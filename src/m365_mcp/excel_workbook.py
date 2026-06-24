"""Microsoft Graph Excel (Workbook) API client for the M365 MCP server.

This module adds *in-place* spreadsheet editing on files stored in OneDrive /
SharePoint, using the Graph Workbook API. Unlike download -> openpyxl -> upload,
the workbook is edited server-side by Excel's own engine, so:

  * formatting, formulas, and data-validation dropdowns are preserved,
  * the binary is never moved (no truncation / OneDrive sync races),
  * every change is versioned by SharePoint and recoverable.

Design mirrors microsoft_graph.MicrosoftGraphClient: it takes the same
MicrosoftAuthService and an optional shared httpx.AsyncClient, and exposes a
small private _request helper with identical error handling.

Required delegated scope:
    Files.ReadWrite.All       # read+write files the user can access, incl.
                              # workbooks in SharePoint document libraries

This single scope authorizes the in-place edits: workbook calls go through
/drives/{id}/items/{id}/workbook, not /sites, so no SharePoint *write*
(Sites.ReadWrite.All) scope is needed. The companion browse client
(sharepoint_files.py) additionally uses the read-only Sites.Read.All for
site/drive discovery.

Endpoint reference (Graph v1.0):
    Resolve a shared file:     GET  /shares/{encoded}/driveItem
    Resolve by path:           GET  /drives/{driveId}/root:/{path}:
    Workbook base:             /drives/{driveId}/items/{itemId}/workbook
    List worksheets:           GET  .../worksheets
    List tables:               GET  .../tables   (or .../worksheets('N')/tables)
    Append a table row:        POST .../tables/{idOrName}/rows/add
    Read a range:              GET  .../worksheets('N')/range(address='A1:O5')
    Write a range:             PATCH .../worksheets('N')/range(address='A2:O2')
    Used range:                GET  .../worksheets('N')/usedRange
    Create a session:          POST .../createSession  {"persistChanges": true}
"""

from __future__ import annotations

import base64
from contextlib import asynccontextmanager
from datetime import date, datetime
from typing import Any
from urllib.parse import quote

import httpx
from pydantic import BaseModel, Field

from .microsoft_auth import MicrosoftAuthService

GRAPH_V1 = "https://graph.microsoft.com/v1.0"

# Excel's day 0 is 1899-12-30 (the well-known Lotus 1-2-3 leap-year bug offset).
_EXCEL_EPOCH = date(1899, 12, 30)


# --------------------------------------------------------------------------- #
# Result models (typed, in the style of models.py)
# --------------------------------------------------------------------------- #
class WorkbookItemRef(BaseModel):
    """Identifies a workbook on a drive. Reuse driveId+itemId for later calls
    to skip re-resolution."""

    driveId: str
    itemId: str
    name: str | None = None
    webUrl: str | None = None


class WorkbookWorksheet(BaseModel):
    id: str
    name: str
    position: int | None = None
    visibility: str | None = None


class WorkbookTable(BaseModel):
    id: str
    name: str
    showHeaders: bool | None = None
    worksheet: str | None = None


class WorkbookListWorksheetsResult(BaseModel):
    item: WorkbookItemRef
    worksheets: list[WorkbookWorksheet]


class WorkbookListTablesResult(BaseModel):
    item: WorkbookItemRef
    tables: list[WorkbookTable]


class WorkbookRangeResult(BaseModel):
    item: WorkbookItemRef
    worksheet: str
    address: str
    values: list[list[Any]] = Field(default_factory=list)
    text: list[list[Any]] | None = None
    rowCount: int | None = None
    columnCount: int | None = None


class WorkbookRowAddResult(BaseModel):
    item: WorkbookItemRef
    table: str
    index: int | None = None
    values: list[list[Any]] = Field(default_factory=list)


class WorkbookSessionResult(BaseModel):
    item: WorkbookItemRef
    sessionId: str
    persistChanges: bool


class WorkbookWriteResult(BaseModel):
    item: WorkbookItemRef
    worksheet: str
    address: str
    updated: bool = True


# --------------------------------------------------------------------------- #
# Client
# --------------------------------------------------------------------------- #
class ExcelWorkbookClient:
    def __init__(
        self,
        auth_service: MicrosoftAuthService,
        http_client: httpx.AsyncClient | None = None,
    ) -> None:
        self._auth_service = auth_service
        self._http_client = http_client

    # ---- public API ------------------------------------------------------- #
    async def resolve_workbook(
        self,
        *,
        shareUrl: str | None = None,
        driveId: str | None = None,
        itemId: str | None = None,
        itemPath: str | None = None,
    ) -> WorkbookItemRef:
        """Resolve a workbook to a (driveId, itemId) pair.

        Provide ONE of:
          * shareUrl  - a SharePoint/OneDrive sharing or browser URL to the file
          * driveId + itemId
          * driveId + itemPath  (path relative to the drive root, e.g.
            "Shared Active Deals/4. Claude Projects/Acquisitions Deals Tracker (UPDATED).xlsx")
        """
        if shareUrl:
            encoded = self._encode_share_url(shareUrl)
            data = await self._request(
                f"/shares/{encoded}/driveItem"
                "?$select=id,name,webUrl,parentReference"
            )
            parent = data.get("parentReference") or {}
            return WorkbookItemRef(
                driveId=str(parent.get("driveId")),
                itemId=str(data["id"]),
                name=data.get("name"),
                webUrl=data.get("webUrl"),
            )
        if driveId and itemId:
            data = await self._request(
                f"/drives/{quote(driveId, safe='')}/items/{quote(itemId, safe='')}"
                "?$select=id,name,webUrl"
            )
            return WorkbookItemRef(
                driveId=driveId,
                itemId=str(data["id"]),
                name=data.get("name"),
                webUrl=data.get("webUrl"),
            )
        if driveId and itemPath:
            path = itemPath.strip("/")
            data = await self._request(
                f"/drives/{quote(driveId, safe='')}/root:/{quote(path)}:"
                "?$select=id,name,webUrl"
            )
            return WorkbookItemRef(
                driveId=driveId,
                itemId=str(data["id"]),
                name=data.get("name"),
                webUrl=data.get("webUrl"),
            )
        raise ValueError(
            "Provide shareUrl, or driveId+itemId, or driveId+itemPath"
        )

    async def list_worksheets(
        self, item: WorkbookItemRef, *, sessionId: str | None = None
    ) -> WorkbookListWorksheetsResult:
        data = await self._request(
            f"{self._wb_base(item)}/worksheets?$select=id,name,position,visibility",
            sessionId=sessionId,
        )
        return WorkbookListWorksheetsResult(
            item=item,
            worksheets=[
                WorkbookWorksheet(
                    id=ws["id"],
                    name=ws["name"],
                    position=ws.get("position"),
                    visibility=ws.get("visibility"),
                )
                for ws in data.get("value", [])
            ],
        )

    async def list_tables(
        self,
        item: WorkbookItemRef,
        *,
        worksheet: str | None = None,
        sessionId: str | None = None,
    ) -> WorkbookListTablesResult:
        if worksheet:
            path = (
                f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')/tables"
            )
        else:
            path = f"{self._wb_base(item)}/tables"
        data = await self._request(
            f"{path}?$select=id,name,showHeaders", sessionId=sessionId
        )
        return WorkbookListTablesResult(
            item=item,
            tables=[
                WorkbookTable(
                    id=t["id"],
                    name=t["name"],
                    showHeaders=t.get("showHeaders"),
                    worksheet=worksheet,
                )
                for t in data.get("value", [])
            ],
        )

    async def add_table_row(
        self,
        item: WorkbookItemRef,
        *,
        table: str,
        values: list[list[Any]],
        index: int | None = None,
        sessionId: str | None = None,
    ) -> WorkbookRowAddResult:
        """Append one or more rows to an Excel table. ``values`` is a 2D array;
        each inner list is one row and must match the table's column count and
        order. ``index=None`` appends at the end; ``index=0`` inserts at top."""
        body: dict[str, Any] = {"values": values}
        if index is not None:
            body["index"] = index
        data = await self._request(
            f"{self._wb_base(item)}/tables/{self._q(table)}/rows/add",
            method="POST",
            json_body=body,
            sessionId=sessionId,
        )
        return WorkbookRowAddResult(
            item=item,
            table=table,
            index=data.get("index"),
            values=data.get("values", []),
        )

    async def get_range(
        self,
        item: WorkbookItemRef,
        *,
        worksheet: str,
        address: str,
        sessionId: str | None = None,
    ) -> WorkbookRangeResult:
        """Read a range like 'A1:O5'. Returns both raw values and display text."""
        data = await self._request(
            f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')"
            f"/range(address='{self._q(address)}')"
            "?$select=address,values,text,rowCount,columnCount",
            sessionId=sessionId,
        )
        return WorkbookRangeResult(
            item=item,
            worksheet=worksheet,
            address=data.get("address", address),
            values=data.get("values", []),
            text=data.get("text"),
            rowCount=data.get("rowCount"),
            columnCount=data.get("columnCount"),
        )

    async def get_used_range(
        self,
        item: WorkbookItemRef,
        *,
        worksheet: str,
        valuesOnly: bool = True,
        sessionId: str | None = None,
    ) -> WorkbookRangeResult:
        suffix = "(valuesOnly=true)" if valuesOnly else ""
        data = await self._request(
            f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')"
            f"/usedRange{suffix}"
            "?$select=address,values,text,rowCount,columnCount",
            sessionId=sessionId,
        )
        return WorkbookRangeResult(
            item=item,
            worksheet=worksheet,
            address=data.get("address", ""),
            values=data.get("values", []),
            text=data.get("text"),
            rowCount=data.get("rowCount"),
            columnCount=data.get("columnCount"),
        )

    async def update_range(
        self,
        item: WorkbookItemRef,
        *,
        worksheet: str,
        address: str,
        values: list[list[Any]] | None = None,
        numberFormat: list[list[Any]] | None = None,
        sessionId: str | None = None,
    ) -> WorkbookWriteResult:
        """Write values and/or number formats into a fixed range. The shape of
        ``values``/``numberFormat`` must match the address dimensions."""
        body: dict[str, Any] = {}
        if values is not None:
            body["values"] = values
        if numberFormat is not None:
            body["numberFormat"] = numberFormat
        if not body:
            raise ValueError("Provide values and/or numberFormat to update")
        await self._request(
            f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')"
            f"/range(address='{self._q(address)}')",
            method="PATCH",
            json_body=body,
            sessionId=sessionId,
        )
        return WorkbookWriteResult(item=item, worksheet=worksheet, address=address)

    async def create_session(
        self,
        item: WorkbookItemRef,
        *,
        persistChanges: bool = True,
    ) -> WorkbookSessionResult:
        """Create a workbook session. Pass the returned sessionId to subsequent
        calls to batch them consistently. persistChanges=True writes to the
        stored file; False is a scratch/read session."""
        data = await self._request(
            f"{self._wb_base(item)}/createSession",
            method="POST",
            json_body={"persistChanges": persistChanges},
        )
        return WorkbookSessionResult(
            item=item,
            sessionId=str(data["id"]),
            persistChanges=persistChanges,
        )

    async def close_session(
        self, item: WorkbookItemRef, *, sessionId: str
    ) -> None:
        await self._request(
            f"{self._wb_base(item)}/closeSession",
            method="POST",
            sessionId=sessionId,
        )

    # ---- helpers ---------------------------------------------------------- #
    @staticmethod
    def excel_serial_date(value: date | datetime | str) -> int:
        """Convert a date to an Excel serial number for date cells. Use the
        result as a cell value together with a date numberFormat (e.g.
        'mm/dd/yy') so Excel stores a real date, not text."""
        if isinstance(value, str):
            value = datetime.fromisoformat(value)
        if isinstance(value, datetime):
            value = value.date()
        return (value - _EXCEL_EPOCH).days

    @staticmethod
    def _encode_share_url(url: str) -> str:
        """Encode a sharing URL per Graph's /shares addressing:
        'u!' + base64url(url) with padding stripped."""
        b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii")
        return "u!" + b64.rstrip("=")

    @staticmethod
    def _q(value: str) -> str:
        """Escape a single quote for use inside an OData string literal like
        worksheets('Name') or range(address='A1')."""
        return value.replace("'", "''")

    @staticmethod
    def _wb_base(item: WorkbookItemRef) -> str:
        return (
            f"/drives/{quote(item.driveId, safe='')}"
            f"/items/{quote(item.itemId, safe='')}/workbook"
        )

    @asynccontextmanager
    async def _client(self) -> Any:
        if self._http_client is not None:
            yield self._http_client
            return
        async with httpx.AsyncClient(follow_redirects=False, timeout=60.0) as client:
            yield client

    async def _request(
        self,
        path: str,
        *,
        method: str = "GET",
        json_body: dict[str, Any] | None = None,
        sessionId: str | None = None,
    ) -> Any:
        access_token = await self._auth_service.get_access_token()
        headers = {
            "Accept": "application/json",
            "Authorization": f"Bearer {access_token}",
            **({"Content-Type": "application/json"} if json_body is not None else {}),
            **({"workbook-session-id": sessionId} if sessionId else {}),
        }
        url = path if path.startswith("https://") else f"{GRAPH_V1}{path}"
        async with self._client() as client:
            response = await client.request(
                method, url, headers=headers, json=json_body
            )
        if response.status_code == 204:
            return None
        data = response.json() if response.text else None
        if not response.is_success:
            detail = self._error_detail(data) or response.reason_phrase
            raise RuntimeError(
                f"Microsoft Graph Workbook request failed "
                f"({response.status_code}): {detail}"
            )
        return data

    @staticmethod
    def _error_detail(data: Any) -> str | None:
        if not isinstance(data, dict):
            return None
        err = data.get("error")
        if isinstance(err, dict):
            code = err.get("code")
            message = err.get("message")
            return f"{code}: {message}" if code else message
        return data.get("error_description")
