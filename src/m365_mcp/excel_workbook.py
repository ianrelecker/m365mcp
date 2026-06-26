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
    Batch read/write:          POST /$batch  (<=20 sub-requests/call)
    Force recalculation:       POST .../application/calculate
    Defined names:             GET  .../names  |  .../worksheets('N')/names
    Resolve a name's range:    GET  .../names('Name')/range
    Clear a range:             POST .../worksheets('N')/range(address)/clear
    Copy into a range:         POST .../worksheets('N')/range(address)/copyFrom
    Insert blank cells:        POST .../worksheets('N')/range(address)/insert
    Create a table:            POST .../worksheets('N')/tables/add
    Sort a table:              POST .../tables/{idOrName}/sort/apply
    Filter a table column:     POST .../tables/{idOrName}/columns('Col')/filter/apply
    Clear table filters:       POST .../tables/{idOrName}/clearFilters
    Format a range:            PATCH .../range(address)/format{,/font,/fill,/borders('Edge')}
    Autofit columns/rows:      POST .../range(address)/format/autofit{Columns,Rows}
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
    formulas: list[list[Any]] | None = None
    numberFormat: list[list[Any]] | None = None
    rowCount: int | None = None
    columnCount: int | None = None


class WorkbookRangeData(BaseModel):
    """One range inside a batch read/write result. ``error`` is set (and the
    data fields left empty) when that individual request failed; the rest of
    the batch is unaffected."""

    worksheet: str
    address: str
    values: list[list[Any]] | None = None
    text: list[list[Any]] | None = None
    formulas: list[list[Any]] | None = None
    numberFormat: list[list[Any]] | None = None
    updated: bool | None = None
    error: str | None = None


class WorkbookRangesResult(BaseModel):
    item: WorkbookItemRef
    ranges: list[WorkbookRangeData] = Field(default_factory=list)


class WorkbookCalculateResult(BaseModel):
    item: WorkbookItemRef
    calculationType: str
    calculated: bool = True


class WorkbookDefinedName(BaseModel):
    name: str
    value: Any | None = None  # the "refers to" formula, e.g. ='Sheet1'!$A$1
    comment: str | None = None
    scope: str | None = None  # "Workbook" or a worksheet name
    type: str | None = None
    visible: bool | None = None


class WorkbookNamesResult(BaseModel):
    item: WorkbookItemRef
    worksheet: str | None = None  # set for worksheet-scoped name lists
    names: list[WorkbookDefinedName] = Field(default_factory=list)


class WorkbookNameRangeResult(BaseModel):
    item: WorkbookItemRef
    name: str
    address: str
    values: list[list[Any]] = Field(default_factory=list)
    text: list[list[Any]] | None = None
    formulas: list[list[Any]] | None = None
    numberFormat: list[list[Any]] | None = None


class WorkbookClearResult(BaseModel):
    item: WorkbookItemRef
    worksheet: str
    address: str
    applyTo: str
    cleared: bool = True


class WorkbookCopyResult(BaseModel):
    item: WorkbookItemRef
    worksheet: str
    address: str
    sourceRange: str
    copyType: str
    copied: bool = True


class WorkbookInsertResult(BaseModel):
    item: WorkbookItemRef
    worksheet: str
    address: str
    shift: str
    inserted: bool = True


class WorkbookDeleteResult(BaseModel):
    item: WorkbookItemRef
    worksheet: str
    address: str
    shift: str
    deleted: bool = True


class WorkbookTableCreateResult(BaseModel):
    item: WorkbookItemRef
    table: WorkbookTable


class WorkbookTableSortResult(BaseModel):
    item: WorkbookItemRef
    table: str
    sorted: bool = True


class WorkbookTableFilterResult(BaseModel):
    item: WorkbookItemRef
    table: str
    column: str
    filtered: bool = True


class WorkbookTableClearFiltersResult(BaseModel):
    item: WorkbookItemRef
    table: str
    cleared: bool = True


class WorkbookFormatResult(BaseModel):
    item: WorkbookItemRef
    worksheet: str
    address: str
    formatted: bool = True


class WorkbookDimensionResult(BaseModel):
    item: WorkbookItemRef
    worksheet: str
    address: str
    autofit: bool = False
    applied: bool = True


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
        """Read a range like 'A1:O5'. Returns raw values, display text, the
        cell formulas, and number formats."""
        data = await self._request(
            f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')"
            f"/range(address='{self._q(address)}')"
            "?$select=address,values,text,formulas,numberFormat,rowCount,columnCount",
            sessionId=sessionId,
        )
        return WorkbookRangeResult(
            item=item,
            worksheet=worksheet,
            address=data.get("address", address),
            values=data.get("values", []),
            text=data.get("text"),
            formulas=data.get("formulas"),
            numberFormat=data.get("numberFormat"),
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
            "?$select=address,values,text,formulas,numberFormat,rowCount,columnCount",
            sessionId=sessionId,
        )
        return WorkbookRangeResult(
            item=item,
            worksheet=worksheet,
            address=data.get("address", ""),
            values=data.get("values", []),
            text=data.get("text"),
            formulas=data.get("formulas"),
            numberFormat=data.get("numberFormat"),
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
        formulas: list[list[Any]] | None = None,
        numberFormat: list[list[Any]] | None = None,
        sessionId: str | None = None,
    ) -> WorkbookWriteResult:
        """Write values, formulas, and/or number formats into a fixed range.
        The shape of ``values``/``formulas``/``numberFormat`` must match the
        address dimensions. ``formulas`` cells may be literal values or formula
        strings (e.g. ``='Unit Mix'!H11``); cross-sheet references are fine."""
        body: dict[str, Any] = {}
        if values is not None:
            body["values"] = values
        if formulas is not None:
            body["formulas"] = formulas
        if numberFormat is not None:
            body["numberFormat"] = numberFormat
        if not body:
            raise ValueError("Provide values, formulas, and/or numberFormat to update")
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

    # ---- batch read / write ---------------------------------------------- #
    async def get_ranges(
        self,
        item: WorkbookItemRef,
        *,
        ranges: list[dict[str, str]],
        sessionId: str | None = None,
    ) -> WorkbookRangesResult:
        """Read many ranges in one shot via Graph ``$batch``.

        ``ranges`` is a list of ``{"worksheet": ..., "address": ...}`` dicts.
        Each result carries values, text, formulas, and numberFormat. Results
        preserve input order; a failed individual range surfaces its ``error``
        without failing the rest of the batch. Auto-chunked to <=20 requests
        per batch call."""
        specs = [self._range_spec(r, i) for i, r in enumerate(ranges)]
        requests = [
            {
                "id": str(i),
                "method": "GET",
                "url": (
                    f"{self._wb_base(item)}/worksheets('{self._q(ws)}')"
                    f"/range(address='{self._q(addr)}')"
                    "?$select=address,values,formulas,text,numberFormat"
                ),
            }
            for i, (ws, addr) in enumerate(specs)
        ]
        responses = await self._batch(requests, sessionId=sessionId)
        out: list[WorkbookRangeData] = []
        for i, (ws, addr) in enumerate(specs):
            resp = responses.get(str(i))
            out.append(self._read_range_response(ws, addr, resp))
        return WorkbookRangesResult(item=item, ranges=out)

    async def update_ranges(
        self,
        item: WorkbookItemRef,
        *,
        updates: list[dict[str, Any]],
        sessionId: str | None = None,
    ) -> WorkbookRangesResult:
        """Write many ranges in one shot via Graph ``$batch`` of PATCH requests.

        ``updates`` is a list of dicts, each with ``worksheet`` and ``address``
        plus any of ``formulas``, ``values``, ``numberFormat``. ``formulas``
        cells may be literal values or formula strings (cross-sheet references
        like ``='Unit Mix'!H11`` are written verbatim). Results preserve input
        order; a failed individual write surfaces its ``error`` without failing
        the rest of the batch. Auto-chunked to <=20 requests per batch call."""
        prepared: list[tuple[str, str, dict[str, Any]]] = []
        for i, upd in enumerate(updates):
            ws = upd.get("worksheet")
            addr = upd.get("address")
            if not ws or not addr:
                raise ValueError(
                    f"updates[{i}] must include worksheet and address"
                )
            body: dict[str, Any] = {}
            for key in ("formulas", "values", "numberFormat"):
                if upd.get(key) is not None:
                    body[key] = upd[key]
            if not body:
                raise ValueError(
                    f"updates[{i}] must include formulas, values, "
                    "and/or numberFormat"
                )
            prepared.append((ws, addr, body))
        requests = [
            {
                "id": str(i),
                "method": "PATCH",
                "url": (
                    f"{self._wb_base(item)}/worksheets('{self._q(ws)}')"
                    f"/range(address='{self._q(addr)}')"
                ),
                "body": body,
            }
            for i, (ws, addr, body) in enumerate(prepared)
        ]
        responses = await self._batch(requests, sessionId=sessionId)
        out: list[WorkbookRangeData] = []
        for i, (ws, addr, _body) in enumerate(prepared):
            resp = responses.get(str(i))
            error = self._batch_error(resp)
            out.append(
                WorkbookRangeData(
                    worksheet=ws,
                    address=addr,
                    updated=error is None,
                    error=error,
                )
            )
        return WorkbookRangesResult(item=item, ranges=out)

    async def calculate(
        self,
        item: WorkbookItemRef,
        *,
        calculationType: str = "Full",
        sessionId: str | None = None,
    ) -> WorkbookCalculateResult:
        """Force a recalculation of the workbook so computed cells are current
        before reading them back. ``calculationType`` is one of ``Recalculate``,
        ``Full``, or ``FullRebuild``."""
        allowed = {"Recalculate", "Full", "FullRebuild"}
        if calculationType not in allowed:
            raise ValueError(
                f"calculationType must be one of {sorted(allowed)}"
            )
        await self._request(
            f"{self._wb_base(item)}/application/calculate",
            method="POST",
            json_body={"calculationType": calculationType},
            sessionId=sessionId,
        )
        return WorkbookCalculateResult(
            item=item, calculationType=calculationType
        )

    # ---- defined names ---------------------------------------------------- #
    async def list_names(
        self,
        item: WorkbookItemRef,
        *,
        worksheet: str | None = None,
        sessionId: str | None = None,
    ) -> WorkbookNamesResult:
        """List defined names. Workbook-scoped names when ``worksheet`` is
        omitted; worksheet-scoped names when it is given."""
        if worksheet:
            path = (
                f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')/names"
            )
        else:
            path = f"{self._wb_base(item)}/names"
        data = await self._request(
            f"{path}?$select=name,value,comment,scope,type,visible",
            sessionId=sessionId,
        )
        return WorkbookNamesResult(
            item=item,
            worksheet=worksheet,
            names=[
                WorkbookDefinedName(
                    name=n["name"],
                    value=n.get("value"),
                    comment=n.get("comment"),
                    scope=n.get("scope"),
                    type=n.get("type"),
                    visible=n.get("visible"),
                )
                for n in data.get("value", [])
            ],
        )

    async def get_name_range(
        self,
        item: WorkbookItemRef,
        *,
        name: str,
        worksheet: str | None = None,
        sessionId: str | None = None,
    ) -> WorkbookNameRangeResult:
        """Resolve a defined name to its range and read it. Provide
        ``worksheet`` for a worksheet-scoped name; omit it for a workbook-scoped
        one."""
        if worksheet:
            base = (
                f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')"
                f"/names('{self._q(name)}')/range"
            )
        else:
            base = f"{self._wb_base(item)}/names('{self._q(name)}')/range"
        data = await self._request(
            f"{base}?$select=address,values,text,formulas,numberFormat",
            sessionId=sessionId,
        )
        return WorkbookNameRangeResult(
            item=item,
            name=name,
            address=data.get("address", ""),
            values=data.get("values", []),
            text=data.get("text"),
            formulas=data.get("formulas"),
            numberFormat=data.get("numberFormat"),
        )

    # ---- range operations ------------------------------------------------- #
    async def clear_range(
        self,
        item: WorkbookItemRef,
        *,
        worksheet: str,
        address: str,
        applyTo: str = "Contents",
        sessionId: str | None = None,
    ) -> WorkbookClearResult:
        """Clear a range. ``applyTo`` is ``Contents`` (values/formulas only),
        ``Formats``, or ``All``."""
        allowed = {"Contents", "Formats", "All"}
        if applyTo not in allowed:
            raise ValueError(f"applyTo must be one of {sorted(allowed)}")
        await self._request(
            f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')"
            f"/range(address='{self._q(address)}')/clear",
            method="POST",
            json_body={"applyTo": applyTo},
            sessionId=sessionId,
        )
        return WorkbookClearResult(
            item=item, worksheet=worksheet, address=address, applyTo=applyTo
        )

    async def copy_range(
        self,
        item: WorkbookItemRef,
        *,
        worksheet: str,
        address: str,
        sourceRange: str,
        copyType: str = "All",
        sessionId: str | None = None,
    ) -> WorkbookCopyResult:
        """Copy into ``address`` (the destination) from ``sourceRange`` (e.g.
        ``'Unit Mix'!A1:B5`` for a cross-sheet source). ``copyType`` is one of
        ``All``, ``Formulas``, ``Values``, or ``Formats``."""
        allowed = {"All", "Formulas", "Values", "Formats"}
        if copyType not in allowed:
            raise ValueError(f"copyType must be one of {sorted(allowed)}")
        await self._request(
            f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')"
            f"/range(address='{self._q(address)}')/copyFrom",
            method="POST",
            json_body={"sourceRange": sourceRange, "copyType": copyType},
            sessionId=sessionId,
        )
        return WorkbookCopyResult(
            item=item,
            worksheet=worksheet,
            address=address,
            sourceRange=sourceRange,
            copyType=copyType,
        )

    async def insert_range(
        self,
        item: WorkbookItemRef,
        *,
        worksheet: str,
        address: str,
        shift: str = "Down",
        sessionId: str | None = None,
    ) -> WorkbookInsertResult:
        """Insert blank cells at ``address``, shifting existing cells. ``shift``
        is ``Down`` or ``Right``."""
        allowed = {"Down", "Right"}
        if shift not in allowed:
            raise ValueError(f"shift must be one of {sorted(allowed)}")
        await self._request(
            f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')"
            f"/range(address='{self._q(address)}')/insert",
            method="POST",
            json_body={"shift": shift},
            sessionId=sessionId,
        )
        return WorkbookInsertResult(
            item=item, worksheet=worksheet, address=address, shift=shift
        )

    async def delete_range(
        self,
        item: WorkbookItemRef,
        *,
        worksheet: str,
        address: str,
        shift: str = "Up",
        sessionId: str | None = None,
    ) -> WorkbookDeleteResult:
        """Delete the cells at ``address`` and shift remaining cells to fill the
        gap. ``shift`` is ``Up`` (default) or ``Left``. Use a full-row address
        (e.g. ``5:5`` or ``A5:Z5``) with ``shift='Up'`` to delete a row. This
        edits cells inside the worksheet via Excel's engine; it never deletes
        the workbook file. To merely blank cells in place without shifting,
        use ``clear_range`` instead."""
        allowed = {"Up", "Left"}
        if shift not in allowed:
            raise ValueError(f"shift must be one of {sorted(allowed)}")
        await self._request(
            f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')"
            f"/range(address='{self._q(address)}')/delete",
            method="POST",
            json_body={"shift": shift},
            sessionId=sessionId,
        )
        return WorkbookDeleteResult(
            item=item, worksheet=worksheet, address=address, shift=shift
        )

    # ---- tables (structure) ---------------------------------------------- #
    async def create_table(
        self,
        item: WorkbookItemRef,
        *,
        worksheet: str,
        address: str,
        hasHeaders: bool = True,
        sessionId: str | None = None,
    ) -> WorkbookTableCreateResult:
        """Create an Excel table over an existing range (a.k.a. convert a range
        to a table). ``address`` is a range on ``worksheet`` like ``A1:D20``;
        ``hasHeaders=True`` treats the first row as column headers. Returns the
        created table's id and name for use with the other table tools."""
        data = await self._request(
            f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')/tables/add",
            method="POST",
            json_body={"address": address, "hasHeaders": hasHeaders},
            sessionId=sessionId,
        )
        return WorkbookTableCreateResult(
            item=item,
            table=WorkbookTable(
                id=data["id"],
                name=data["name"],
                showHeaders=data.get("showHeaders"),
                worksheet=worksheet,
            ),
        )

    async def sort_table(
        self,
        item: WorkbookItemRef,
        *,
        table: str,
        fields: list[dict[str, Any]],
        matchCase: bool = False,
        sessionId: str | None = None,
    ) -> WorkbookTableSortResult:
        """Sort a table by one or more columns. ``fields`` is a list of sort
        conditions, each ``{"key": <zero-based column index in the table>,
        "ascending": true|false}`` (an optional ``"sortOn"`` may be ``Value``,
        ``CellColor``, ``FontColor``, or ``Icon``). Earlier fields take
        priority. ``matchCase`` makes text comparison case-sensitive."""
        if not fields:
            raise ValueError(
                "Provide at least one sort field, e.g. {'key': 0, 'ascending': True}."
            )
        await self._request(
            f"{self._wb_base(item)}/tables/{self._q(table)}/sort/apply",
            method="POST",
            json_body={"fields": fields, "matchCase": matchCase},
            sessionId=sessionId,
        )
        return WorkbookTableSortResult(item=item, table=table)

    async def filter_table(
        self,
        item: WorkbookItemRef,
        *,
        table: str,
        column: str,
        criteria: dict[str, Any],
        sessionId: str | None = None,
    ) -> WorkbookTableFilterResult:
        """Apply a filter to one table column. ``column`` is the column's header
        name. ``criteria`` is a Graph filter-criteria object, e.g.
        ``{"filterOn": "values", "values": ["A", "B"]}`` to keep matching rows,
        or ``{"filterOn": "custom", "criterion1": ">100", "operator": "And"}``.
        Use ``clear_table_filters`` to remove filters afterwards."""
        await self._request(
            f"{self._wb_base(item)}/tables/{self._q(table)}"
            f"/columns('{self._q(column)}')/filter/apply",
            method="POST",
            json_body={"criteria": criteria},
            sessionId=sessionId,
        )
        return WorkbookTableFilterResult(item=item, table=table, column=column)

    async def clear_table_filters(
        self,
        item: WorkbookItemRef,
        *,
        table: str,
        sessionId: str | None = None,
    ) -> WorkbookTableClearFiltersResult:
        """Clear all column filters on a table, restoring every row to view."""
        await self._request(
            f"{self._wb_base(item)}/tables/{self._q(table)}/clearFilters",
            method="POST",
            sessionId=sessionId,
        )
        return WorkbookTableClearFiltersResult(item=item, table=table)

    # ---- formatting ------------------------------------------------------ #
    async def format_range(
        self,
        item: WorkbookItemRef,
        *,
        worksheet: str,
        address: str,
        fillColor: str | None = None,
        fontBold: bool | None = None,
        fontItalic: bool | None = None,
        fontUnderline: str | None = None,
        fontColor: str | None = None,
        fontSize: float | None = None,
        fontName: str | None = None,
        horizontalAlignment: str | None = None,
        verticalAlignment: str | None = None,
        wrapText: bool | None = None,
        borderStyle: str | None = None,
        borderColor: str | None = None,
        borderEdges: list[str] | None = None,
        sessionId: str | None = None,
    ) -> WorkbookFormatResult:
        """Apply visual formatting to a range. Colors are hex strings like
        ``#FFFF00``. ``fontUnderline`` is ``None``/``Single``/``Double``.
        ``horizontalAlignment`` is ``General``/``Left``/``Center``/``Right`` (etc.)
        and ``verticalAlignment`` is ``Top``/``Center``/``Bottom``. Setting
        ``borderStyle`` (e.g. ``Continuous``) draws borders on ``borderEdges``
        (default the four outer edges: ``EdgeTop``/``EdgeBottom``/``EdgeLeft``/
        ``EdgeRight``; inside edges are ``InsideVertical``/``InsideHorizontal``).
        Number formats are set via ``update_range``. Issues one Graph PATCH per
        sub-property changed; at least one property must be provided."""
        base = (
            f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')"
            f"/range(address='{self._q(address)}')/format"
        )
        did_any = False

        font_body: dict[str, Any] = {}
        if fontBold is not None:
            font_body["bold"] = fontBold
        if fontItalic is not None:
            font_body["italic"] = fontItalic
        if fontUnderline is not None:
            font_body["underline"] = fontUnderline
        if fontColor is not None:
            font_body["color"] = fontColor
        if fontSize is not None:
            font_body["size"] = fontSize
        if fontName is not None:
            font_body["name"] = fontName
        if font_body:
            await self._request(
                f"{base}/font", method="PATCH", json_body=font_body,
                sessionId=sessionId,
            )
            did_any = True

        if fillColor is not None:
            await self._request(
                f"{base}/fill", method="PATCH",
                json_body={"color": fillColor}, sessionId=sessionId,
            )
            did_any = True

        fmt_body: dict[str, Any] = {}
        if horizontalAlignment is not None:
            fmt_body["horizontalAlignment"] = horizontalAlignment
        if verticalAlignment is not None:
            fmt_body["verticalAlignment"] = verticalAlignment
        if wrapText is not None:
            fmt_body["wrapText"] = wrapText
        if fmt_body:
            await self._request(
                base, method="PATCH", json_body=fmt_body, sessionId=sessionId
            )
            did_any = True

        if borderStyle is not None:
            edges = borderEdges or [
                "EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"
            ]
            border_body: dict[str, Any] = {"style": borderStyle}
            if borderColor is not None:
                border_body["color"] = borderColor
            for edge in edges:
                await self._request(
                    f"{base}/borders('{self._q(edge)}')",
                    method="PATCH", json_body=border_body, sessionId=sessionId,
                )
            did_any = True

        if not did_any:
            raise ValueError(
                "Provide at least one formatting property to apply."
            )
        return WorkbookFormatResult(
            item=item, worksheet=worksheet, address=address
        )

    async def set_column_width(
        self,
        item: WorkbookItemRef,
        *,
        worksheet: str,
        columns: str,
        width: float | None = None,
        autofit: bool = False,
        sessionId: str | None = None,
    ) -> WorkbookDimensionResult:
        """Set the width of one or more columns, addressed like ``A:A`` or
        ``A:C``. Pass ``width`` (in points) to set a fixed width, or
        ``autofit=True`` to size columns to their content."""
        base = (
            f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')"
            f"/range(address='{self._q(columns)}')/format"
        )
        if autofit:
            await self._request(
                f"{base}/autofitColumns", method="POST", sessionId=sessionId
            )
        elif width is not None:
            await self._request(
                base, method="PATCH",
                json_body={"columnWidth": width}, sessionId=sessionId,
            )
        else:
            raise ValueError("Provide width, or set autofit=True.")
        return WorkbookDimensionResult(
            item=item, worksheet=worksheet, address=columns, autofit=autofit
        )

    async def set_row_height(
        self,
        item: WorkbookItemRef,
        *,
        worksheet: str,
        rows: str,
        height: float | None = None,
        autofit: bool = False,
        sessionId: str | None = None,
    ) -> WorkbookDimensionResult:
        """Set the height of one or more rows, addressed like ``1:1`` or
        ``1:10``. Pass ``height`` (in points) to set a fixed height, or
        ``autofit=True`` to size rows to their content."""
        base = (
            f"{self._wb_base(item)}/worksheets('{self._q(worksheet)}')"
            f"/range(address='{self._q(rows)}')/format"
        )
        if autofit:
            await self._request(
                f"{base}/autofitRows", method="POST", sessionId=sessionId
            )
        elif height is not None:
            await self._request(
                base, method="PATCH",
                json_body={"rowHeight": height}, sessionId=sessionId,
            )
        else:
            raise ValueError("Provide height, or set autofit=True.")
        return WorkbookDimensionResult(
            item=item, worksheet=worksheet, address=rows, autofit=autofit
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

    @staticmethod
    def _range_spec(spec: dict[str, str], index: int) -> tuple[str, str]:
        ws = spec.get("worksheet")
        addr = spec.get("address")
        if not ws or not addr:
            raise ValueError(
                f"ranges[{index}] must include worksheet and address"
            )
        return ws, addr

    @staticmethod
    def _read_range_response(
        worksheet: str, address: str, resp: dict[str, Any] | None
    ) -> WorkbookRangeData:
        error = ExcelWorkbookClient._batch_error(resp)
        if error is not None or resp is None:
            return WorkbookRangeData(
                worksheet=worksheet, address=address, error=error
            )
        body = resp.get("body") or {}
        return WorkbookRangeData(
            worksheet=worksheet,
            address=body.get("address", address),
            values=body.get("values"),
            text=body.get("text"),
            formulas=body.get("formulas"),
            numberFormat=body.get("numberFormat"),
        )

    @staticmethod
    def _batch_error(resp: dict[str, Any] | None) -> str | None:
        """Return an error string for a failed inner $batch response, else
        None. A missing response (request id dropped from the batch) is an
        error too."""
        if resp is None:
            return "no response returned for this request"
        status = resp.get("status")
        if isinstance(status, int) and 200 <= status < 300:
            return None
        detail = ExcelWorkbookClient._error_detail(resp.get("body"))
        return f"{status}: {detail}" if detail else f"HTTP {status}"

    @staticmethod
    def _chunked(items: list[Any], size: int) -> list[list[Any]]:
        return [items[i : i + size] for i in range(0, len(items), size)]

    async def _batch(
        self,
        requests: list[dict[str, Any]],
        *,
        sessionId: str | None = None,
    ) -> dict[str, dict[str, Any]]:
        """Send Graph ``POST /$batch`` requests, auto-chunked to <=20 per call.

        Each entry in ``requests`` is ``{"id", "method", "url", "body"?,
        "headers"?}`` with ``url`` relative to ``/v1.0``. The workbook session
        header is attached to every inner request when ``sessionId`` is given.
        Returns a dict keyed by request id -> raw inner response, aggregated
        across chunks so the caller can reassemble in input order."""
        responses: dict[str, dict[str, Any]] = {}
        for chunk in self._chunked(requests, 20):
            payload: dict[str, Any] = {"requests": []}
            for req in chunk:
                entry: dict[str, Any] = {
                    "id": req["id"],
                    "method": req["method"],
                    "url": req["url"],
                }
                headers = dict(req.get("headers") or {})
                if req.get("body") is not None:
                    entry["body"] = req["body"]
                    headers.setdefault("Content-Type", "application/json")
                if sessionId:
                    headers["workbook-session-id"] = sessionId
                if headers:
                    entry["headers"] = headers
                payload["requests"].append(entry)
            data = await self._request("/$batch", method="POST", json_body=payload)
            for resp in (data or {}).get("responses", []):
                responses[str(resp.get("id"))] = resp
        return responses

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
