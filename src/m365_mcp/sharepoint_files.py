"""Read-only SharePoint / OneDrive browsing for the M365 MCP server.

Lets Claude find folders anywhere the signed-in user has access and list their
contents (Excel, PDF, etc.) — without mounting the folder locally. Pairs with
excel_workbook.py: use this to *locate* a workbook, then hand the returned
driveId+itemId to the Workbook tools to edit it in place.

Same design as the other clients: shared MicrosoftAuthService + httpx client,
private _request with identical error handling, typed pydantic results.

Scopes (delegated):
    Sites.Read.All   +  Files.ReadWrite.All
  This client is read-only and would work with Files.Read.All, but the
  companion Workbook client (excel_workbook.py) needs Files.ReadWrite.All to
  edit, so the shared config requests the write file scope and the read-only
  Sites.Read.All. No SharePoint *write* (Sites.ReadWrite.All) scope is used.

Endpoint reference (Graph v1.0):
    Search sites:        GET  /sites?search={q}
    Site by path:        GET  /sites/{hostname}:/sites/{path}
    Site's libraries:    GET  /sites/{siteId}/drives
    Folder children:     GET  /drives/{driveId}/root/children
                         GET  /drives/{driveId}/items/{itemId}/children
                         GET  /drives/{driveId}/root:/{path}:/children
    Search in a drive:   GET  /drives/{driveId}/root/search(q='{q}')
    Search everywhere:   POST /search/query   (entityTypes: driveItem)
    Resolve a link:      GET  /shares/{u!encoded}/driveItem
"""

from __future__ import annotations

import base64
from contextlib import asynccontextmanager
from typing import Any
from urllib.parse import quote

import httpx
from pydantic import BaseModel, Field

from .microsoft_auth import MicrosoftAuthService

GRAPH_V1 = "https://graph.microsoft.com/v1.0"


# --------------------------------------------------------------------------- #
# Models
# --------------------------------------------------------------------------- #
class SiteInfo(BaseModel):
    id: str
    name: str | None = None
    displayName: str | None = None
    webUrl: str | None = None


class DriveInfo(BaseModel):
    id: str
    name: str | None = None
    driveType: str | None = None
    webUrl: str | None = None


class DriveItemInfo(BaseModel):
    name: str
    itemId: str
    driveId: str | None = None
    isFolder: bool = False
    childCount: int | None = None
    size: int | None = None
    fileExtension: str | None = None
    webUrl: str | None = None
    lastModifiedDateTime: str | None = None
    path: str | None = None  # parentReference.path, e.g. /drive/root:/Folder/Sub


class SitesResult(BaseModel):
    query: str | None = None
    sites: list[SiteInfo] = Field(default_factory=list)


class DrivesResult(BaseModel):
    siteId: str
    drives: list[DriveInfo] = Field(default_factory=list)


class DriveItemsResult(BaseModel):
    driveId: str | None = None
    parentItemId: str | None = None
    path: str | None = None
    items: list[DriveItemInfo] = Field(default_factory=list)
    nextLink: str | None = None


# --------------------------------------------------------------------------- #
# Client
# --------------------------------------------------------------------------- #
class SharePointFilesClient:
    def __init__(
        self,
        auth_service: MicrosoftAuthService,
        http_client: httpx.AsyncClient | None = None,
    ) -> None:
        self._auth_service = auth_service
        self._http_client = http_client

    # ---- discovery -------------------------------------------------------- #
    async def search_sites(self, *, query: str, top: int = 25) -> SitesResult:
        """Find SharePoint sites by keyword (name/title)."""
        data = await self._request(
            f"/sites?search={quote(query)}&$top={min(top, 100)}"
            "&$select=id,name,displayName,webUrl"
        )
        return SitesResult(
            query=query,
            sites=[self._map_site(s) for s in data.get("value", [])],
        )

    async def get_site_by_path(
        self, *, hostname: str, sitePath: str
    ) -> SiteInfo:
        """Resolve a known site, e.g. hostname='kcbm.sharepoint.com',
        sitePath='Acquisitions' -> /sites/{host}:/sites/Acquisitions."""
        path = sitePath.strip("/")
        data = await self._request(
            f"/sites/{quote(hostname, safe='')}:/sites/{quote(path)}"
            "?$select=id,name,displayName,webUrl"
        )
        return self._map_site(data)

    async def list_drives(self, *, siteId: str) -> DrivesResult:
        """List a site's document libraries (each library is a 'drive')."""
        data = await self._request(
            f"/sites/{quote(siteId, safe='')}/drives"
            "?$select=id,name,driveType,webUrl"
        )
        return DrivesResult(
            siteId=siteId,
            drives=[
                DriveInfo(
                    id=d["id"],
                    name=d.get("name"),
                    driveType=d.get("driveType"),
                    webUrl=d.get("webUrl"),
                )
                for d in data.get("value", [])
            ],
        )

    # ---- browse ----------------------------------------------------------- #
    async def list_children(
        self,
        *,
        driveId: str,
        itemId: str | None = None,
        path: str | None = None,
        top: int = 200,
        extensions: list[str] | None = None,
        foldersOnly: bool = False,
    ) -> DriveItemsResult:
        """List the contents of a folder.

        Target the folder by itemId, or by path relative to the drive root
        (e.g. 'Shared Active Deals/4. Claude Projects'); omit both for the root.
        Optionally filter to file extensions (e.g. ['xlsx','pdf']) or folders.
        """
        if itemId:
            base = f"/drives/{quote(driveId, safe='')}/items/{quote(itemId, safe='')}/children"
        elif path:
            clean = path.strip("/")
            base = f"/drives/{quote(driveId, safe='')}/root:/{quote(clean)}:/children"
        else:
            base = f"/drives/{quote(driveId, safe='')}/root/children"
        url = (
            f"{base}?$top={min(top, 999)}"
            "&$select=id,name,folder,file,size,webUrl,lastModifiedDateTime,parentReference"
        )
        data = await self._request(url)
        items = [self._map_item(i, driveId) for i in data.get("value", [])]
        items = self._filter_items(items, extensions=extensions, foldersOnly=foldersOnly)
        return DriveItemsResult(
            driveId=driveId,
            parentItemId=itemId,
            path=path,
            items=items,
            nextLink=data.get("@odata.nextLink"),
        )

    async def search_in_drive(
        self,
        *,
        driveId: str,
        query: str,
        top: int = 50,
        extensions: list[str] | None = None,
    ) -> DriveItemsResult:
        """Search for files/folders by name within a single document library."""
        data = await self._request(
            f"/drives/{quote(driveId, safe='')}/root/search(q='{self._q(query)}')"
            f"?$top={min(top, 200)}"
            "&$select=id,name,folder,file,size,webUrl,lastModifiedDateTime,parentReference"
        )
        items = [self._map_item(i, driveId) for i in data.get("value", [])]
        items = self._filter_items(items, extensions=extensions)
        return DriveItemsResult(driveId=driveId, items=items)

    async def search_items(
        self,
        *,
        query: str,
        top: int = 25,
        extensions: list[str] | None = None,
    ) -> DriveItemsResult:
        """Search across everything the user can access (all sites + OneDrive),
        using the Graph Search API. Best for 'find this folder/file anywhere'."""
        body = {
            "requests": [
                {
                    "entityTypes": ["driveItem"],
                    "query": {"queryString": query},
                    "from": 0,
                    "size": min(top, 200),
                }
            ]
        }
        data = await self._request("/search/query", method="POST", json_body=body)
        items: list[DriveItemInfo] = []
        for response in data.get("value", []):
            for container in response.get("hitsContainers", []):
                for hit in container.get("hits", []):
                    resource = hit.get("resource") or {}
                    parent = resource.get("parentReference") or {}
                    items.append(
                        self._map_item(resource, parent.get("driveId"))
                    )
        items = self._filter_items(items, extensions=extensions)
        return DriveItemsResult(items=items)

    async def get_item_by_share_url(self, *, shareUrl: str) -> DriveItemInfo:
        """Resolve a SharePoint/OneDrive sharing or browser URL to a driveItem
        (with driveId+itemId you can then browse or edit)."""
        encoded = self._encode_share_url(shareUrl)
        data = await self._request(
            f"/shares/{encoded}/driveItem"
            "?$select=id,name,folder,file,size,webUrl,lastModifiedDateTime,parentReference"
        )
        parent = data.get("parentReference") or {}
        return self._map_item(data, parent.get("driveId"))

    # ---- mapping / helpers ------------------------------------------------ #
    @staticmethod
    def _map_site(data: dict[str, Any]) -> SiteInfo:
        return SiteInfo(
            id=str(data.get("id")),
            name=data.get("name"),
            displayName=data.get("displayName"),
            webUrl=data.get("webUrl"),
        )

    @staticmethod
    def _map_item(data: dict[str, Any], driveId: str | None) -> DriveItemInfo:
        folder = data.get("folder")
        name = data.get("name") or ""
        ext = None
        if "." in name and not folder:
            ext = name.rsplit(".", 1)[-1].lower()
        parent = data.get("parentReference") or {}
        return DriveItemInfo(
            name=name,
            itemId=str(data.get("id")),
            driveId=driveId or parent.get("driveId"),
            isFolder=folder is not None,
            childCount=(folder or {}).get("childCount") if folder else None,
            size=data.get("size"),
            fileExtension=ext,
            webUrl=data.get("webUrl"),
            lastModifiedDateTime=data.get("lastModifiedDateTime"),
            path=parent.get("path"),
        )

    @staticmethod
    def _filter_items(
        items: list[DriveItemInfo],
        *,
        extensions: list[str] | None = None,
        foldersOnly: bool = False,
    ) -> list[DriveItemInfo]:
        result = items
        if foldersOnly:
            result = [i for i in result if i.isFolder]
        if extensions:
            wanted = {e.lower().lstrip(".") for e in extensions}
            result = [
                i
                for i in result
                if i.isFolder or (i.fileExtension and i.fileExtension in wanted)
            ]
        return result

    @staticmethod
    def _encode_share_url(url: str) -> str:
        b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii")
        return "u!" + b64.rstrip("=")

    @staticmethod
    def _q(value: str) -> str:
        return value.replace("'", "''")

    @asynccontextmanager
    async def _client(self) -> Any:
        if self._http_client is not None:
            yield self._http_client
            return
        async with httpx.AsyncClient(follow_redirects=False, timeout=30.0) as client:
            yield client

    async def _request(
        self,
        path: str,
        *,
        method: str = "GET",
        json_body: dict[str, Any] | None = None,
    ) -> Any:
        access_token = await self._auth_service.get_access_token()
        # No Prefer: IdType="ImmutableId" here. That header is an Outlook
        # (mail/calendar/contacts) feature and is ignored by Graph for
        # /drives and /sites driveItems. Omitting it keeps the IDs returned by
        # this browse client identical to the default driveItem IDs the
        # excel_workbook.py client expects, so a browsed itemId round-trips
        # cleanly into the workbook tools (verified against live Graph).
        headers = {
            "Accept": "application/json",
            "Authorization": f"Bearer {access_token}",
            **({"Content-Type": "application/json"} if json_body is not None else {}),
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
                f"Microsoft Graph Files request failed "
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
