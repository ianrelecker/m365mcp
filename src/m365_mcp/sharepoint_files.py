"""SharePoint / OneDrive browsing and sharing for the M365 MCP server.

Lets Claude find folders anywhere the signed-in user has access and list their
contents (Excel, PDF, etc.) — without mounting the folder locally. Pairs with
excel_workbook.py: use this to *locate* a workbook, then hand the returned
driveId+itemId to the Workbook tools to edit it in place.

Same design as the other clients: shared MicrosoftAuthService + httpx client,
private _request with identical error handling, typed pydantic results.

Scopes (delegated):
    Sites.Read.All   +  Files.ReadWrite.All
  Browsing would work with Files.Read.All, but sharing and the companion
  Workbook client (excel_workbook.py) need Files.ReadWrite.All. The shared
  config also requests the read-only Sites.Read.All; no SharePoint-wide write
  scope (Sites.ReadWrite.All) is used.

Endpoint reference (Graph v1.0):
    Search sites:        GET  /sites?search={q}
    Site by path:        GET  /sites/{hostname}:/sites/{path}
    Site's libraries:    GET  /sites/{siteId}/drives
    Folder children:     GET  /drives/{driveId}/root/children
                         GET  /drives/{driveId}/items/{itemId}/children
                         GET  /drives/{driveId}/root:/{path}:/children
    Search in a drive:   GET  /drives/{driveId}/root/search(q='{q}')
    Search everywhere:   POST /search/query   (entityTypes: driveItem)
    Resolve a link:      GET    /shares/{u!encoded}/driveItem
    List permissions:    GET    /drives/{driveId}/items/{itemId}/permissions
    Create link:         POST   /drives/{driveId}/items/{itemId}/createLink
    Grant access:        POST   /drives/{driveId}/items/{itemId}/invite
    Revoke permission:   DELETE /drives/{driveId}/items/{itemId}/permissions/{id}
"""

from __future__ import annotations

import base64
from contextlib import asynccontextmanager
from typing import Any
from urllib.parse import quote

import httpx
from pydantic import BaseModel, Field

from .microsoft_auth import MicrosoftAuthService
from .pid_policy import Location, PidPolicy, labels_from_graph

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


class SharingIdentity(BaseModel):
    identityType: str | None = None
    id: str | None = None
    displayName: str | None = None
    email: str | None = None
    loginName: str | None = None


class SharingPermissionInfo(BaseModel):
    permissionId: str
    roles: list[str] = Field(default_factory=list)
    linkType: str | None = None
    linkScope: str | None = None
    shareUrl: str | None = None
    expirationDateTime: str | None = None
    inherited: bool = False
    grantedTo: list[SharingIdentity] = Field(default_factory=list)


class SharingPermissionsResult(BaseModel):
    driveId: str
    itemId: str
    permissions: list[SharingPermissionInfo] = Field(default_factory=list)
    nextLink: str | None = None


class SharingPermissionResult(BaseModel):
    driveId: str
    itemId: str
    permission: SharingPermissionInfo


class SharingGrantResult(BaseModel):
    driveId: str
    itemId: str
    permissions: list[SharingPermissionInfo] = Field(default_factory=list)
    failures: list["SharingGrantFailure"] = Field(default_factory=list)


class SharingGrantFailure(BaseModel):
    recipient: str | None = None
    code: str | None = None
    message: str


class SharingRevokeResult(BaseModel):
    driveId: str
    itemId: str
    permissionId: str
    revoked: bool = True


# --------------------------------------------------------------------------- #
# Client
# --------------------------------------------------------------------------- #
class SharePointFilesClient:
    def __init__(
        self,
        auth_service: MicrosoftAuthService,
        http_client: httpx.AsyncClient | None = None,
        pid_policy: PidPolicy | None = None,
    ) -> None:
        self._auth_service = auth_service
        self._http_client = http_client
        self._pid_policy = pid_policy or PidPolicy.disabled()

    # ---- discovery -------------------------------------------------------- #
    async def search_sites(self, *, query: str, top: int = 25) -> SitesResult:
        """Find SharePoint sites by keyword (name/title)."""
        data = await self._request(
            f"/sites?search={quote(query)}&$top={min(top, 100)}"
            "&$select=id,name,displayName,webUrl"
        )
        sites = [
            site
            for site in (self._map_site(s) for s in data.get("value", []))
            if self._pid_policy.allows_location(self._site_location(site))
        ]
        return SitesResult(
            query=query,
            sites=sites,
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
        site = self._map_site(data)
        self._pid_policy.require_location(self._site_location(site))
        return site

    async def list_drives(self, *, siteId: str) -> DrivesResult:
        """List a site's document libraries (each library is a 'drive')."""
        self._pid_policy.require_location(Location(site_id=siteId))
        data = await self._request(
            f"/sites/{quote(siteId, safe='')}/drives"
            "?$select=id,name,driveType,webUrl"
        )
        drives = [
            DriveInfo(
                id=d["id"],
                name=d.get("name"),
                driveType=d.get("driveType"),
                webUrl=d.get("webUrl"),
            )
            for d in data.get("value", [])
        ]
        drives = [
            drive
            for drive in drives
            if self._pid_policy.allows_location(
                Location(
                    site_id=siteId,
                    drive_id=drive.id,
                    web_url=drive.webUrl,
                )
            )
        ]
        return DrivesResult(
            siteId=siteId,
            drives=drives,
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
        self._pid_policy.require_location(
            Location(drive_id=driveId, item_id=itemId, path=path)
        )
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
        items = [
            item
            for item in (
                self._keep_mapped(i, driveId) for i in data.get("value", [])
            )
            if item is not None
        ]
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
        self._pid_policy.require_location(Location(drive_id=driveId))
        data = await self._request(
            f"/drives/{quote(driveId, safe='')}/root/search(q='{self._q(query)}')"
            f"?$top={min(top, 200)}"
            "&$select=id,name,folder,file,size,webUrl,lastModifiedDateTime,parentReference"
        )
        items = [
            item
            for item in (
                self._keep_mapped(i, driveId) for i in data.get("value", [])
            )
            if item is not None
        ]
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
                    item = self._keep_mapped(resource, parent.get("driveId"))
                    if item is not None:
                        items.append(item)
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
        item = self._map_item(data, parent.get("driveId"))
        self._pid_policy.require_location(
            self._item_location(item, labels=labels_from_graph(data))
        )
        return item

    # ---- sharing ---------------------------------------------------------- #
    async def list_permissions(
        self, *, driveId: str, itemId: str
    ) -> SharingPermissionsResult:
        """List the effective permissions visible to the signed-in user."""
        self._pid_policy.require_location(Location(drive_id=driveId, item_id=itemId))
        data = await self._request(f"{self._item_path(driveId, itemId)}/permissions")
        raw_permissions = list(data.get("value", []))
        next_link = data.get("@odata.nextLink")
        seen_links: set[str] = set()
        while next_link:
            if next_link in seen_links:
                raise RuntimeError("Microsoft Graph returned a repeated permissions page.")
            seen_links.add(next_link)
            data = await self._request(next_link)
            raw_permissions.extend(data.get("value", []))
            next_link = data.get("@odata.nextLink")
        return SharingPermissionsResult(
            driveId=driveId,
            itemId=itemId,
            permissions=[
                self._map_permission(permission)
                for permission in raw_permissions
            ],
            nextLink=None,
        )

    async def create_link(
        self,
        *,
        driveId: str,
        itemId: str,
        linkType: str,
        scope: str,
        expirationDateTime: str | None = None,
    ) -> SharingPermissionResult:
        """Create or return an existing sharing link for an item."""
        self._pid_policy.require_location(Location(drive_id=driveId, item_id=itemId))
        body: dict[str, Any] = {
            "type": linkType,
            "scope": scope,
            "retainInheritedPermissions": True,
        }
        if expirationDateTime:
            body["expirationDateTime"] = expirationDateTime
        data = await self._request(
            f"{self._item_path(driveId, itemId)}/createLink",
            method="POST",
            json_body=body,
        )
        return SharingPermissionResult(
            driveId=driveId,
            itemId=itemId,
            permission=self._map_permission(data),
        )

    async def grant_access(
        self,
        *,
        driveId: str,
        itemId: str,
        recipients: list[str],
        role: str,
        sendInvitation: bool = False,
        message: str | None = None,
    ) -> SharingGrantResult:
        """Grant named recipients read or write access to an item."""
        self._pid_policy.require_location(Location(drive_id=driveId, item_id=itemId))
        addresses = [address.strip() for address in recipients if address.strip()]
        if not addresses:
            raise ValueError("At least one recipient email address is required.")
        body: dict[str, Any] = {
            "recipients": [{"email": address} for address in addresses],
            "roles": [role],
            "requireSignIn": True,
            "sendInvitation": sendInvitation,
            "retainInheritedPermissions": True,
        }
        if message:
            body["message"] = message
        data = await self._request(
            f"{self._item_path(driveId, itemId)}/invite",
            method="POST",
            json_body=body,
        )
        values = data.get("value", []) if isinstance(data, dict) else []
        permissions: list[SharingPermissionInfo] = []
        failures: list[SharingGrantFailure] = []
        for index, value in enumerate(values):
            error = value.get("error") if isinstance(value, dict) else None
            permission = self._map_permission(value)
            if permission.permissionId or permission.roles or permission.grantedTo:
                permissions.append(permission)
            if isinstance(error, dict):
                recipient = (value.get("recipient") or {}).get("email")
                if not recipient:
                    recipient = next(
                        (
                            identity.email
                            for identity in permission.grantedTo
                            if identity.email
                        ),
                        None,
                    )
                failures.append(
                    SharingGrantFailure(
                        recipient=recipient
                        or (addresses[index] if index < len(addresses) else None),
                        code=error.get("code"),
                        message=str(error.get("message") or "Unknown Graph error"),
                    )
                )
        return SharingGrantResult(
            driveId=driveId,
            itemId=itemId,
            permissions=permissions,
            failures=failures,
        )

    async def revoke_permission(
        self, *, driveId: str, itemId: str, permissionId: str
    ) -> SharingRevokeResult:
        """Revoke a non-inherited direct permission or entire sharing link."""
        self._pid_policy.require_location(Location(drive_id=driveId, item_id=itemId))
        await self._request(
            f"{self._item_path(driveId, itemId)}/permissions/"
            f"{quote(permissionId, safe='')}",
            method="DELETE",
        )
        return SharingRevokeResult(
            driveId=driveId,
            itemId=itemId,
            permissionId=permissionId,
        )

    # ---- mapping / helpers ------------------------------------------------ #
    def _site_location(self, site: SiteInfo) -> Location:
        return Location(site_id=site.id, web_url=site.webUrl)

    def _item_location(
        self,
        item: DriveItemInfo,
        *,
        labels: tuple[str, ...] = (),
    ) -> Location:
        return Location(
            drive_id=item.driveId,
            item_id=item.itemId,
            path=item.path,
            web_url=item.webUrl,
            labels=labels,
        )

    def _keep_mapped(
        self, data: dict[str, Any], driveId: str | None
    ) -> DriveItemInfo | None:
        item = self._map_item(data, driveId)
        if not self._pid_policy.allows_location(
            self._item_location(item, labels=labels_from_graph(data))
        ):
            return None
        return item

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

    @classmethod
    def _map_permission(cls, data: dict[str, Any]) -> SharingPermissionInfo:
        link = data.get("link") or {}
        identities: list[SharingIdentity] = []
        raw_identities: list[dict[str, Any]] = []
        if data.get("grantedToV2"):
            raw_identities.append(data["grantedToV2"])
        raw_identities.extend(data.get("grantedToIdentitiesV2") or [])
        invitation = data.get("invitation") or {}
        if invitation.get("email"):
            raw_identities.append({"user": {"email": invitation["email"]}})
        for raw in raw_identities:
            for identity_type in (
                "user",
                "siteUser",
                "group",
                "siteGroup",
                "application",
                "siteApplication",
                "device",
            ):
                identity = raw.get(identity_type)
                if not isinstance(identity, dict):
                    continue
                identities.append(
                    SharingIdentity(
                        identityType=identity_type,
                        id=identity.get("id"),
                        displayName=identity.get("displayName"),
                        email=identity.get("email"),
                        loginName=identity.get("loginName"),
                    )
                )
        return SharingPermissionInfo(
            permissionId=str(data.get("id") or ""),
            roles=[str(role) for role in data.get("roles", [])],
            linkType=link.get("type"),
            linkScope=link.get("scope"),
            shareUrl=link.get("webUrl"),
            expirationDateTime=data.get("expirationDateTime"),
            inherited=data.get("inheritedFrom") is not None,
            grantedTo=identities,
        )

    @staticmethod
    def _item_path(driveId: str, itemId: str) -> str:
        return (
            f"/drives/{quote(driveId, safe='')}/items/"
            f"{quote(itemId, safe='')}"
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
