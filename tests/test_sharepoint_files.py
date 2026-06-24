from __future__ import annotations

import base64
import json

import httpx
import pytest

from m365_mcp.sharepoint_files import SharePointFilesClient


class StaticAuthService:
    async def get_access_token(self) -> str:
        return "access-token"


def _make_client(handler) -> tuple[SharePointFilesClient, httpx.AsyncClient]:
    http_client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    return SharePointFilesClient(StaticAuthService(), http_client), http_client


@pytest.mark.anyio
async def test_search_sites_builds_query_and_maps_response() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        assert request.headers["authorization"] == "Bearer access-token"
        # Item 3 regression guard: no Outlook ImmutableId Prefer header here.
        assert "prefer" not in request.headers
        assert request.url.path == "/v1.0/sites"
        assert request.url.params["search"] == "acme deals"
        assert request.url.params["$top"] == "100"  # clamped from 500
        return httpx.Response(
            200,
            json={
                "value": [
                    {
                        "id": "site-1",
                        "name": "acme",
                        "displayName": "Acme Deals",
                        "webUrl": "https://contoso.sharepoint.com/sites/acme",
                    }
                ]
            },
        )

    client, http_client = _make_client(handler)
    result = await client.search_sites(query="acme deals", top=500)
    assert result.query == "acme deals"
    assert result.sites[0].id == "site-1"
    assert result.sites[0].displayName == "Acme Deals"
    await http_client.aclose()


@pytest.mark.anyio
async def test_get_site_by_path_constructs_colon_path() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        assert (
            request.url.path
            == "/v1.0/sites/contoso.sharepoint.com:/sites/Acquisitions"
        )
        return httpx.Response(200, json={"id": "site-9", "displayName": "Acq"})

    client, http_client = _make_client(handler)
    site = await client.get_site_by_path(
        hostname="contoso.sharepoint.com", sitePath="/Acquisitions/"
    )
    assert site.id == "site-9"
    await http_client.aclose()


@pytest.mark.anyio
async def test_list_drives_maps_libraries() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        assert request.url.path == "/v1.0/sites/site-1/drives"
        return httpx.Response(
            200,
            json={
                "value": [
                    {
                        "id": "drive-1",
                        "name": "Documents",
                        "driveType": "documentLibrary",
                        "webUrl": "https://contoso.sharepoint.com/Docs",
                    }
                ]
            },
        )

    client, http_client = _make_client(handler)
    result = await client.list_drives(siteId="site-1")
    assert result.siteId == "site-1"
    assert result.drives[0].driveType == "documentLibrary"
    await http_client.aclose()


@pytest.mark.anyio
async def test_list_children_by_item_maps_folders_and_files() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        assert request.url.path == "/v1.0/drives/drive-1/items/item-1/children"
        assert request.url.params["$top"] == "999"  # clamped from 5000
        return httpx.Response(
            200,
            json={
                "value": [
                    {
                        "id": "f1",
                        "name": "Reports",
                        "folder": {"childCount": 3},
                        "parentReference": {"path": "/drive/root:/Reports"},
                    },
                    {
                        "id": "x1",
                        "name": "Budget.xlsx",
                        "file": {"mimeType": "application/vnd.ms-excel"},
                        "size": 4096,
                        "parentReference": {
                            "driveId": "drive-1",
                            "path": "/drive/root:",
                        },
                    },
                ],
                "@odata.nextLink": "https://graph.microsoft.com/next",
            },
        )

    client, http_client = _make_client(handler)
    result = await client.list_children(
        driveId="drive-1", itemId="item-1", top=5000
    )
    assert result.parentItemId == "item-1"
    assert result.nextLink == "https://graph.microsoft.com/next"
    folder, xlsx = result.items
    assert folder.isFolder is True
    assert folder.childCount == 3
    assert xlsx.isFolder is False
    assert xlsx.fileExtension == "xlsx"
    assert xlsx.driveId == "drive-1"
    await http_client.aclose()


@pytest.mark.anyio
async def test_list_children_by_path_uses_root_colon_addressing() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        assert (
            request.url.path
            == "/v1.0/drives/drive-1/root:/Shared Active Deals/Q1:/children"
        )
        return httpx.Response(200, json={"value": []})

    client, http_client = _make_client(handler)
    result = await client.list_children(
        driveId="drive-1", path="/Shared Active Deals/Q1/"
    )
    assert result.items == []
    await http_client.aclose()


@pytest.mark.anyio
async def test_list_children_extension_and_folder_filters() -> None:
    value = {
        "value": [
            {"id": "f1", "name": "Sub", "folder": {"childCount": 0}},
            {"id": "a1", "name": "a.xlsx", "file": {}},
            {"id": "b1", "name": "b.pdf", "file": {}},
            {"id": "c1", "name": "c.docx", "file": {}},
        ]
    }

    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(200, json=value)

    client, http_client = _make_client(handler)

    only_xlsx = await client.list_children(
        driveId="d", itemId="i", extensions=[".XLSX"]
    )
    names = {i.name for i in only_xlsx.items}
    # Folders are always kept; only xlsx files pass the extension filter.
    assert names == {"Sub", "a.xlsx"}

    folders = await client.list_children(driveId="d", itemId="i", foldersOnly=True)
    assert [i.name for i in folders.items] == ["Sub"]
    await http_client.aclose()


@pytest.mark.anyio
async def test_search_in_drive_escapes_single_quotes() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        # OData single-quote escaping: O'Brien -> O''Brien inside search(q='...').
        assert "search(q='O''Brien')" in str(request.url)
        return httpx.Response(200, json={"value": []})

    client, http_client = _make_client(handler)
    result = await client.search_in_drive(driveId="drive-1", query="O'Brien")
    assert result.driveId == "drive-1"
    await http_client.aclose()


@pytest.mark.anyio
async def test_search_items_posts_search_query_and_parses_hits() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        assert request.method == "POST"
        assert request.url.path == "/v1.0/search/query"
        body = json.loads(request.content)
        req = body["requests"][0]
        assert req["entityTypes"] == ["driveItem"]
        assert req["query"]["queryString"] == "tracker"
        assert req["size"] == 200  # clamped from 999
        return httpx.Response(
            200,
            json={
                "value": [
                    {
                        "hitsContainers": [
                            {
                                "hits": [
                                    {
                                        "resource": {
                                            "id": "hit-1",
                                            "name": "Tracker.xlsx",
                                            "file": {},
                                            "parentReference": {"driveId": "drive-9"},
                                        }
                                    }
                                ]
                            }
                        ]
                    }
                ]
            },
        )

    client, http_client = _make_client(handler)
    result = await client.search_items(query="tracker", top=999)
    assert result.items[0].itemId == "hit-1"
    assert result.items[0].driveId == "drive-9"
    assert result.items[0].fileExtension == "xlsx"
    await http_client.aclose()


@pytest.mark.anyio
async def test_get_item_by_share_url_encodes_url() -> None:
    share_url = "https://contoso.sharepoint.com/:x:/r/sites/a/Doc.xlsx?web=1"
    expected = "u!" + base64.urlsafe_b64encode(
        share_url.encode("utf-8")
    ).decode("ascii").rstrip("=")

    def handler(request: httpx.Request) -> httpx.Response:
        assert request.url.path == f"/v1.0/shares/{expected}/driveItem"
        return httpx.Response(
            200,
            json={
                "id": "drv-item",
                "name": "Doc.xlsx",
                "file": {},
                "parentReference": {"driveId": "drive-7"},
            },
        )

    client, http_client = _make_client(handler)
    item = await client.get_item_by_share_url(shareUrl=share_url)
    assert item.itemId == "drv-item"
    assert item.driveId == "drive-7"
    await http_client.aclose()


@pytest.mark.anyio
async def test_request_error_raises_with_graph_detail() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(
            403,
            json={
                "error": {
                    "code": "accessDenied",
                    "message": "Access denied to the resource.",
                }
            },
        )

    client, http_client = _make_client(handler)
    with pytest.raises(RuntimeError) as excinfo:
        await client.list_drives(siteId="site-1")
    message = str(excinfo.value)
    assert "403" in message
    assert "accessDenied" in message
    assert "Access denied to the resource." in message
    await http_client.aclose()


def test_encode_share_url_matches_graph_addressing() -> None:
    encoded = SharePointFilesClient._encode_share_url("https://a/b c?d=1")
    assert encoded.startswith("u!")
    assert "=" not in encoded  # padding stripped
    # Round-trips back to the original URL.
    payload = encoded[2:]
    padded = payload + "=" * (-len(payload) % 4)
    assert base64.urlsafe_b64decode(padded).decode("utf-8") == "https://a/b c?d=1"


def test_q_escapes_single_quote() -> None:
    assert SharePointFilesClient._q("it's a 'test'") == "it''s a ''test''"
