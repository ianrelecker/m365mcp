from __future__ import annotations

import base64
import json
from datetime import date

import httpx
import pytest

from m365_mcp.excel_workbook import ExcelWorkbookClient, WorkbookItemRef


class StaticAuthService:
    async def get_access_token(self) -> str:
        return "access-token"


def _make_client(handler) -> tuple[ExcelWorkbookClient, httpx.AsyncClient]:
    http_client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    return ExcelWorkbookClient(StaticAuthService(), http_client), http_client


def _item() -> WorkbookItemRef:
    return WorkbookItemRef(driveId="drive-1", itemId="item-1")


@pytest.mark.anyio
async def test_resolve_workbook_by_share_url_encodes_and_maps() -> None:
    share_url = "https://contoso.sharepoint.com/:x:/r/Doc.xlsx"
    expected = "u!" + base64.urlsafe_b64encode(
        share_url.encode("utf-8")
    ).decode("ascii").rstrip("=")

    def handler(request: httpx.Request) -> httpx.Response:
        assert request.headers["authorization"] == "Bearer access-token"
        assert request.url.path == f"/v1.0/shares/{expected}/driveItem"
        return httpx.Response(
            200,
            json={
                "id": "wb-1",
                "name": "Doc.xlsx",
                "webUrl": "https://contoso.sharepoint.com/Doc.xlsx",
                "parentReference": {"driveId": "drive-42"},
            },
        )

    client, http_client = _make_client(handler)
    ref = await client.resolve_workbook(shareUrl=share_url)
    assert ref.driveId == "drive-42"
    assert ref.itemId == "wb-1"
    assert ref.name == "Doc.xlsx"
    await http_client.aclose()


@pytest.mark.anyio
async def test_resolve_workbook_by_drive_and_path() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        assert request.url.path == "/v1.0/drives/drive-1/root:/Folder/Book.xlsx:"
        return httpx.Response(200, json={"id": "wb-2", "name": "Book.xlsx"})

    client, http_client = _make_client(handler)
    ref = await client.resolve_workbook(
        driveId="drive-1", itemPath="/Folder/Book.xlsx/"
    )
    assert ref.itemId == "wb-2"
    assert ref.driveId == "drive-1"
    await http_client.aclose()


@pytest.mark.anyio
async def test_resolve_workbook_requires_an_identifier() -> None:
    client, http_client = _make_client(
        lambda request: httpx.Response(500)
    )
    with pytest.raises(ValueError):
        await client.resolve_workbook()
    await http_client.aclose()


@pytest.mark.anyio
async def test_list_worksheets_maps_and_targets_workbook_base() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        assert (
            request.url.path
            == "/v1.0/drives/drive-1/items/item-1/workbook/worksheets"
        )
        return httpx.Response(
            200,
            json={
                "value": [
                    {
                        "id": "{ABC}",
                        "name": "Sheet1",
                        "position": 0,
                        "visibility": "Visible",
                    }
                ]
            },
        )

    client, http_client = _make_client(handler)
    result = await client.list_worksheets(_item())
    assert result.worksheets[0].name == "Sheet1"
    assert result.worksheets[0].position == 0
    await http_client.aclose()


@pytest.mark.anyio
async def test_list_tables_for_worksheet_escapes_name() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        # Worksheet name with an apostrophe must be OData-escaped (' -> '').
        # The space is percent-encoded by httpx; assert on the decoded path.
        assert "worksheets('Bob''s Sheet')/tables" in request.url.path
        return httpx.Response(
            200,
            json={"value": [{"id": "t1", "name": "Data", "showHeaders": True}]},
        )

    client, http_client = _make_client(handler)
    result = await client.list_tables(_item(), worksheet="Bob's Sheet")
    assert result.tables[0].name == "Data"
    assert result.tables[0].worksheet == "Bob's Sheet"
    await http_client.aclose()


@pytest.mark.anyio
async def test_get_range_escapes_worksheet_and_address() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        url = str(request.url)
        assert "worksheets('Sheet1')" in url
        assert "range(address='A1:O5')" in url
        return httpx.Response(
            200,
            json={
                "address": "Sheet1!A1:O5",
                "values": [[1, 2], [3, 4]],
                "text": [["1", "2"], ["3", "4"]],
                "rowCount": 2,
                "columnCount": 2,
            },
        )

    client, http_client = _make_client(handler)
    result = await client.get_range(_item(), worksheet="Sheet1", address="A1:O5")
    assert result.values == [[1, 2], [3, 4]]
    assert result.rowCount == 2
    assert result.address == "Sheet1!A1:O5"
    await http_client.aclose()


@pytest.mark.anyio
async def test_get_used_range_values_only_suffix() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        assert "usedRange(valuesOnly=true)" in str(request.url)
        return httpx.Response(
            200,
            json={"address": "Sheet1!A1:B2", "values": [["x"]], "rowCount": 1},
        )

    client, http_client = _make_client(handler)
    result = await client.get_used_range(_item(), worksheet="Sheet1")
    assert result.address == "Sheet1!A1:B2"
    await http_client.aclose()


@pytest.mark.anyio
async def test_update_range_patches_values_and_number_format() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        assert request.method == "PATCH"
        assert "range(address='A2:B2')" in str(request.url)
        body = json.loads(request.content)
        assert body["values"] == [[1, 2]]
        assert body["numberFormat"] == [["General", "mm/dd/yy"]]
        return httpx.Response(200, json={"address": "Sheet1!A2:B2"})

    client, http_client = _make_client(handler)
    result = await client.update_range(
        _item(),
        worksheet="Sheet1",
        address="A2:B2",
        values=[[1, 2]],
        numberFormat=[["General", "mm/dd/yy"]],
    )
    assert result.updated is True
    assert result.address == "A2:B2"
    await http_client.aclose()


@pytest.mark.anyio
async def test_update_range_requires_payload() -> None:
    client, http_client = _make_client(lambda request: httpx.Response(500))
    with pytest.raises(ValueError):
        await client.update_range(_item(), worksheet="Sheet1", address="A1")
    await http_client.aclose()


@pytest.mark.anyio
async def test_add_table_row_posts_values_and_index() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        assert request.method == "POST"
        assert (
            request.url.path
            == "/v1.0/drives/drive-1/items/item-1/workbook/tables/Table1/rows/add"
        )
        body = json.loads(request.content)
        assert body["values"] == [["a", "b"]]
        assert body["index"] == 0
        return httpx.Response(200, json={"index": 0, "values": [["a", "b"]]})

    client, http_client = _make_client(handler)
    result = await client.add_table_row(
        _item(), table="Table1", values=[["a", "b"]], index=0
    )
    assert result.table == "Table1"
    assert result.index == 0
    await http_client.aclose()


@pytest.mark.anyio
async def test_session_header_is_sent_when_session_id_present() -> None:
    seen: dict[str, str | None] = {}

    def handler(request: httpx.Request) -> httpx.Response:
        seen["session"] = request.headers.get("workbook-session-id")
        return httpx.Response(200, json={"value": []})

    client, http_client = _make_client(handler)
    await client.list_worksheets(_item(), sessionId="session-xyz")
    assert seen["session"] == "session-xyz"
    await http_client.aclose()


@pytest.mark.anyio
async def test_create_and_close_session() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path.endswith("/workbook/createSession"):
            assert json.loads(request.content) == {"persistChanges": True}
            return httpx.Response(200, json={"id": "sess-1"})
        if request.url.path.endswith("/workbook/closeSession"):
            assert request.headers["workbook-session-id"] == "sess-1"
            return httpx.Response(204)
        raise AssertionError(f"Unexpected request: {request.url}")

    client, http_client = _make_client(handler)
    session = await client.create_session(_item(), persistChanges=True)
    assert session.sessionId == "sess-1"
    assert session.persistChanges is True
    await client.close_session(_item(), sessionId="sess-1")
    await http_client.aclose()


@pytest.mark.anyio
async def test_request_error_raises_with_graph_detail() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(
            404,
            json={
                "error": {
                    "code": "ItemNotFound",
                    "message": "The requested resource doesn't exist.",
                }
            },
        )

    client, http_client = _make_client(handler)
    with pytest.raises(RuntimeError) as excinfo:
        await client.get_range(_item(), worksheet="Nope", address="A1")
    message = str(excinfo.value)
    assert "404" in message
    assert "ItemNotFound" in message
    await http_client.aclose()


def test_excel_serial_date_conversion() -> None:
    # Excel's epoch sentinel: 1899-12-31 is serial 1.
    assert ExcelWorkbookClient.excel_serial_date(date(1899, 12, 31)) == 1
    # Known reference: 2024-01-01 is Excel serial 45292.
    assert ExcelWorkbookClient.excel_serial_date(date(2024, 1, 1)) == 45292
    # Accepts ISO strings too.
    assert ExcelWorkbookClient.excel_serial_date("2024-01-01") == 45292


def test_encode_share_url_strips_padding() -> None:
    encoded = ExcelWorkbookClient._encode_share_url("https://a/b?c=1")
    assert encoded.startswith("u!")
    assert "=" not in encoded
    payload = encoded[2:]
    padded = payload + "=" * (-len(payload) % 4)
    assert base64.urlsafe_b64decode(padded).decode("utf-8") == "https://a/b?c=1"


def test_wb_base_quotes_ids() -> None:
    ref = WorkbookItemRef(driveId="d/1", itemId="i 1")
    base = ExcelWorkbookClient._wb_base(ref)
    # driveId and itemId are quoted with safe='' so '/' and ' ' are escaped.
    assert base == "/drives/d%2F1/items/i%201/workbook"
