from __future__ import annotations

from urllib.parse import unquote

import httpx
import pytest

from m365_mcp.audit import audit_metadata, classify_tool, redact_error_message
from m365_mcp.microsoft_graph import MicrosoftGraphClient
from m365_mcp.sharepoint_files import (
    MAX_UPLOAD_NAME_CHARS,
    SharePointFilesClient,
    safe_drive_item_name,
)

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
XLSM = "application/vnd.ms-excel.sheet.macroEnabled.12"
PAYLOAD = b"PK\x03\x04 pretend workbook bytes"


class StaticAuthService:
    async def get_access_token(self) -> str:
        return "access-token"


def _copy_graph(
    *,
    name: str = "Deal Model.xlsx",
    content_type: str = XLSX,
    odata_type: str = "#microsoft.graph.fileAttachment",
    size: int | None = None,
    payload: bytes = PAYLOAD,
    uploaded_name: str | None = None,
    uploads: list[httpx.Request] | None = None,
    downloads: list[int] | None = None,
) -> tuple[httpx.AsyncClient, MicrosoftGraphClient, SharePointFilesClient]:
    """Mock Graph: attachment metadata, its /$value bytes, and the drive PUT.

    ``uploaded_name`` is the name Graph gives the new file (a clash renames
    it); by default it keeps the requested name.
    """

    def handler(request: httpx.Request) -> httpx.Response:
        assert request.headers["authorization"] == "Bearer access-token"
        path = request.url.path
        if path.endswith("/messages/msg-1/attachments/xl-1"):
            select = request.url.params.get("$select", "")
            assert select and "contentBytes" not in select
            return httpx.Response(
                200,
                json={
                    "@odata.type": odata_type,
                    "id": "xl-1",
                    "name": name,
                    "contentType": content_type,
                    "size": len(payload) if size is None else size,
                },
            )
        if path.endswith("/messages/msg-1/attachments/xl-1/$value"):
            if downloads is not None:
                downloads.append(len(payload))
            return httpx.Response(200, content=payload)
        if request.method == "PUT" and path.startswith("/v1.0/drives/drive-1/"):
            if uploads is not None:
                uploads.append(request)
            requested = unquote(path.rsplit(":/", 1)[0].rsplit("/", 1)[-1])
            return httpx.Response(
                201,
                json={
                    "id": "new-item",
                    "name": uploaded_name or requested,
                    "size": len(request.content),
                    "webUrl": "https://kcbm.sharepoint.com/new-item",
                    "file": {"mimeType": request.headers["content-type"]},
                    "parentReference": {"driveId": "drive-1", "id": "folder-1"},
                },
            )
        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    return (
        client,
        MicrosoftGraphClient(StaticAuthService(), client),
        SharePointFilesClient(StaticAuthService(), client),
    )


@pytest.mark.anyio
async def test_copy_uploads_attachment_bytes_into_folder_by_id() -> None:
    uploads: list[httpx.Request] = []
    client, graph, files = _copy_graph(uploads=uploads)

    result = await graph.copy_attachment_to_sharepoint(
        files=files,
        mailbox="acquisitions@example.com",
        messageId="msg-1",
        attachmentId="xl-1",
        driveId="drive-1",
        folderItemId="folder-1",
    )

    assert result.unsupportedReason is None
    assert (result.driveId, result.itemId) == ("drive-1", "new-item")
    assert result.fileName == result.requestedFileName == "Deal Model.xlsx"
    assert result.renamed is False
    assert result.macroEnabled is False
    assert result.size == len(PAYLOAD)
    assert result.mailbox == "acquisitions@example.com"
    (upload,) = uploads
    assert upload.url.path == "/v1.0/drives/drive-1/items/folder-1:/Deal Model.xlsx:/content"
    assert upload.url.params["@microsoft.graph.conflictBehavior"] == "rename"
    assert upload.headers["content-type"] == XLSX
    assert upload.content == PAYLOAD

    await client.aclose()


@pytest.mark.anyio
async def test_copy_addresses_folder_by_path_from_library_root() -> None:
    uploads: list[httpx.Request] = []
    client, graph, files = _copy_graph(uploads=uploads)

    await graph.copy_attachment_to_sharepoint(
        files=files,
        messageId="msg-1",
        attachmentId="xl-1",
        driveId="drive-1",
        folderPath="/Shared Active Deals/4. Claude Projects/Email Screening/Excel Saves/",
    )

    (upload,) = uploads
    assert upload.url.path == (
        "/v1.0/drives/drive-1/root:/Shared Active Deals/4. Claude Projects/"
        "Email Screening/Excel Saves/Deal Model.xlsx:/content"
    )
    assert b"%20" in upload.url.raw_path

    await client.aclose()


@pytest.mark.anyio
async def test_copy_reports_a_name_clash_as_renamed() -> None:
    client, graph, files = _copy_graph(uploaded_name="Deal Model 1.xlsx")

    result = await graph.copy_attachment_to_sharepoint(
        files=files,
        messageId="msg-1",
        attachmentId="xl-1",
        driveId="drive-1",
        folderItemId="folder-1",
    )

    assert result.renamed is True
    assert result.fileName == "Deal Model 1.xlsx"
    assert result.requestedFileName == "Deal Model.xlsx"

    await client.aclose()


@pytest.mark.anyio
async def test_copy_file_name_override_is_made_safe_and_keeps_extension() -> None:
    uploads: list[httpx.Request] = []
    client, graph, files = _copy_graph(uploads=uploads)

    result = await graph.copy_attachment_to_sharepoint(
        files=files,
        messageId="msg-1",
        attachmentId="xl-1",
        driveId="drive-1",
        folderItemId="folder-1",
        fileName="..\\..\\Woodlake: T12 v2",
    )

    assert result.requestedFileName == "Woodlake_ T12 v2.xlsx"
    assert uploads[0].url.path.endswith("/folder-1:/Woodlake_ T12 v2.xlsx:/content")

    await client.aclose()


@pytest.mark.anyio
async def test_copy_refuses_a_file_name_that_changes_the_extension() -> None:
    uploads: list[httpx.Request] = []
    downloads: list[int] = []
    client, graph, files = _copy_graph(
        name="Macro Model.xlsm", content_type=XLSM, uploads=uploads, downloads=downloads
    )

    with pytest.raises(ValueError, match=r"keep the attachment's \.xlsm extension"):
        await graph.copy_attachment_to_sharepoint(
            files=files,
            messageId="msg-1",
            attachmentId="xl-1",
            driveId="drive-1",
            folderItemId="folder-1",
            fileName="Macro Model.xlsx",
        )
    assert uploads == [] and downloads == []

    await client.aclose()


@pytest.mark.anyio
async def test_copy_names_an_extensionless_macro_workbook_from_its_type() -> None:
    uploads: list[httpx.Request] = []
    client, graph, files = _copy_graph(name="Model", content_type=XLSM, uploads=uploads)

    result = await graph.copy_attachment_to_sharepoint(
        files=files,
        messageId="msg-1",
        attachmentId="xl-1",
        driveId="drive-1",
        folderItemId="folder-1",
    )

    assert result.fileName == "Model.xlsm"
    assert result.macroEnabled is True
    assert uploads[0].headers["content-type"] == XLSM

    await client.aclose()


@pytest.mark.anyio
@pytest.mark.parametrize(
    ("overrides", "reason"),
    [
        ({"name": "legacy.xls", "content_type": "application/vnd.ms-excel"}, "Legacy .xls"),
        ({"name": "binary.xlsb", "content_type": "application/octet-stream"}, "Legacy .xls"),
        ({"name": "template.xltx", "content_type": "application/octet-stream"}, "templates"),
        ({"name": "notes.txt", "content_type": "text/plain"}, "not an .xlsx"),
        ({"size": 30_000_000}, "maxBytes=25000000"),
        ({"odata_type": "#microsoft.graph.itemAttachment"}, "Only file attachments"),
    ],
)
async def test_copy_refuses_before_downloading(overrides: dict, reason: str) -> None:
    uploads: list[httpx.Request] = []
    downloads: list[int] = []
    client, graph, files = _copy_graph(uploads=uploads, downloads=downloads, **overrides)

    result = await graph.copy_attachment_to_sharepoint(
        files=files,
        messageId="msg-1",
        attachmentId="xl-1",
        driveId="drive-1",
        folderItemId="folder-1",
    )

    assert reason in result.unsupportedReason
    assert result.itemId is None
    assert uploads == [] and downloads == []

    await client.aclose()


@pytest.mark.anyio
async def test_copy_stops_a_download_that_passes_max_bytes() -> None:
    # The reported size is small, but the bytes keep coming.
    uploads: list[httpx.Request] = []
    client, graph, files = _copy_graph(
        size=10, payload=b"x" * 2_000, uploads=uploads
    )

    result = await graph.copy_attachment_to_sharepoint(
        files=files,
        messageId="msg-1",
        attachmentId="xl-1",
        driveId="drive-1",
        folderItemId="folder-1",
        maxBytes=1_000,
    )

    assert "maxBytes=1000" in result.unsupportedReason
    assert uploads == []

    await client.aclose()


@pytest.mark.anyio
@pytest.mark.parametrize(
    "folder", [{}, {"folderItemId": "folder-1", "folderPath": "Excel Saves"}]
)
async def test_copy_needs_exactly_one_folder(folder: dict) -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)
    files = SharePointFilesClient(StaticAuthService(), client)

    with pytest.raises(ValueError, match="exactly one of folderItemId or folderPath"):
        await graph.copy_attachment_to_sharepoint(
            files=files, messageId="msg-1", attachmentId="xl-1", driveId="drive-1", **folder
        )

    await client.aclose()


@pytest.mark.anyio
async def test_upload_surfaces_graph_errors() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(
            403, json={"error": {"code": "accessDenied", "message": "No write access"}}
        )

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    files = SharePointFilesClient(StaticAuthService(), client)

    with pytest.raises(RuntimeError, match="403.*accessDenied"):
        await files.upload_new_file(
            driveId="drive-1",
            folderItemId="folder-1",
            fileName="a.xlsx",
            content=b"x",
            contentType=XLSX,
        )

    await client.aclose()


@pytest.mark.parametrize(
    ("name", "expected"),
    [
        ("Deal Model.xlsx", "Deal Model.xlsx"),
        ("Deal Model.XLSX", "Deal Model.xlsx"),
        ("Deal Model", "Deal Model.xlsx"),
        ("Deal.v2", "Deal.v2.xlsx"),
        ("C:\\Users\\x\\Startup\\run.xlsx", "run.xlsx"),
        ("../../etc/run.xlsx", "run.xlsx"),
        ('a"*:<>?|b.xlsx', "a_______b.xlsx"),
        ("~$Deal.xlsx", "Deal.xlsx"),
        ("  spaced   out . .xlsx", "spaced out.xlsx"),
        ("tab\tname.xlsx", "tab_name.xlsx"),
        ("", "attachment.xlsx"),
        ("...", "attachment.xlsx"),
        ("Deal Model.xlsx ", "Deal Model.xlsx"),
        ("Deal Model.xlsx.", "Deal Model.xlsx"),
    ],
)
def test_safe_drive_item_name(name: str, expected: str) -> None:
    assert safe_drive_item_name(name, extension=".xlsx") == expected


def test_safe_drive_item_name_caps_length() -> None:
    result = safe_drive_item_name("x" * 500 + ".xlsm", extension=".xlsm")
    assert len(result) == MAX_UPLOAD_NAME_CHARS
    assert result.endswith(".xlsm")


def test_copy_is_audited_as_a_write_with_file_name_redacted() -> None:
    arguments = {
        "messageId": "msg-1",
        "attachmentId": "xl-1",
        "driveId": "drive-1",
        "folderItemId": "folder-1",
        "fileName": "Woodlake T12.xlsx",
    }
    assert classify_tool("mail_copy_attachment_to_sharepoint") == "write"
    assert audit_metadata("mail_copy_attachment_to_sharepoint", arguments)["ids"] == {
        "messageId": "msg-1",
        "attachmentId": "xl-1",
        "driveId": "drive-1",
        "folderItemId": "folder-1",
    }
    assert "Woodlake" not in redact_error_message(
        "upload of Woodlake T12.xlsx failed", arguments
    )
