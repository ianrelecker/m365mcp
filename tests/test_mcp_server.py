from __future__ import annotations

import importlib.util
import socket
from pathlib import Path

import httpx
import pytest
from mcp.shared.memory import create_connected_server_and_client_session

from m365_mcp.models import (
    AccountInfo,
    AuthStatusResult,
    CalendarCreateEventResult,
    CalendarDateTime,
    CalendarEvent,
    MailCreateDraftResult,
    MessageBody,
    MessageSummary,
    MicrosoftConnectionStatus,
)
from m365_mcp.server import RuntimeServices, _can_bind_localhost, create_mcp_server


class StubAuthService:
    async def get_status(self) -> MicrosoftConnectionStatus:
        return MicrosoftConnectionStatus(
            connected=True,
            account=AccountInfo(preferredUsername="user@example.com"),
            expiresAt=1712345678901,
            knownMailboxes=["shared@example.com"],
        )


class StubGraphClient:
    def __init__(self) -> None:
        self.last_from: str | None = None

    async def list_messages(self, **kwargs):
        raise AssertionError("Not expected in this test")

    async def search_messages(self, **kwargs):
        raise AssertionError("Not expected in this test")

    async def get_message(self, **kwargs):
        raise AssertionError("Not expected in this test")

    async def list_drafts(self, **kwargs):
        raise AssertionError("Not expected in this test")

    async def create_draft(self, **kwargs) -> MailCreateDraftResult:
        self.last_from = kwargs["from_"]
        return MailCreateDraftResult(
            mailbox=kwargs["mailbox"] or "me",
            draft=MessageSummary(
                id="draft-1",
                subject=kwargs["subject"],
                from_=kwargs["from_"],
                receivedDateTime=None,
                sentDateTime=None,
                bodyPreview="Draft preview",
                webLink=None,
                isDraft=True,
                conversationId=None,
            ),
        )

    async def send_draft(self, **kwargs):
        raise AssertionError("Not expected in this test")

    async def move_message(self, **kwargs):
        raise AssertionError("Not expected in this test")

    async def list_events(self, **kwargs):
        raise AssertionError("Not expected in this test")

    async def create_event(self, **kwargs) -> CalendarCreateEventResult:
        return CalendarCreateEventResult(
            mailbox=kwargs["mailbox"] or "me",
            event=CalendarEvent(
                id="event-1",
                subject=kwargs["subject"],
                webLink=None,
                start=CalendarDateTime(dateTime=kwargs["start"], timeZone=kwargs["timeZone"]),
                end=CalendarDateTime(dateTime=kwargs["end"], timeZone=kwargs["timeZone"]),
                location=kwargs["location"],
                attendees=[],
                bodyPreview="",
                body=MessageBody(contentType="text", content=""),
            ),
        )


@pytest.mark.anyio
async def test_mcp_server_exposes_expected_tools_and_structured_outputs(config_factory) -> None:
    graph = StubGraphClient()
    http_client = httpx.AsyncClient()
    runtime = RuntimeServices(
        config=config_factory(localBaseUrl="http://localhost:8787"),
        microsoft_auth=StubAuthService(),
        graph=graph,
        http_client=http_client,
        owns_http_client=False,
        start_helper_server=False,
    )
    server = create_mcp_server(runtime)

    async with create_connected_server_and_client_session(server, raise_exceptions=True) as session:
        tools = await session.list_tools()
        assert {tool.name for tool in tools.tools} == {
            "m365_capabilities",
            "auth_status",
            "mail_check_inbox",
            "mail_list_folders",
            "mail_folder_tree",
            "mail_resolve_folder",
            "mail_create_folder",
            "mail_rename_folder",
            "mail_delete_folder",
            "mail_list_rules",
            "mail_create_rule",
            "mail_update_rule",
            "mail_delete_rule",
            "mail_list",
            "mail_search",
            "mail_get",
            "mail_list_drafts",
            "mail_create_draft",
            "mail_send",
            "mail_send_draft",
            "mail_move",
            "mail_list_attachments",
            "mail_get_attachment_content",
            "mail_get_thread",
            "mail_create_reply_draft",
            "mail_send_reply",
            "mail_list_categories",
            "mail_set_categories",
            "mail_add_categories",
            "mail_remove_categories",
            "mail_clear_categories",
            "mail_create_category",
            "mail_update_category",
            "mail_delete_category",
            "mail_mark_read",
            "mail_set_flag",
            "contacts_list",
            "contacts_search",
            "contacts_get",
            "contacts_create",
            "contacts_update",
            "contacts_delete",
            "contacts_set_categories",
            "contacts_add_categories",
            "contacts_remove_categories",
            "contacts_clear_categories",
            "contacts_list_folders",
            "calendar_list_events",
            "calendar_create_event",
            "calendar_update_event",
            "calendar_delete_event",
        }
        tool_by_name = {tool.name: tool for tool in tools.tools}
        contact_create_schema = tool_by_name["contacts_create"].inputSchema
        assert "categories" in contact_create_schema["properties"]
        assert "businessAddress" in contact_create_schema["properties"]
        assert "countryOrRegion" in str(contact_create_schema)

        resources = await session.list_resources()
        assert {str(resource.uri) for resource in resources.resources} == {
            "m365://capabilities"
        }

        capabilities = await session.call_tool("m365_capabilities", {})
        assert "M365 MCP Capabilities" in capabilities.structuredContent["content"]
        assert (
            "contacts_set_categories" in capabilities.structuredContent["content"]
        )

        resource = await session.read_resource("m365://capabilities")
        assert "M365 MCP Capabilities" in resource.contents[0].text

        auth_status = await session.call_tool("auth_status", {})
        assert auth_status.structuredContent == AuthStatusResult(
            connected=True,
            account=AccountInfo(preferredUsername="user@example.com"),
            expiresAt=1712345678901,
            knownMailboxes=["shared@example.com"],
            localStatusUrl="http://localhost:8787",
            microsoftConnectUrl="http://localhost:8787/auth/microsoft/start",
            microsoftDisconnectUrl="http://localhost:8787/auth/microsoft/disconnect",
        ).model_dump(mode="json", by_alias=True)

        draft = await session.call_tool(
            "mail_create_draft",
            {
                "subject": "Hello",
                "body": "Draft body",
                "mailbox": "shared@example.com",
                "to": ["a@example.com"],
                "from": "delegated@example.com",
            },
        )
        assert draft.structuredContent["draft"]["from"] == "delegated@example.com"
        assert graph.last_from == "delegated@example.com"

        event = await session.call_tool(
            "calendar_create_event",
            {
                "subject": "Planning",
                "start": "2026-04-22T16:00:00",
                "end": "2026-04-22T17:00:00",
                "mailbox": "shared@example.com",
            },
        )
        assert event.structuredContent["event"]["subject"] == "Planning"

    await http_client.aclose()


def test_server_file_imports_the_way_mcp_cli_imports_it() -> None:
    server_path = Path(__file__).parents[1] / "src" / "m365_mcp" / "server.py"
    spec = importlib.util.spec_from_file_location("mcp_cli_server_import_test", server_path)
    assert spec is not None
    assert spec.loader is not None

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    assert module.mcp is module.app


def test_can_bind_localhost_detects_busy_port() -> None:
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.bind(("127.0.0.1", 0))
    sock.listen()
    try:
        assert _can_bind_localhost(sock.getsockname()[1]) is False
    finally:
        sock.close()


@pytest.mark.anyio
async def test_mcp_server_stays_up_when_helper_port_is_busy(config_factory) -> None:
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.bind(("127.0.0.1", 0))
    sock.listen()
    port = sock.getsockname()[1]

    http_client = httpx.AsyncClient()
    runtime = RuntimeServices(
        config=config_factory(
            port=port,
            localBaseUrl=f"http://localhost:{port}",
        ),
        microsoft_auth=StubAuthService(),
        graph=StubGraphClient(),
        http_client=http_client,
        owns_http_client=False,
        start_helper_server=True,
    )
    server = create_mcp_server(runtime)

    try:
        async with create_connected_server_and_client_session(
            server,
            raise_exceptions=True,
        ) as session:
            auth_status = await session.call_tool("auth_status", {})
            assert auth_status.structuredContent["connected"] is True
    finally:
        sock.close()
        await http_client.aclose()
