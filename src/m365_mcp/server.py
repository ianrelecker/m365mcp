import contextlib
import socket
import sys
from collections.abc import Callable, Iterator
from dataclasses import dataclass
from typing import Annotated, Literal
from urllib.parse import urljoin

import anyio
import httpx
import uvicorn
from mcp.server.fastmcp import FastMCP
from pydantic import Field

from m365_mcp.config import AppConfig, load_config
from m365_mcp.helper_app import create_helper_app
from m365_mcp.microsoft_auth import MicrosoftAuthService
from m365_mcp.microsoft_graph import MicrosoftGraphClient
from m365_mcp.models import (
    AuthStatusResult,
    CalendarCreateEventResult,
    CalendarListEventsResult,
    MailCreateDraftResult,
    MailGetResult,
    MailListDraftsResult,
    MailListResult,
    MailMoveResult,
    MailSearchResult,
    MailSendDraftResult,
)


@dataclass
class RuntimeServices:
    config: AppConfig
    microsoft_auth: MicrosoftAuthService
    graph: MicrosoftGraphClient
    http_client: httpx.AsyncClient
    owns_http_client: bool = False
    start_helper_server: bool = True


def create_runtime(
    config: AppConfig | None = None,
    *,
    http_client: httpx.AsyncClient | None = None,
    start_helper_server: bool = True,
) -> RuntimeServices:
    resolved_config = config or load_config()
    resolved_http_client = http_client or httpx.AsyncClient(
        follow_redirects=False,
        timeout=30.0,
    )
    auth = MicrosoftAuthService(resolved_config, resolved_http_client)
    graph = MicrosoftGraphClient(auth, resolved_http_client)
    return RuntimeServices(
        config=resolved_config,
        microsoft_auth=auth,
        graph=graph,
        http_client=resolved_http_client,
        owns_http_client=http_client is None,
        start_helper_server=start_helper_server,
    )


class _RuntimeProvider:
    def __init__(self, factory: Callable[[], RuntimeServices]) -> None:
        self._factory = factory
        self._runtime: RuntimeServices | None = None

    def get(self) -> RuntimeServices:
        if self._runtime is None:
            self._runtime = self._factory()
        return self._runtime

    def reset(self) -> None:
        self._runtime = None


class _HelperServerRunner:
    def __init__(self, runtime: RuntimeServices) -> None:
        self._runtime = runtime
        self._server: uvicorn.Server | None = None

    async def run(self, *, task_status: anyio.abc.TaskStatus[None]) -> None:
        if not _can_bind_localhost(self._runtime.config.port):
            print(
                "Claude M365 MCP local helper was not started because "
                f"localhost:{self._runtime.config.port} is already in use. "
                "Close the other process using that port, or set PORT and "
                "LOCAL_BASE_URL to a different localhost port that also matches "
                "the Azure redirect URI.",
                file=sys.stderr,
            )
            task_status.started()
            return

        app = create_helper_app(self._runtime.config, self._runtime.microsoft_auth)
        config = uvicorn.Config(
            app,
            host="localhost",
            port=self._runtime.config.port,
            log_level="warning",
            access_log=False,
            lifespan="off",
        )
        server = uvicorn.Server(config)
        server.install_signal_handlers = lambda: None
        self._server = server

        async def wait_until_started() -> None:
            with anyio.fail_after(5):
                while not server.started and not server.should_exit:
                    await anyio.sleep(0.05)

            if not server.started:
                raise RuntimeError(
                    f"Failed to start local helper server on port {self._runtime.config.port}"
                )

            print(
                f"Claude M365 MCP local helper listening on port {self._runtime.config.port}",
                file=sys.stderr,
            )
            print(
                f"Local helper URL: {self._runtime.config.localBaseUrl}",
                file=sys.stderr,
            )
            print(
                f"Microsoft callback URI: {self._runtime.config.microsoft.redirectUri}",
                file=sys.stderr,
            )
            task_status.started()

        async with anyio.create_task_group() as task_group:
            task_group.start_soon(wait_until_started)
            try:
                await server.serve()
            finally:
                task_group.cancel_scope.cancel()

    async def stop(self) -> None:
        if self._server is not None:
            self._server.should_exit = True


def _can_bind_localhost(port: int) -> bool:
    addresses: list[tuple[int, tuple[str, int] | tuple[str, int, int, int]]] = [
        (socket.AF_INET, ("127.0.0.1", port)),
    ]
    if socket.has_ipv6:
        addresses.append((socket.AF_INET6, ("::1", port, 0, 0)))

    for family, address in addresses:
        sock = socket.socket(family, socket.SOCK_STREAM)
        try:
            sock.bind(address)
        except OSError:
            return False
        finally:
            sock.close()

    return True


def _create_server(runtime_provider: _RuntimeProvider) -> FastMCP:
    @contextlib.asynccontextmanager
    async def lifespan(_app: FastMCP) -> Iterator[dict[str, object]]:
        runtime = runtime_provider.get()
        helper_runner = (
            _HelperServerRunner(runtime) if runtime.start_helper_server else None
        )

        async with anyio.create_task_group() as task_group:
            if helper_runner is not None:
                await task_group.start(helper_runner.run)
            try:
                yield {"config": runtime.config}
            finally:
                if helper_runner is not None:
                    await helper_runner.stop()
                if runtime.owns_http_client:
                    await runtime.http_client.aclose()
                runtime_provider.reset()

    mcp = FastMCP("claude-m365-mcp", lifespan=lifespan)

    @mcp.tool(
        name="auth_status",
        description=(
            "Check whether the server is connected to Microsoft 365 and see any "
            "known delegated mailbox hints."
        ),
    )
    async def auth_status() -> AuthStatusResult:
        runtime = runtime_provider.get()
        status = await runtime.microsoft_auth.get_status()
        return AuthStatusResult(
            connected=status.connected,
            account=status.account,
            expiresAt=status.expiresAt,
            knownMailboxes=status.knownMailboxes,
            localStatusUrl=runtime.config.localBaseUrl,
            microsoftConnectUrl=urljoin(
                runtime.config.localBaseUrl, "/auth/microsoft/start"
            ),
            microsoftDisconnectUrl=urljoin(
                runtime.config.localBaseUrl, "/auth/microsoft/disconnect"
            ),
        )

    @mcp.tool(
        name="mail_list",
        description=(
            "List messages from a mailbox folder. Use mailbox for shared/delegated "
            "mailboxes the signed-in Microsoft user can access."
        ),
    )
    async def mail_list(
        mailbox: str | None = None,
        folder: str = "Inbox",
        top: int = 25,
    ) -> MailListResult:
        runtime = runtime_provider.get()
        return await runtime.graph.list_messages(mailbox=mailbox, folder=folder, top=top)

    @mcp.tool(
        name="mail_search",
        description=(
            "Search a mailbox using Microsoft Graph $search. Use mailbox for "
            "shared/delegated mailboxes."
        ),
    )
    async def mail_search(
        query: str,
        mailbox: str | None = None,
        top: int = 10,
    ) -> MailSearchResult:
        runtime = runtime_provider.get()
        return await runtime.graph.search_messages(mailbox=mailbox, query=query, top=top)

    @mcp.tool(
        name="mail_get",
        description=(
            "Get the full details and body of one message by ID. Use mailbox for "
            "shared/delegated mailboxes."
        ),
    )
    async def mail_get(
        messageId: str,
        mailbox: str | None = None,
    ) -> MailGetResult:
        runtime = runtime_provider.get()
        return await runtime.graph.get_message(mailbox=mailbox, messageId=messageId)

    @mcp.tool(
        name="mail_list_drafts",
        description=(
            "List draft messages from the default Drafts folder. Use mailbox for "
            "shared/delegated mailboxes."
        ),
    )
    async def mail_list_drafts(
        mailbox: str | None = None,
        top: int = 25,
    ) -> MailListDraftsResult:
        runtime = runtime_provider.get()
        return await runtime.graph.list_drafts(mailbox=mailbox, top=top)

    @mcp.tool(
        name="mail_create_draft",
        description=(
            "Create a new draft email. Use mailbox for shared/delegated mailboxes "
            "and optionally set from when you need a specific sender."
        ),
    )
    async def mail_create_draft(
        subject: str,
        body: str,
        mailbox: str | None = None,
        to: list[str] = [],
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        bodyType: Literal["text", "html"] = "text",
        from_: Annotated[
            str | None,
            Field(validation_alias="from", serialization_alias="from_"),
        ] = None,
    ) -> MailCreateDraftResult:
        runtime = runtime_provider.get()
        return await runtime.graph.create_draft(
            mailbox=mailbox,
            subject=subject,
            to=to,
            cc=cc,
            bcc=bcc,
            body=body,
            bodyType=bodyType,
            from_=from_,
        )

    @mcp.tool(
        name="mail_send_draft",
        description=(
            "Send an existing draft message by ID. Use mailbox for shared/delegated mailboxes."
        ),
    )
    async def mail_send_draft(
        messageId: str,
        mailbox: str | None = None,
    ) -> MailSendDraftResult:
        runtime = runtime_provider.get()
        return await runtime.graph.send_draft(mailbox=mailbox, messageId=messageId)

    @mcp.tool(
        name="mail_move",
        description=(
            "Move a message to another folder. Pass a well-known folder name like "
            "Archive or DeletedItems, or set destinationFolderIsId when passing a raw folder ID."
        ),
    )
    async def mail_move(
        messageId: str,
        destinationFolder: str,
        mailbox: str | None = None,
        destinationFolderIsId: bool = False,
    ) -> MailMoveResult:
        runtime = runtime_provider.get()
        return await runtime.graph.move_message(
            mailbox=mailbox,
            messageId=messageId,
            destinationFolder=destinationFolder,
            destinationFolderIsId=destinationFolderIsId,
        )

    @mcp.tool(
        name="calendar_list_events",
        description=(
            "List events in the default calendar over a time window. Use mailbox for "
            "shared/delegated calendars."
        ),
    )
    async def calendar_list_events(
        mailbox: str | None = None,
        start: str | None = None,
        end: str | None = None,
        top: int = 25,
    ) -> CalendarListEventsResult:
        runtime = runtime_provider.get()
        return await runtime.graph.list_events(
            mailbox=mailbox,
            start=start,
            end=end,
            top=top,
        )

    @mcp.tool(
        name="calendar_create_event",
        description=(
            "Create an event in the default calendar. Use mailbox for shared/delegated calendars."
        ),
    )
    async def calendar_create_event(
        subject: str,
        start: str,
        end: str,
        mailbox: str | None = None,
        timeZone: str = "UTC",
        attendees: list[str] | None = None,
        body: str | None = None,
        bodyType: Literal["text", "html"] = "text",
        location: str | None = None,
    ) -> CalendarCreateEventResult:
        runtime = runtime_provider.get()
        return await runtime.graph.create_event(
            mailbox=mailbox,
            subject=subject,
            start=start,
            end=end,
            timeZone=timeZone,
            attendees=attendees,
            body=body,
            bodyType=bodyType,
            location=location,
        )

    return mcp


def create_mcp_server(runtime: RuntimeServices) -> FastMCP:
    return _create_server(_RuntimeProvider(lambda: runtime))


def create_default_server() -> FastMCP:
    return _create_server(_RuntimeProvider(create_runtime))


mcp = create_default_server()
app = mcp


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
