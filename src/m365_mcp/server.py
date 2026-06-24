import contextlib
import socket
import sys
from collections.abc import Callable, Iterator, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Annotated, Any, Literal
from urllib.parse import urljoin

import anyio
import httpx
import uvicorn
from mcp.server.fastmcp import FastMCP
from mcp.types import ContentBlock
from pydantic import Field

from m365_mcp.audit import LocalAuditLogger
from m365_mcp.config import AppConfig, load_config
from m365_mcp.helper_app import create_helper_app
from m365_mcp.excel_workbook import (
    ExcelWorkbookClient,
    WorkbookCalculateResult,
    WorkbookClearResult,
    WorkbookCopyResult,
    WorkbookInsertResult,
    WorkbookItemRef,
    WorkbookListTablesResult,
    WorkbookListWorksheetsResult,
    WorkbookNameRangeResult,
    WorkbookNamesResult,
    WorkbookRangeResult,
    WorkbookRangesResult,
    WorkbookRowAddResult,
    WorkbookSessionResult,
    WorkbookWriteResult,
)
from m365_mcp.microsoft_auth import MicrosoftAuthService
from m365_mcp.microsoft_graph import MicrosoftGraphClient
from m365_mcp.sharepoint_files import (
    DriveItemInfo,
    DriveItemsResult,
    DrivesResult,
    SharePointFilesClient,
    SiteInfo,
    SitesResult,
)
from m365_mcp.models import (
    AuthStatusResult,
    CalendarCreateEventResult,
    CalendarDeleteEventResult,
    CalendarListEventsResult,
    CalendarUpdateEventResult,
    ContactAddress,
    ContactFoldersResult,
    ContactGetResult,
    ContactMutationResult,
    ContactsListResult,
    ContactsSearchResult,
    M365CapabilitiesResult,
    MailAttachmentContentResult,
    MailCategoryResult,
    MailCheckInboxResult,
    MailCreateDraftResult,
    MailFolderMutationResult,
    MailFolderTreeResult,
    MailGetResult,
    MailListAttachmentsResult,
    MailListCategoriesResult,
    MailListDraftsResult,
    MailListFoldersResult,
    MailListResult,
    MailListRulesResult,
    MailMoveResult,
    MailResolveFolderResult,
    MailRuleResult,
    MailSearchResult,
    MailSendDraftResult,
    MailSendResult,
    MailThreadResult,
    MailUpdateMessageResult,
)


CAPABILITIES_PATH = Path(__file__).parents[2] / "M365_MCP_CAPABILITIES.md"


@dataclass
class RuntimeServices:
    config: AppConfig
    microsoft_auth: MicrosoftAuthService
    graph: MicrosoftGraphClient
    sharepoint: SharePointFilesClient
    excel: ExcelWorkbookClient
    http_client: httpx.AsyncClient
    audit_logger: LocalAuditLogger | None = None
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
    sharepoint = SharePointFilesClient(auth, resolved_http_client)
    excel = ExcelWorkbookClient(auth, resolved_http_client)
    audit_logger = LocalAuditLogger(
        enabled=resolved_config.auditLogEnabled,
        file_path=resolved_config.auditLogFile,
    )
    return RuntimeServices(
        config=resolved_config,
        microsoft_auth=auth,
        graph=graph,
        sharepoint=sharepoint,
        excel=excel,
        http_client=resolved_http_client,
        audit_logger=audit_logger,
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


def _get_audit_logger(runtime: RuntimeServices) -> LocalAuditLogger:
    if runtime.audit_logger is None:
        runtime.audit_logger = LocalAuditLogger(
            enabled=runtime.config.auditLogEnabled,
            file_path=runtime.config.auditLogFile,
        )
    return runtime.audit_logger


class _AuditedFastMCP(FastMCP):
    def __init__(
        self,
        *args: object,
        runtime_provider: _RuntimeProvider,
        **kwargs: object,
    ) -> None:
        self._audit_runtime_provider = runtime_provider
        super().__init__(*args, **kwargs)

    async def call_tool(
        self,
        name: str,
        arguments: dict[str, Any],
    ) -> Sequence[ContentBlock] | dict[str, Any]:
        try:
            result = await super().call_tool(name, arguments)
        except Exception as error:
            await self._record_audit_event(
                tool_name=name,
                arguments=arguments,
                outcome="error",
                error=error,
            )
            raise

        await self._record_audit_event(
            tool_name=name,
            arguments=arguments,
            outcome="success",
        )
        return result

    async def _record_audit_event(
        self,
        *,
        tool_name: str,
        arguments: dict[str, Any],
        outcome: str,
        error: BaseException | None = None,
    ) -> None:
        try:
            runtime = self._audit_runtime_provider.get()
            await _get_audit_logger(runtime).record_tool_call(
                tool_name=tool_name,
                arguments=arguments,
                outcome=outcome,
                error=error,
            )
        except Exception as audit_error:
            print(
                f"Claude M365 MCP audit logging failed: {audit_error}",
                file=sys.stderr,
            )


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


def _load_capabilities_text() -> str:
    return CAPABILITIES_PATH.read_text("utf-8")


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

    mcp = _AuditedFastMCP(
        "claude-m365-mcp",
        lifespan=lifespan,
        runtime_provider=runtime_provider,
    )

    @mcp.resource(
        "m365://capabilities",
        name="m365_capabilities",
        title="M365 MCP Capabilities",
        description="Model-facing guide for using the M365 MCP server safely and effectively.",
        mime_type="text/markdown",
    )
    def m365_capabilities_resource() -> str:
        return _load_capabilities_text()

    @mcp.tool(
        name="m365_capabilities",
        description=(
            "Read the model-facing guide for what this M365 MCP server can do "
            "and how to use its inbox, folder, attachment, thread, category, "
            "contact, and calendar tools."
        ),
    )
    async def m365_capabilities() -> M365CapabilitiesResult:
        return M365CapabilitiesResult(content=_load_capabilities_text())

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
            requiredScopes=status.requiredScopes,
            grantedScopes=status.grantedScopes,
            missingScopes=status.missingScopes,
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
        folderId: str | None = None,
        folderPath: str | None = None,
        top: int = 25,
        isRead: bool | None = None,
        hasAttachments: bool | None = None,
        importance: Literal["low", "normal", "high"] | None = None,
        categories: list[str] | None = None,
        flagStatus: Literal["notFlagged", "flagged", "complete"] | None = None,
        inferenceClassification: Literal["focused", "other"] | None = None,
    ) -> MailListResult:
        runtime = runtime_provider.get()
        return await runtime.graph.list_messages(
            mailbox=mailbox,
            folder=folder,
            folderId=folderId,
            folderPath=folderPath,
            top=top,
            isRead=isRead,
            hasAttachments=hasAttachments,
            importance=importance,
            categories=categories,
            flagStatus=flagStatus,
            inferenceClassification=inferenceClassification,
        )

    @mcp.tool(
        name="mail_check_inbox",
        description=(
            "Quickly check an inbox or shared inbox folder. Defaults to unread "
            "messages in Inbox and includes triage metadata."
        ),
    )
    async def mail_check_inbox(
        mailbox: str | None = None,
        folderPath: str = "Inbox",
        folderId: str | None = None,
        top: int = 25,
        includeRead: bool = False,
        inferenceClassification: Literal["focused", "other"] | None = None,
    ) -> MailCheckInboxResult:
        runtime = runtime_provider.get()
        return await runtime.graph.check_inbox(
            mailbox=mailbox,
            folderPath=folderPath,
            folderId=folderId,
            top=top,
            includeRead=includeRead,
            inferenceClassification=inferenceClassification,
        )

    @mcp.tool(
        name="mail_list_folders",
        description=(
            "List top-level mail folders or direct child folders for own or shared mailboxes."
        ),
    )
    async def mail_list_folders(
        mailbox: str | None = None,
        parentFolderId: str | None = None,
        top: int = 100,
    ) -> MailListFoldersResult:
        runtime = runtime_provider.get()
        return await runtime.graph.list_mail_folders(
            mailbox=mailbox,
            parentFolderId=parentFolderId,
            top=top,
        )

    @mcp.tool(
        name="mail_folder_tree",
        description=(
            "Return a nested folder tree so the model can navigate subfolders by path or ID."
        ),
    )
    async def mail_folder_tree(
        mailbox: str | None = None,
        rootFolderId: str | None = None,
        maxDepth: int = 4,
    ) -> MailFolderTreeResult:
        runtime = runtime_provider.get()
        return await runtime.graph.mail_folder_tree(
            mailbox=mailbox,
            rootFolderId=rootFolderId,
            maxDepth=maxDepth,
        )

    @mcp.tool(
        name="mail_resolve_folder",
        description=(
            "Resolve a folder path like Inbox/Clients/Acme or a child display name to a folder ID."
        ),
    )
    async def mail_resolve_folder(
        mailbox: str | None = None,
        folderPath: str | None = None,
        parentFolderId: str | None = None,
        displayName: str | None = None,
    ) -> MailResolveFolderResult:
        runtime = runtime_provider.get()
        return await runtime.graph.resolve_mail_folder(
            mailbox=mailbox,
            folderPath=folderPath,
            parentFolderId=parentFolderId,
            displayName=displayName,
        )

    @mcp.tool(
        name="mail_create_folder",
        description=(
            "Create a top-level mail folder or a subfolder under parentFolderId/"
            "parentFolderPath."
        ),
    )
    async def mail_create_folder(
        displayName: str,
        mailbox: str | None = None,
        parentFolderId: str | None = None,
        parentFolderPath: str | None = None,
        isHidden: bool | None = None,
    ) -> MailFolderMutationResult:
        runtime = runtime_provider.get()
        return await runtime.graph.create_mail_folder(
            mailbox=mailbox,
            displayName=displayName,
            parentFolderId=parentFolderId,
            parentFolderPath=parentFolderPath,
            isHidden=isHidden,
        )

    @mcp.tool(
        name="mail_rename_folder",
        description="Rename a mail folder by folderId or folderPath.",
    )
    async def mail_rename_folder(
        displayName: str,
        mailbox: str | None = None,
        folderId: str | None = None,
        folderPath: str | None = None,
    ) -> MailFolderMutationResult:
        runtime = runtime_provider.get()
        return await runtime.graph.rename_mail_folder(
            mailbox=mailbox,
            displayName=displayName,
            folderId=folderId,
            folderPath=folderPath,
        )

    @mcp.tool(
        name="mail_delete_folder",
        description="Delete a mail folder by folderId or folderPath.",
    )
    async def mail_delete_folder(
        mailbox: str | None = None,
        folderId: str | None = None,
        folderPath: str | None = None,
    ) -> MailFolderMutationResult:
        runtime = runtime_provider.get()
        return await runtime.graph.delete_mail_folder(
            mailbox=mailbox,
            folderId=folderId,
            folderPath=folderPath,
        )

    @mcp.tool(
        name="mail_list_rules",
        description="List Outlook Inbox message rules for the mailbox.",
    )
    async def mail_list_rules(
        mailbox: str | None = None,
        top: int = 100,
    ) -> MailListRulesResult:
        runtime = runtime_provider.get()
        return await runtime.graph.list_mail_rules(mailbox=mailbox, top=top)

    @mcp.tool(
        name="mail_create_rule",
        description=(
            "Create an Outlook Inbox message rule. Use raw Graph conditions/actions "
            "or convenience fields such as senderContains and moveToFolderPath."
        ),
    )
    async def mail_create_rule(
        displayName: str,
        mailbox: str | None = None,
        sequence: int = 1,
        isEnabled: bool = True,
        conditions: dict[str, object] | None = None,
        actions: dict[str, object] | None = None,
        exceptions: dict[str, object] | None = None,
        fromAddresses: list[str] | None = None,
        senderContains: list[str] | None = None,
        subjectContains: list[str] | None = None,
        bodyContains: list[str] | None = None,
        sentToAddresses: list[str] | None = None,
        moveToFolderId: str | None = None,
        moveToFolderPath: str | None = None,
        markAsRead: bool | None = None,
        assignCategories: list[str] | None = None,
        stopProcessingRules: bool | None = None,
    ) -> MailRuleResult:
        runtime = runtime_provider.get()
        return await runtime.graph.create_mail_rule(
            mailbox=mailbox,
            displayName=displayName,
            sequence=sequence,
            isEnabled=isEnabled,
            conditions=conditions,
            actions=actions,
            exceptions=exceptions,
            fromAddresses=fromAddresses,
            senderContains=senderContains,
            subjectContains=subjectContains,
            bodyContains=bodyContains,
            sentToAddresses=sentToAddresses,
            moveToFolderId=moveToFolderId,
            moveToFolderPath=moveToFolderPath,
            markAsRead=markAsRead,
            assignCategories=assignCategories,
            stopProcessingRules=stopProcessingRules,
        )

    @mcp.tool(
        name="mail_update_rule",
        description="Update an Outlook Inbox message rule by ID.",
    )
    async def mail_update_rule(
        ruleId: str,
        mailbox: str | None = None,
        displayName: str | None = None,
        sequence: int | None = None,
        isEnabled: bool | None = None,
        conditions: dict[str, object] | None = None,
        actions: dict[str, object] | None = None,
        exceptions: dict[str, object] | None = None,
        fromAddresses: list[str] | None = None,
        senderContains: list[str] | None = None,
        subjectContains: list[str] | None = None,
        bodyContains: list[str] | None = None,
        sentToAddresses: list[str] | None = None,
        moveToFolderId: str | None = None,
        moveToFolderPath: str | None = None,
        markAsRead: bool | None = None,
        assignCategories: list[str] | None = None,
        stopProcessingRules: bool | None = None,
    ) -> MailRuleResult:
        runtime = runtime_provider.get()
        return await runtime.graph.update_mail_rule(
            mailbox=mailbox,
            ruleId=ruleId,
            displayName=displayName,
            sequence=sequence,
            isEnabled=isEnabled,
            conditions=conditions,
            actions=actions,
            exceptions=exceptions,
            fromAddresses=fromAddresses,
            senderContains=senderContains,
            subjectContains=subjectContains,
            bodyContains=bodyContains,
            sentToAddresses=sentToAddresses,
            moveToFolderId=moveToFolderId,
            moveToFolderPath=moveToFolderPath,
            markAsRead=markAsRead,
            assignCategories=assignCategories,
            stopProcessingRules=stopProcessingRules,
        )

    @mcp.tool(
        name="mail_delete_rule",
        description="Delete an Outlook Inbox message rule by ID.",
    )
    async def mail_delete_rule(
        ruleId: str,
        mailbox: str | None = None,
    ) -> MailRuleResult:
        runtime = runtime_provider.get()
        return await runtime.graph.delete_mail_rule(mailbox=mailbox, ruleId=ruleId)

    @mcp.tool(
        name="mail_search",
        description=(
            "Search a mailbox using Microsoft Graph $search. Use mailbox for "
            "shared/delegated mailboxes. Prefer mail_list filters for quick "
            "inbox triage because Graph search can be slower on large mailboxes."
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
        name="mail_send",
        description=(
            "Send a new email in one call. Prefer mail_create_draft first when "
            "the user has not explicitly approved sending."
        ),
    )
    async def mail_send(
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
        saveToSentItems: bool = True,
    ) -> MailSendResult:
        runtime = runtime_provider.get()
        return await runtime.graph.send_mail(
            mailbox=mailbox,
            subject=subject,
            to=to,
            cc=cc,
            bcc=bcc,
            body=body,
            bodyType=bodyType,
            from_=from_,
            saveToSentItems=saveToSentItems,
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
        destinationFolderId: str | None = None,
        destinationFolderPath: str | None = None,
    ) -> MailMoveResult:
        runtime = runtime_provider.get()
        return await runtime.graph.move_message(
            mailbox=mailbox,
            messageId=messageId,
            destinationFolder=destinationFolder,
            destinationFolderIsId=destinationFolderIsId,
            destinationFolderId=destinationFolderId,
            destinationFolderPath=destinationFolderPath,
        )

    @mcp.tool(
        name="mail_list_attachments",
        description="List attachment metadata for a message without downloading content.",
    )
    async def mail_list_attachments(
        messageId: str,
        mailbox: str | None = None,
        includeInline: bool = False,
    ) -> MailListAttachmentsResult:
        runtime = runtime_provider.get()
        return await runtime.graph.list_attachments(
            mailbox=mailbox,
            messageId=messageId,
            includeInline=includeInline,
        )

    @mcp.tool(
        name="mail_get_attachment_content",
        description=(
            "Read small text-like attachment content. Large or binary attachments "
            "return metadata with an unsupportedReason instead of content."
        ),
    )
    async def mail_get_attachment_content(
        messageId: str,
        attachmentId: str,
        mailbox: str | None = None,
        maxBytes: int = 1_000_000,
        maxChars: int = 100_000,
    ) -> MailAttachmentContentResult:
        runtime = runtime_provider.get()
        return await runtime.graph.get_attachment_content(
            mailbox=mailbox,
            messageId=messageId,
            attachmentId=attachmentId,
            maxBytes=maxBytes,
            maxChars=maxChars,
        )

    @mcp.tool(
        name="mail_get_thread",
        description=(
            "Get messages in the same conversation by messageId or conversationId."
        ),
    )
    async def mail_get_thread(
        messageId: str | None = None,
        conversationId: str | None = None,
        mailbox: str | None = None,
        top: int = 50,
    ) -> MailThreadResult:
        runtime = runtime_provider.get()
        return await runtime.graph.get_thread(
            mailbox=mailbox,
            messageId=messageId,
            conversationId=conversationId,
            top=top,
        )

    @mcp.tool(
        name="mail_create_reply_draft",
        description=(
            "Create a reply or reply-all draft in the original message thread. "
            "Use mail_send_draft later to send it."
        ),
    )
    async def mail_create_reply_draft(
        messageId: str,
        comment: str,
        mailbox: str | None = None,
        replyAll: bool = False,
        bodyType: Literal["text", "html"] = "html",
    ) -> MailCreateDraftResult:
        runtime = runtime_provider.get()
        return await runtime.graph.create_reply_draft(
            mailbox=mailbox,
            messageId=messageId,
            comment=comment,
            replyAll=replyAll,
            bodyType=bodyType,
        )

    @mcp.tool(
        name="mail_send_reply",
        description=(
            "Send a reply or reply-all immediately in the original thread. "
            "Prefer mail_create_reply_draft first when approval is not explicit."
        ),
    )
    async def mail_send_reply(
        messageId: str,
        comment: str,
        mailbox: str | None = None,
        replyAll: bool = False,
        bodyType: Literal["text", "html"] = "html",
    ) -> MailSendResult:
        runtime = runtime_provider.get()
        return await runtime.graph.send_reply(
            mailbox=mailbox,
            messageId=messageId,
            comment=comment,
            replyAll=replyAll,
            bodyType=bodyType,
        )

    @mcp.tool(
        name="mail_list_categories",
        description="List Outlook master categories for a mailbox.",
    )
    async def mail_list_categories(
        mailbox: str | None = None,
    ) -> MailListCategoriesResult:
        runtime = runtime_provider.get()
        return await runtime.graph.list_categories(mailbox=mailbox)

    @mcp.tool(
        name="mail_set_categories",
        description="Replace the categories assigned to a message.",
    )
    async def mail_set_categories(
        messageId: str,
        categories: list[str],
        mailbox: str | None = None,
    ) -> MailUpdateMessageResult:
        runtime = runtime_provider.get()
        return await runtime.graph.set_message_categories(
            mailbox=mailbox,
            messageId=messageId,
            categories=categories,
        )

    @mcp.tool(
        name="mail_add_categories",
        description="Add categories to a message without removing existing categories.",
    )
    async def mail_add_categories(
        messageId: str,
        categories: list[str],
        mailbox: str | None = None,
    ) -> MailUpdateMessageResult:
        runtime = runtime_provider.get()
        return await runtime.graph.add_message_categories(
            mailbox=mailbox,
            messageId=messageId,
            categories=categories,
        )

    @mcp.tool(
        name="mail_remove_categories",
        description="Remove selected categories from a message.",
    )
    async def mail_remove_categories(
        messageId: str,
        categories: list[str],
        mailbox: str | None = None,
    ) -> MailUpdateMessageResult:
        runtime = runtime_provider.get()
        return await runtime.graph.remove_message_categories(
            mailbox=mailbox,
            messageId=messageId,
            categories=categories,
        )

    @mcp.tool(
        name="mail_clear_categories",
        description="Remove all categories from a message.",
    )
    async def mail_clear_categories(
        messageId: str,
        mailbox: str | None = None,
    ) -> MailUpdateMessageResult:
        runtime = runtime_provider.get()
        return await runtime.graph.clear_message_categories(
            mailbox=mailbox,
            messageId=messageId,
        )

    @mcp.tool(
        name="mail_create_category",
        description="Create an Outlook master category definition.",
    )
    async def mail_create_category(
        displayName: str,
        color: str = "preset0",
        mailbox: str | None = None,
    ) -> MailCategoryResult:
        runtime = runtime_provider.get()
        return await runtime.graph.create_category(
            mailbox=mailbox,
            displayName=displayName,
            color=color,
        )

    @mcp.tool(
        name="mail_update_category",
        description=(
            "Update an Outlook master category color. Microsoft Graph does not "
            "support renaming existing master categories."
        ),
    )
    async def mail_update_category(
        categoryId: str,
        displayName: Annotated[
            str | None,
            Field(
                description=(
                    "Unsupported for existing Outlook master categories; create "
                    "a new category instead of renaming."
                )
            ),
        ] = None,
        color: Annotated[
            str | None,
            Field(description="New Outlook category color, for example preset3."),
        ] = None,
        mailbox: str | None = None,
    ) -> MailCategoryResult:
        runtime = runtime_provider.get()
        return await runtime.graph.update_category(
            mailbox=mailbox,
            categoryId=categoryId,
            displayName=displayName,
            color=color,
        )

    @mcp.tool(
        name="mail_delete_category",
        description="Delete an Outlook master category definition.",
    )
    async def mail_delete_category(
        categoryId: str,
        mailbox: str | None = None,
    ) -> MailCategoryResult:
        runtime = runtime_provider.get()
        return await runtime.graph.delete_category(
            mailbox=mailbox,
            categoryId=categoryId,
        )

    @mcp.tool(
        name="mail_mark_read",
        description="Mark a message read or unread.",
    )
    async def mail_mark_read(
        messageId: str,
        isRead: bool = True,
        mailbox: str | None = None,
    ) -> MailUpdateMessageResult:
        runtime = runtime_provider.get()
        return await runtime.graph.mark_message_read(
            mailbox=mailbox,
            messageId=messageId,
            isRead=isRead,
        )

    @mcp.tool(
        name="mail_set_flag",
        description="Set a message follow-up flag status.",
    )
    async def mail_set_flag(
        messageId: str,
        flagStatus: Literal["notFlagged", "flagged", "complete"],
        mailbox: str | None = None,
        startDateTime: str | None = None,
        dueDateTime: str | None = None,
    ) -> MailUpdateMessageResult:
        runtime = runtime_provider.get()
        return await runtime.graph.set_message_flag(
            mailbox=mailbox,
            messageId=messageId,
            flagStatus=flagStatus,
            startDateTime=startDateTime,
            dueDateTime=dueDateTime,
        )

    @mcp.tool(
        name="contacts_list",
        description="List Outlook contacts from the default or specified contact folder.",
    )
    async def contacts_list(
        mailbox: str | None = None,
        folderId: str | None = None,
        top: int = 25,
    ) -> ContactsListResult:
        runtime = runtime_provider.get()
        return await runtime.graph.list_contacts(
            mailbox=mailbox,
            folderId=folderId,
            top=top,
        )

    @mcp.tool(
        name="contacts_search",
        description="Search Outlook contacts by exact email or client-side name/email matching.",
    )
    async def contacts_search(
        query: str,
        mailbox: str | None = None,
        folderId: str | None = None,
        top: int = 25,
        maxPages: int = 5,
    ) -> ContactsSearchResult:
        runtime = runtime_provider.get()
        return await runtime.graph.search_contacts(
            mailbox=mailbox,
            query=query,
            folderId=folderId,
            top=top,
            maxPages=maxPages,
        )

    @mcp.tool(
        name="contacts_get",
        description="Get one Outlook contact by ID.",
    )
    async def contacts_get(
        contactId: str,
        mailbox: str | None = None,
        folderId: str | None = None,
    ) -> ContactGetResult:
        runtime = runtime_provider.get()
        return await runtime.graph.get_contact(
            mailbox=mailbox,
            contactId=contactId,
            folderId=folderId,
        )

    @mcp.tool(
        name="contacts_create",
        description="Create an Outlook contact with optional categories and addresses.",
    )
    async def contacts_create(
        mailbox: str | None = None,
        folderId: str | None = None,
        displayName: str | None = None,
        givenName: str | None = None,
        surname: str | None = None,
        emailAddresses: list[str] | None = None,
        companyName: str | None = None,
        jobTitle: str | None = None,
        personalHomePage: str | None = None,
        personalNotes: str | None = None,
        businessPhones: list[str] | None = None,
        mobilePhone: str | None = None,
        categories: list[str] | None = None,
        businessAddress: ContactAddress | None = None,
        homeAddress: ContactAddress | None = None,
        otherAddress: ContactAddress | None = None,
    ) -> ContactMutationResult:
        runtime = runtime_provider.get()
        return await runtime.graph.create_contact(
            mailbox=mailbox,
            folderId=folderId,
            displayName=displayName,
            givenName=givenName,
            surname=surname,
            emailAddresses=emailAddresses,
            companyName=companyName,
            jobTitle=jobTitle,
            personalHomePage=personalHomePage,
            personalNotes=personalNotes,
            businessPhones=businessPhones,
            mobilePhone=mobilePhone,
            categories=categories,
            businessAddress=businessAddress,
            homeAddress=homeAddress,
            otherAddress=otherAddress,
        )

    @mcp.tool(
        name="contacts_update",
        description=(
            "Update an Outlook contact by ID, including names, phones, email "
            "addresses, categories, and physical addresses."
        ),
    )
    async def contacts_update(
        contactId: str,
        mailbox: str | None = None,
        folderId: str | None = None,
        displayName: str | None = None,
        givenName: str | None = None,
        surname: str | None = None,
        emailAddresses: list[str] | None = None,
        companyName: str | None = None,
        jobTitle: str | None = None,
        personalHomePage: str | None = None,
        personalNotes: str | None = None,
        businessPhones: list[str] | None = None,
        mobilePhone: str | None = None,
        categories: list[str] | None = None,
        businessAddress: ContactAddress | None = None,
        homeAddress: ContactAddress | None = None,
        otherAddress: ContactAddress | None = None,
    ) -> ContactMutationResult:
        runtime = runtime_provider.get()
        return await runtime.graph.update_contact(
            mailbox=mailbox,
            contactId=contactId,
            folderId=folderId,
            displayName=displayName,
            givenName=givenName,
            surname=surname,
            emailAddresses=emailAddresses,
            companyName=companyName,
            jobTitle=jobTitle,
            personalHomePage=personalHomePage,
            personalNotes=personalNotes,
            businessPhones=businessPhones,
            mobilePhone=mobilePhone,
            categories=categories,
            businessAddress=businessAddress,
            homeAddress=homeAddress,
            otherAddress=otherAddress,
        )

    @mcp.tool(
        name="contacts_delete",
        description="Delete an Outlook contact by ID.",
    )
    async def contacts_delete(
        contactId: str,
        mailbox: str | None = None,
        folderId: str | None = None,
    ) -> ContactMutationResult:
        runtime = runtime_provider.get()
        return await runtime.graph.delete_contact(
            mailbox=mailbox,
            contactId=contactId,
            folderId=folderId,
        )

    @mcp.tool(
        name="contacts_set_categories",
        description="Replace all categories on an Outlook contact.",
    )
    async def contacts_set_categories(
        contactId: str,
        categories: list[str],
        mailbox: str | None = None,
        folderId: str | None = None,
    ) -> ContactMutationResult:
        runtime = runtime_provider.get()
        return await runtime.graph.set_contact_categories(
            mailbox=mailbox,
            contactId=contactId,
            folderId=folderId,
            categories=categories,
        )

    @mcp.tool(
        name="contacts_add_categories",
        description="Add categories to an Outlook contact without removing existing categories.",
    )
    async def contacts_add_categories(
        contactId: str,
        categories: list[str],
        mailbox: str | None = None,
        folderId: str | None = None,
    ) -> ContactMutationResult:
        runtime = runtime_provider.get()
        return await runtime.graph.add_contact_categories(
            mailbox=mailbox,
            contactId=contactId,
            folderId=folderId,
            categories=categories,
        )

    @mcp.tool(
        name="contacts_remove_categories",
        description="Remove selected categories from an Outlook contact.",
    )
    async def contacts_remove_categories(
        contactId: str,
        categories: list[str],
        mailbox: str | None = None,
        folderId: str | None = None,
    ) -> ContactMutationResult:
        runtime = runtime_provider.get()
        return await runtime.graph.remove_contact_categories(
            mailbox=mailbox,
            contactId=contactId,
            folderId=folderId,
            categories=categories,
        )

    @mcp.tool(
        name="contacts_clear_categories",
        description="Remove all categories from an Outlook contact.",
    )
    async def contacts_clear_categories(
        contactId: str,
        mailbox: str | None = None,
        folderId: str | None = None,
    ) -> ContactMutationResult:
        runtime = runtime_provider.get()
        return await runtime.graph.clear_contact_categories(
            mailbox=mailbox,
            contactId=contactId,
            folderId=folderId,
        )

    @mcp.tool(
        name="contacts_list_folders",
        description="List Outlook contact folders or direct child contact folders.",
    )
    async def contacts_list_folders(
        mailbox: str | None = None,
        parentFolderId: str | None = None,
        top: int = 100,
    ) -> ContactFoldersResult:
        runtime = runtime_provider.get()
        return await runtime.graph.list_contact_folders(
            mailbox=mailbox,
            parentFolderId=parentFolderId,
            top=top,
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

    @mcp.tool(
        name="calendar_update_event",
        description="Update an event in the default calendar by event ID.",
    )
    async def calendar_update_event(
        eventId: str,
        mailbox: str | None = None,
        subject: str | None = None,
        start: str | None = None,
        end: str | None = None,
        timeZone: str = "UTC",
        attendees: list[str] | None = None,
        body: str | None = None,
        bodyType: Literal["text", "html"] = "text",
        location: str | None = None,
    ) -> CalendarUpdateEventResult:
        runtime = runtime_provider.get()
        return await runtime.graph.update_event(
            mailbox=mailbox,
            eventId=eventId,
            subject=subject,
            start=start,
            end=end,
            timeZone=timeZone,
            attendees=attendees,
            body=body,
            bodyType=bodyType,
            location=location,
        )

    @mcp.tool(
        name="calendar_delete_event",
        description="Delete an event from the default calendar by event ID.",
    )
    async def calendar_delete_event(
        eventId: str,
        mailbox: str | None = None,
    ) -> CalendarDeleteEventResult:
        runtime = runtime_provider.get()
        return await runtime.graph.delete_event(mailbox=mailbox, eventId=eventId)

    @mcp.tool(
        name="sharepoint_search_items",
        description=(
            "Search across everything the signed-in user can access (all "
            "SharePoint sites and OneDrive) for files and folders by name or "
            "content. Best starting point for 'find this file/folder anywhere'. "
            "Optionally filter to file extensions like ['xlsx', 'pdf']."
        ),
    )
    async def sharepoint_search_items(
        query: str,
        top: int = 25,
        extensions: list[str] | None = None,
    ) -> DriveItemsResult:
        runtime = runtime_provider.get()
        return await runtime.sharepoint.search_items(
            query=query,
            top=top,
            extensions=extensions,
        )

    @mcp.tool(
        name="sharepoint_search_sites",
        description="Find SharePoint sites by keyword in their name or title.",
    )
    async def sharepoint_search_sites(
        query: str,
        top: int = 25,
    ) -> SitesResult:
        runtime = runtime_provider.get()
        return await runtime.sharepoint.search_sites(query=query, top=top)

    @mcp.tool(
        name="sharepoint_get_site",
        description=(
            "Resolve a known SharePoint site by hostname and site path, for "
            "example hostname 'contoso.sharepoint.com' and sitePath "
            "'Acquisitions'. Returns the site ID for listing its libraries."
        ),
    )
    async def sharepoint_get_site(
        hostname: str,
        sitePath: str,
    ) -> SiteInfo:
        runtime = runtime_provider.get()
        return await runtime.sharepoint.get_site_by_path(
            hostname=hostname,
            sitePath=sitePath,
        )

    @mcp.tool(
        name="sharepoint_list_drives",
        description=(
            "List a SharePoint site's document libraries. Each library is a "
            "drive; use its driveId to browse contents."
        ),
    )
    async def sharepoint_list_drives(
        siteId: str,
    ) -> DrivesResult:
        runtime = runtime_provider.get()
        return await runtime.sharepoint.list_drives(siteId=siteId)

    @mcp.tool(
        name="sharepoint_list_children",
        description=(
            "List the contents of a folder in a document library. Target the "
            "folder by itemId or by path relative to the drive root; omit both "
            "for the root. Optionally filter to file extensions or folders only."
        ),
    )
    async def sharepoint_list_children(
        driveId: str,
        itemId: str | None = None,
        path: str | None = None,
        top: int = 200,
        extensions: list[str] | None = None,
        foldersOnly: bool = False,
    ) -> DriveItemsResult:
        runtime = runtime_provider.get()
        return await runtime.sharepoint.list_children(
            driveId=driveId,
            itemId=itemId,
            path=path,
            top=top,
            extensions=extensions,
            foldersOnly=foldersOnly,
        )

    @mcp.tool(
        name="sharepoint_search_in_drive",
        description=(
            "Search for files and folders by name within a single document "
            "library identified by driveId."
        ),
    )
    async def sharepoint_search_in_drive(
        driveId: str,
        query: str,
        top: int = 50,
        extensions: list[str] | None = None,
    ) -> DriveItemsResult:
        runtime = runtime_provider.get()
        return await runtime.sharepoint.search_in_drive(
            driveId=driveId,
            query=query,
            top=top,
            extensions=extensions,
        )

    @mcp.tool(
        name="sharepoint_get_item_by_url",
        description=(
            "Resolve a SharePoint or OneDrive sharing/browser URL to a drive "
            "item, returning its driveId and itemId for further browsing."
        ),
    )
    async def sharepoint_get_item_by_url(
        shareUrl: str,
    ) -> DriveItemInfo:
        runtime = runtime_provider.get()
        return await runtime.sharepoint.get_item_by_share_url(shareUrl=shareUrl)

    @mcp.tool(
        name="workbook_resolve",
        description=(
            "Resolve an Excel workbook to a driveId + itemId pair for the other "
            "workbook tools. Provide a shareUrl, or driveId + itemId, or "
            "driveId + itemPath (path relative to the drive root). The returned "
            "driveId and itemId can be reused on later workbook calls."
        ),
    )
    async def workbook_resolve(
        shareUrl: str | None = None,
        driveId: str | None = None,
        itemId: str | None = None,
        itemPath: str | None = None,
    ) -> WorkbookItemRef:
        runtime = runtime_provider.get()
        return await runtime.excel.resolve_workbook(
            shareUrl=shareUrl,
            driveId=driveId,
            itemId=itemId,
            itemPath=itemPath,
        )

    @mcp.tool(
        name="workbook_list_worksheets",
        description="List the worksheets (tabs) in an Excel workbook.",
    )
    async def workbook_list_worksheets(
        driveId: str,
        itemId: str,
        sessionId: str | None = None,
    ) -> WorkbookListWorksheetsResult:
        runtime = runtime_provider.get()
        return await runtime.excel.list_worksheets(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            sessionId=sessionId,
        )

    @mcp.tool(
        name="workbook_list_tables",
        description=(
            "List Excel tables in a workbook, or in a single worksheet when "
            "worksheet is given."
        ),
    )
    async def workbook_list_tables(
        driveId: str,
        itemId: str,
        worksheet: str | None = None,
        sessionId: str | None = None,
    ) -> WorkbookListTablesResult:
        runtime = runtime_provider.get()
        return await runtime.excel.list_tables(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            worksheet=worksheet,
            sessionId=sessionId,
        )

    @mcp.tool(
        name="workbook_get_range",
        description=(
            "Read a fixed range like 'A1:O5' from a worksheet. Returns raw "
            "values, display text, cell formulas, number formats, and the "
            "resolved address."
        ),
    )
    async def workbook_get_range(
        driveId: str,
        itemId: str,
        worksheet: str,
        address: str,
        sessionId: str | None = None,
    ) -> WorkbookRangeResult:
        runtime = runtime_provider.get()
        return await runtime.excel.get_range(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            worksheet=worksheet,
            address=address,
            sessionId=sessionId,
        )

    @mcp.tool(
        name="workbook_get_used_range",
        description=(
            "Read the used (non-empty) range of a worksheet. Use this to "
            "discover the data extent before reading or writing specific cells."
        ),
    )
    async def workbook_get_used_range(
        driveId: str,
        itemId: str,
        worksheet: str,
        valuesOnly: bool = True,
        sessionId: str | None = None,
    ) -> WorkbookRangeResult:
        runtime = runtime_provider.get()
        return await runtime.excel.get_used_range(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            worksheet=worksheet,
            valuesOnly=valuesOnly,
            sessionId=sessionId,
        )

    @mcp.tool(
        name="workbook_update_range",
        description=(
            "Write values, formulas, and/or number formats into a fixed range "
            "in place. The shape of values/formulas/numberFormat must match the "
            "address dimensions. formulas cells may be literal values or formula "
            "strings like ='Unit Mix'!H11 (cross-sheet references are fine). "
            "This edits the stored file via Excel's engine, preserving "
            "formatting and validation."
        ),
    )
    async def workbook_update_range(
        driveId: str,
        itemId: str,
        worksheet: str,
        address: str,
        values: list[list[Any]] | None = None,
        formulas: list[list[Any]] | None = None,
        numberFormat: list[list[Any]] | None = None,
        sessionId: str | None = None,
    ) -> WorkbookWriteResult:
        runtime = runtime_provider.get()
        return await runtime.excel.update_range(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            worksheet=worksheet,
            address=address,
            values=values,
            formulas=formulas,
            numberFormat=numberFormat,
            sessionId=sessionId,
        )

    @mcp.tool(
        name="workbook_add_table_row",
        description=(
            "Append one or more rows to an Excel table. values is a 2D array; "
            "each inner list is one row and must match the table's column count "
            "and order. index=None appends at the end; index=0 inserts at top."
        ),
    )
    async def workbook_add_table_row(
        driveId: str,
        itemId: str,
        table: str,
        values: list[list[Any]],
        index: int | None = None,
        sessionId: str | None = None,
    ) -> WorkbookRowAddResult:
        runtime = runtime_provider.get()
        return await runtime.excel.add_table_row(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            table=table,
            values=values,
            index=index,
            sessionId=sessionId,
        )

    @mcp.tool(
        name="workbook_create_session",
        description=(
            "Create a workbook session and return its sessionId. Pass that "
            "sessionId to subsequent workbook calls to batch them consistently. "
            "persistChanges=True writes to the stored file; False is a "
            "scratch/read session."
        ),
    )
    async def workbook_create_session(
        driveId: str,
        itemId: str,
        persistChanges: bool = True,
    ) -> WorkbookSessionResult:
        runtime = runtime_provider.get()
        return await runtime.excel.create_session(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            persistChanges=persistChanges,
        )

    @mcp.tool(
        name="workbook_close_session",
        description="Close a workbook session previously opened with workbook_create_session.",
    )
    async def workbook_close_session(
        driveId: str,
        itemId: str,
        sessionId: str,
    ) -> dict[str, Any]:
        runtime = runtime_provider.get()
        await runtime.excel.close_session(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            sessionId=sessionId,
        )
        return {"closed": True, "sessionId": sessionId}

    @mcp.tool(
        name="workbook_get_ranges",
        description=(
            "Batch-read many ranges in one call. ranges is a list of "
            "{worksheet, address} objects. Each result carries values, text, "
            "formulas, numberFormat, and the resolved address, in input order. "
            "A failed individual range surfaces its error without failing the "
            "rest. Auto-chunked to <=20 sub-requests per Graph batch."
        ),
    )
    async def workbook_get_ranges(
        driveId: str,
        itemId: str,
        ranges: list[dict[str, str]],
        sessionId: str | None = None,
    ) -> WorkbookRangesResult:
        runtime = runtime_provider.get()
        return await runtime.excel.get_ranges(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            ranges=ranges,
            sessionId=sessionId,
        )

    @mcp.tool(
        name="workbook_update_ranges",
        description=(
            "Batch-write many ranges in one call. updates is a list of objects, "
            "each with worksheet and address plus any of formulas, values, "
            "numberFormat. formulas cells may be literal values or formula "
            "strings like ='Unit Mix'!H11 (cross-sheet references are written "
            "verbatim). Results preserve input order; a failed individual write "
            "surfaces its error without failing the rest. Auto-chunked to <=20 "
            "sub-requests per Graph batch."
        ),
    )
    async def workbook_update_ranges(
        driveId: str,
        itemId: str,
        updates: list[dict[str, Any]],
        sessionId: str | None = None,
    ) -> WorkbookRangesResult:
        runtime = runtime_provider.get()
        return await runtime.excel.update_ranges(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            updates=updates,
            sessionId=sessionId,
        )

    @mcp.tool(
        name="workbook_calculate",
        description=(
            "Force a recalculation of the workbook so computed cells are current "
            "before reading them back. calculationType is one of Recalculate, "
            "Full (default), or FullRebuild. Pass a sessionId to recalc inside "
            "an open session."
        ),
    )
    async def workbook_calculate(
        driveId: str,
        itemId: str,
        calculationType: str = "Full",
        sessionId: str | None = None,
    ) -> WorkbookCalculateResult:
        runtime = runtime_provider.get()
        return await runtime.excel.calculate(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            calculationType=calculationType,
            sessionId=sessionId,
        )

    @mcp.tool(
        name="workbook_list_names",
        description=(
            "List defined names. Workbook-scoped names when worksheet is "
            "omitted; worksheet-scoped names when worksheet is given. Each name "
            "includes its 'refers to' value/formula and scope."
        ),
    )
    async def workbook_list_names(
        driveId: str,
        itemId: str,
        worksheet: str | None = None,
        sessionId: str | None = None,
    ) -> WorkbookNamesResult:
        runtime = runtime_provider.get()
        return await runtime.excel.list_names(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            worksheet=worksheet,
            sessionId=sessionId,
        )

    @mcp.tool(
        name="workbook_get_name_range",
        description=(
            "Resolve a defined name to its range and read it. Returns the "
            "resolved address plus values, text, formulas, and numberFormat. "
            "Provide worksheet for a worksheet-scoped name; omit it for a "
            "workbook-scoped one."
        ),
    )
    async def workbook_get_name_range(
        driveId: str,
        itemId: str,
        name: str,
        worksheet: str | None = None,
        sessionId: str | None = None,
    ) -> WorkbookNameRangeResult:
        runtime = runtime_provider.get()
        return await runtime.excel.get_name_range(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            name=name,
            worksheet=worksheet,
            sessionId=sessionId,
        )

    @mcp.tool(
        name="workbook_clear_range",
        description=(
            "Clear a range in place. applyTo is Contents (values/formulas "
            "only, default), Formats, or All. This edits the stored file via "
            "Excel's engine; it never deletes the workbook file."
        ),
    )
    async def workbook_clear_range(
        driveId: str,
        itemId: str,
        worksheet: str,
        address: str,
        applyTo: str = "Contents",
        sessionId: str | None = None,
    ) -> WorkbookClearResult:
        runtime = runtime_provider.get()
        return await runtime.excel.clear_range(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            worksheet=worksheet,
            address=address,
            applyTo=applyTo,
            sessionId=sessionId,
        )

    @mcp.tool(
        name="workbook_copy_range",
        description=(
            "Copy into address (the destination range) from sourceRange (e.g. "
            "'Unit Mix'!A1:B5 for a cross-sheet source). copyType is one of All "
            "(default), Formulas, Values, or Formats."
        ),
    )
    async def workbook_copy_range(
        driveId: str,
        itemId: str,
        worksheet: str,
        address: str,
        sourceRange: str,
        copyType: str = "All",
        sessionId: str | None = None,
    ) -> WorkbookCopyResult:
        runtime = runtime_provider.get()
        return await runtime.excel.copy_range(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            worksheet=worksheet,
            address=address,
            sourceRange=sourceRange,
            copyType=copyType,
            sessionId=sessionId,
        )

    @mcp.tool(
        name="workbook_insert_range",
        description=(
            "Insert blank cells at address, shifting existing cells. shift is "
            "Down (default) or Right. Useful for insert-at-top patterns."
        ),
    )
    async def workbook_insert_range(
        driveId: str,
        itemId: str,
        worksheet: str,
        address: str,
        shift: str = "Down",
        sessionId: str | None = None,
    ) -> WorkbookInsertResult:
        runtime = runtime_provider.get()
        return await runtime.excel.insert_range(
            WorkbookItemRef(driveId=driveId, itemId=itemId),
            worksheet=worksheet,
            address=address,
            shift=shift,
            sessionId=sessionId,
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
