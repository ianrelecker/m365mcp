from __future__ import annotations

import base64
import io
from contextlib import asynccontextmanager
from datetime import UTC, datetime, timedelta
from typing import Any
from urllib.parse import quote

import httpx

try:  # pragma: no cover - exercised through dependency-aware tests
    from pypdf import PdfReader
except Exception:  # pragma: no cover - pypdf is an optional import at module load
    PdfReader = None  # type: ignore[assignment]

try:  # pragma: no cover - page rendering degrades to a reason string without these
    import pypdfium2
except Exception:  # pragma: no cover
    pypdfium2 = None  # type: ignore[assignment]

try:  # pragma: no cover
    from PIL import Image as PILImage
except Exception:  # pragma: no cover
    PILImage = None  # type: ignore[assignment]

from .models import (
    AttachmentImage,
    AttachmentInfo,
    CalendarAttendee,
    CalendarCreateEventResult,
    CalendarDateTime,
    CalendarDeleteEventResult,
    CalendarEvent,
    CalendarListEventsResult,
    CalendarUpdateEventResult,
    CalendarWindow,
    ContactAddress,
    ContactFoldersResult,
    ContactFolderInfo,
    ContactGetResult,
    ContactInfo,
    ContactMutationResult,
    ContactsListResult,
    ContactsSearchResult,
    FullMessage,
    MailAttachmentContentResult,
    MailAttachmentImageResult,
    MailAttachmentPdfResult,
    MailCategoryInfo,
    MailCategoryResult,
    MailCheckInboxResult,
    MailCreateDraftResult,
    MailFolderMutationResult,
    MailFolderInfo,
    MailFolderTreeNode,
    MailFolderTreeResult,
    MailGetResult,
    MailInlineImagesResult,
    MailListAttachmentsResult,
    MailListCategoriesResult,
    MailListDraftsResult,
    MailListFoldersResult,
    MailListResult,
    MailListRulesResult,
    MailMoveResult,
    MailResolveFolderResult,
    MailRuleInfo,
    MailRuleResult,
    MailSearchResult,
    MailSendDraftResult,
    MailSendResult,
    MailThreadResult,
    MailUpdateMessageResult,
    MessageBody,
    MessageSummary,
    PdfPageImage,
    SkippedAttachment,
)
from .microsoft_auth import MicrosoftAuthService
from .pid_policy import PidPolicy


def _utc_now_iso() -> str:
    return datetime.now(UTC).isoformat().replace("+00:00", "Z")


class _PdfRenderError(RuntimeError):
    """Raised when a PDF cannot be rasterized; surfaced as unsupportedReason."""


MESSAGE_SUMMARY_SELECT = (
    "id,subject,from,sender,replyTo,receivedDateTime,sentDateTime,"
    "bodyPreview,webLink,isDraft,isRead,hasAttachments,importance,"
    "categories,flag,inferenceClassification,parentFolderId,internetMessageId,"
    "conversationId"
)
MESSAGE_FULL_SELECT = (
    "id,subject,from,sender,replyTo,toRecipients,ccRecipients,bccRecipients,"
    "receivedDateTime,sentDateTime,bodyPreview,body,webLink,isDraft,isRead,"
    "hasAttachments,importance,categories,flag,inferenceClassification,"
    "parentFolderId,internetMessageId,conversationId"
)
MAIL_FOLDER_SELECT = (
    "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount,isHidden"
)
CONTACT_SELECT = (
    "id,displayName,givenName,surname,companyName,jobTitle,"
    "personalNotes,businessPhones,mobilePhone,emailAddresses,categories,"
    "businessAddress,homeAddress,otherAddress,parentFolderId"
)
PERSONAL_HOME_PAGE_EXTENDED_PROPERTY_ID = "String 0x3A50"
CONTACT_EXTENDED_PROPERTIES_EXPAND = (
    "singleValueExtendedProperties($filter=id eq "
    f"'{PERSONAL_HOME_PAGE_EXTENDED_PROPERTY_ID}')"
)
CONTACT_FOLDER_SELECT = "id,displayName,parentFolderId"


def _build_contact_read_query(*, top: int | None = None) -> str:
    """Build a query string for contact reads that preserves the literal
    ``$filter=...`` inside ``$expand`` without URL-encoding.

    Microsoft Graph rejects (or silently ignores) the expand clause when the
    ``$`` in ``$filter`` is percent-encoded, so we hand-format the query string
    instead of routing it through ``httpx.QueryParams``. All values here come
    from server-controlled constants (``CONTACT_SELECT``,
    ``CONTACT_EXTENDED_PROPERTIES_EXPAND``) or a bounded integer, so there is
    no user input that needs escaping.
    """

    parts: list[str] = []
    if top is not None:
        parts.append(f"$top={top}")
    parts.append(f"$select={CONTACT_SELECT}")
    parts.append(f"$expand={CONTACT_EXTENDED_PROPERTIES_EXPAND}")
    return "&".join(parts)


MESSAGE_RULE_SELECT = (
    "id,displayName,sequence,isEnabled,hasError,isReadOnly,"
    "conditions,actions,exceptions"
)
SAFE_ATTACHMENT_CONTENT_TYPES = {
    "application/calendar",
    "application/csv",
    "application/json",
    "application/pdf",
    "application/xml",
    "text/calendar",
    "text/csv",
    "text/html",
    "text/markdown",
    "text/plain",
    "text/tab-separated-values",
    "text/xml",
}
# Image types every MCP client can render as image content.
SUPPORTED_IMAGE_CONTENT_TYPES = {
    "image/png",
    "image/jpeg",
    "image/gif",
    "image/webp",
}
IMAGE_EXTENSION_CONTENT_TYPES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".jfif": "image/jpeg",
    ".gif": "image/gif",
    ".webp": "image/webp",
}
IMAGE_CONTENT_TYPE_ALIASES = {
    "image/jpg": "image/jpeg",
    "image/pjpeg": "image/jpeg",
    "image/x-png": "image/png",
}
# Claude's per-image ceiling is 10 MB of base64 on the Claude API (5 MB on
# Bedrock/Vertex); 4 MB of raw bytes encodes to ~5.3 MB, so a single image stays
# under both. Graph rarely returns inline pictures anywhere near this size.
DEFAULT_IMAGE_MAX_BYTES = 4_000_000
# Per-image caps do not bound a whole message: ten 4 MB pictures would encode to
# ~53 MB and blow the 32 MB request limit, so cap the batch too.
DEFAULT_MAX_TOTAL_IMAGE_BYTES = 8_000_000
DEFAULT_MAX_INLINE_IMAGES = 10
# A rendered page is one image, so the long edge targets the same viewing
# resolution as a photo attachment; 1600px keeps body text legible.
DEFAULT_PDF_PAGE_LONG_EDGE = 1600
MAX_PDF_RENDER_SCALE = 4.0
# JPEG rather than PNG: a rendered page compresses to a fraction of the size
# with no meaningful loss of legibility.
PDF_PAGE_JPEG_QUALITY = 80
# Pages are expensive in both bytes and tokens, so default to a handful and let
# the caller page through a longer document.
DEFAULT_MAX_PDF_PAGES = 5
# Source PDFs are downloaded, not viewed, so this is far larger than an image
# cap; the page images it renders to are bounded separately.
DEFAULT_PDF_MAX_BYTES = 25_000_000
SAFE_ATTACHMENT_EXTENSIONS = {
    ".csv",
    ".ics",
    ".json",
    ".log",
    ".md",
    ".pdf",
    ".txt",
    ".tsv",
    ".xml",
    ".html",
    ".htm",
}


class MicrosoftGraphClient:
    def __init__(
        self,
        auth_service: MicrosoftAuthService,
        http_client: httpx.AsyncClient | None = None,
        pid_policy: PidPolicy | None = None,
    ) -> None:
        self._auth_service = auth_service
        self._http_client = http_client
        self._pid_policy = pid_policy or PidPolicy.disabled()

    @asynccontextmanager
    async def _client(self) -> Any:
        if self._http_client is not None:
            yield self._http_client
            return

        async with httpx.AsyncClient(follow_redirects=False, timeout=30.0) as client:
            yield client

    async def list_messages(
        self,
        *,
        mailbox: str | None = None,
        folder: str = "Inbox",
        folderId: str | None = None,
        folderPath: str | None = None,
        top: int = 25,
        isRead: bool | None = None,
        hasAttachments: bool | None = None,
        importance: str | None = None,
        categories: list[str] | None = None,
        flagStatus: str | None = None,
        inferenceClassification: str | None = None,
    ) -> MailListResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        resolved_folder_id = folderId
        resolved_folder_path = folderPath
        legacy_folder = folder

        if folderPath is None and folderId is None and "/" in folder:
            resolved_folder_path = folder

        if resolved_folder_path and not resolved_folder_id:
            resolved = await self._resolve_mail_folder_by_path(base, resolved_folder_path)
            resolved_folder_id = resolved.id

        filters = self._message_filters(
            isRead=isRead,
            hasAttachments=hasAttachments,
            importance=importance,
            categories=categories,
            flagStatus=flagStatus,
            inferenceClassification=inferenceClassification,
        )
        params: dict[str, str] = {
            "$top": str(min(top, 100)),
            "$select": MESSAGE_SUMMARY_SELECT,
        }
        if filters:
            params["$filter"] = " and ".join(filters)
        query = httpx.QueryParams(params)
        result = await self._request(
            f"{self._mail_folder_messages_path(base, legacy_folder, resolved_folder_id)}?{query}"
        )

        return MailListResult(
            mailbox=normalized_mailbox or "me",
            folder=folder,
            folderId=resolved_folder_id,
            folderPath=resolved_folder_path,
            messages=[self._map_message_summary(message) for message in result["value"]],
        )

    async def check_inbox(
        self,
        *,
        mailbox: str | None = None,
        folderPath: str = "Inbox",
        folderId: str | None = None,
        top: int = 25,
        includeRead: bool = False,
        inferenceClassification: str | None = None,
    ) -> MailCheckInboxResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        folder_info = (
            await self._get_mail_folder(base, folderId)
            if folderId
            else await self._resolve_mail_folder_by_path(base, folderPath)
        )
        listed = await self.list_messages(
            mailbox=normalized_mailbox,
            folder=folderPath,
            folderId=folder_info.id,
            folderPath=folderPath,
            top=top,
            isRead=None if includeRead else False,
            inferenceClassification=inferenceClassification,
        )
        return MailCheckInboxResult(
            mailbox=normalized_mailbox or "me",
            folder=folder_info,
            messages=listed.messages,
        )

    async def search_messages(
        self,
        *,
        mailbox: str | None = None,
        query: str,
        top: int = 10,
    ) -> MailSearchResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        escaped_query = query.replace('"', '\\"')
        params = httpx.QueryParams(
            {
                "$top": str(min(top, 50)),
                "$search": f'"{escaped_query}"',
                "$select": MESSAGE_SUMMARY_SELECT,
            }
        )
        result = await self._request(
            f"{base}/messages?{params}",
            headers={"ConsistencyLevel": "eventual"},
        )

        return MailSearchResult(
            mailbox=normalized_mailbox or "me",
            query=query,
            messages=[self._map_message_summary(message) for message in result["value"]],
        )

    async def get_message(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
    ) -> MailGetResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        params = httpx.QueryParams(
            {
                "$select": MESSAGE_FULL_SELECT,
            }
        )
        message = await self._request(
            f"{base}/messages/{quote(messageId, safe='')}?{params}"
        )
        return MailGetResult(
            mailbox=normalized_mailbox or "me",
            message=self._map_full_message(message),
        )

    async def list_drafts(
        self,
        *,
        mailbox: str | None = None,
        top: int = 25,
    ) -> MailListDraftsResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        params = httpx.QueryParams(
            {
                "$top": str(min(top, 100)),
                "$select": MESSAGE_SUMMARY_SELECT,
            }
        )
        result = await self._request(f"{base}/mailFolders('Drafts')/messages?{params}")
        return MailListDraftsResult(
            mailbox=normalized_mailbox or "me",
            drafts=[self._map_message_summary(message) for message in result["value"]],
        )

    async def create_draft(
        self,
        *,
        mailbox: str | None = None,
        subject: str,
        to: list[str],
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        body: str,
        bodyType: str = "text",
        from_: str | None = None,
    ) -> MailCreateDraftResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)

        message = await self._request(
            f"{base}/messages",
            method="POST",
            json_body={
                "subject": subject,
                "body": {
                    "contentType": self._graph_body_content_type(bodyType),
                    "content": body,
                },
                "toRecipients": self._to_recipients(to),
                "ccRecipients": self._to_recipients(cc),
                "bccRecipients": self._to_recipients(bcc),
                "from": (
                    {"emailAddress": {"address": from_}}
                    if from_
                    else (
                        {"emailAddress": {"address": normalized_mailbox}}
                        if normalized_mailbox
                        else None
                    )
                ),
            },
        )

        return MailCreateDraftResult(
            mailbox=normalized_mailbox or "me",
            draft=self._map_message_summary(message),
        )

    async def send_mail(
        self,
        *,
        mailbox: str | None = None,
        subject: str,
        to: list[str],
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        body: str,
        bodyType: str = "text",
        from_: str | None = None,
        saveToSentItems: bool = True,
    ) -> MailSendResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        await self._request(
            f"{base}/sendMail",
            method="POST",
            json_body={
                "message": {
                    "subject": subject,
                    "body": {
                        "contentType": self._graph_body_content_type(bodyType),
                        "content": body,
                    },
                    "toRecipients": self._to_recipients(to),
                    "ccRecipients": self._to_recipients(cc),
                    "bccRecipients": self._to_recipients(bcc),
                    "from": (
                        {"emailAddress": {"address": from_}} if from_ else None
                    ),
                },
                "saveToSentItems": saveToSentItems,
            },
        )

        return MailSendResult(
            mailbox=normalized_mailbox or "me",
            subject=subject,
            sent=True,
        )

    async def send_draft(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
    ) -> MailSendDraftResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        await self._request(
            f"{base}/messages/{quote(messageId, safe='')}/send",
            method="POST",
        )

        return MailSendDraftResult(
            mailbox=normalized_mailbox or "me",
            messageId=messageId,
            sent=True,
        )

    async def move_message(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
        destinationFolder: str,
        destinationFolderIsId: bool = False,
        destinationFolderId: str | None = None,
        destinationFolderPath: str | None = None,
    ) -> MailMoveResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        if destinationFolderId:
            destination_id = destinationFolderId
        elif destinationFolderPath:
            destination_id = (
                await self._resolve_mail_folder_by_path(base, destinationFolderPath)
            ).id
        elif destinationFolderIsId:
            destination_id = destinationFolder
        elif "/" in destinationFolder:
            destinationFolderPath = destinationFolder
            destination_id = (
                await self._resolve_mail_folder_by_path(base, destinationFolder)
            ).id
        else:
            destination_id = await self._resolve_folder_id(base, destinationFolder)
        moved = await self._request(
            f"{base}/messages/{quote(messageId, safe='')}/move",
            method="POST",
            json_body={"destinationId": destination_id},
        )

        return MailMoveResult(
            mailbox=normalized_mailbox or "me",
            destinationFolder=destinationFolder,
            destinationFolderId=destination_id,
            destinationFolderPath=destinationFolderPath,
            movedMessage=self._map_message_summary(moved),
        )

    async def list_mail_folders(
        self,
        *,
        mailbox: str | None = None,
        parentFolderId: str | None = None,
        top: int = 100,
    ) -> MailListFoldersResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        folders = await self._list_mail_folder_infos(
            base,
            parentFolderId=parentFolderId,
            top=top,
        )
        return MailListFoldersResult(
            mailbox=normalized_mailbox or "me",
            parentFolderId=parentFolderId,
            folders=folders,
        )

    async def mail_folder_tree(
        self,
        *,
        mailbox: str | None = None,
        rootFolderId: str | None = None,
        maxDepth: int = 4,
    ) -> MailFolderTreeResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        depth = max(1, min(maxDepth, 8))
        roots = await self._list_mail_folder_tree(
            base,
            parentFolderId=rootFolderId,
            maxDepth=depth,
            currentPath=None,
        )
        return MailFolderTreeResult(
            mailbox=normalized_mailbox or "me",
            rootFolderId=rootFolderId,
            maxDepth=depth,
            folders=roots,
        )

    async def resolve_mail_folder(
        self,
        *,
        mailbox: str | None = None,
        folderPath: str | None = None,
        parentFolderId: str | None = None,
        displayName: str | None = None,
    ) -> MailResolveFolderResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        if folderPath:
            folder = await self._resolve_mail_folder_by_path(base, folderPath)
        elif displayName:
            folder = await self._find_mail_folder_child(
                base,
                parentFolderId=parentFolderId,
                displayName=displayName,
            )
        elif parentFolderId:
            folder = await self._get_mail_folder(base, parentFolderId)
        else:
            raise ValueError("Provide folderPath, displayName, or parentFolderId")

        return MailResolveFolderResult(
            mailbox=normalized_mailbox or "me",
            folder=folder,
        )

    async def create_mail_folder(
        self,
        *,
        mailbox: str | None = None,
        displayName: str,
        parentFolderId: str | None = None,
        parentFolderPath: str | None = None,
        isHidden: bool | None = None,
    ) -> MailFolderMutationResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        resolved_parent_id = parentFolderId
        if parentFolderPath and not resolved_parent_id:
            resolved_parent_id = (
                await self._resolve_mail_folder_by_path(base, parentFolderPath)
            ).id
        path = (
            f"{base}/mailFolders/{quote(resolved_parent_id, safe='')}/childFolders"
            if resolved_parent_id
            else f"{base}/mailFolders"
        )
        folder = await self._request(
            path,
            method="POST",
            json_body={"displayName": displayName, "isHidden": isHidden},
        )
        return MailFolderMutationResult(
            mailbox=normalized_mailbox or "me",
            folder=self._map_mail_folder(folder),
        )

    async def rename_mail_folder(
        self,
        *,
        mailbox: str | None = None,
        displayName: str,
        folderId: str | None = None,
        folderPath: str | None = None,
    ) -> MailFolderMutationResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        resolved_folder_id = await self._resolve_mail_folder_identifier(
            base,
            folderId=folderId,
            folderPath=folderPath,
        )
        folder = await self._request(
            f"{base}/mailFolders/{quote(resolved_folder_id, safe='')}",
            method="PATCH",
            json_body={"displayName": displayName},
        )
        return MailFolderMutationResult(
            mailbox=normalized_mailbox or "me",
            folder=self._map_mail_folder(folder),
        )

    async def delete_mail_folder(
        self,
        *,
        mailbox: str | None = None,
        folderId: str | None = None,
        folderPath: str | None = None,
    ) -> MailFolderMutationResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        resolved_folder_id = await self._resolve_mail_folder_identifier(
            base,
            folderId=folderId,
            folderPath=folderPath,
        )
        await self._request(
            f"{base}/mailFolders/{quote(resolved_folder_id, safe='')}",
            method="DELETE",
        )
        return MailFolderMutationResult(
            mailbox=normalized_mailbox or "me",
            folderId=resolved_folder_id,
            deleted=True,
        )

    async def list_mail_rules(
        self,
        *,
        mailbox: str | None = None,
        top: int = 100,
    ) -> MailListRulesResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        params = httpx.QueryParams(
            {"$top": str(min(top, 100)), "$select": MESSAGE_RULE_SELECT}
        )
        result = await self._request(f"{self._mail_rules_path(base)}?{params}")
        return MailListRulesResult(
            mailbox=normalized_mailbox or "me",
            rules=[self._map_rule(rule) for rule in result.get("value", [])],
        )

    async def create_mail_rule(
        self,
        *,
        mailbox: str | None = None,
        displayName: str,
        sequence: int = 1,
        isEnabled: bool = True,
        conditions: dict[str, Any] | None = None,
        actions: dict[str, Any] | None = None,
        exceptions: dict[str, Any] | None = None,
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
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        payload = await self._mail_rule_payload(
            base,
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
        rule = await self._request(
            self._mail_rules_path(base),
            method="POST",
            json_body=payload,
        )
        return MailRuleResult(
            mailbox=normalized_mailbox or "me",
            rule=self._map_rule(rule),
        )

    async def update_mail_rule(
        self,
        *,
        mailbox: str | None = None,
        ruleId: str,
        displayName: str | None = None,
        sequence: int | None = None,
        isEnabled: bool | None = None,
        conditions: dict[str, Any] | None = None,
        actions: dict[str, Any] | None = None,
        exceptions: dict[str, Any] | None = None,
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
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        payload = await self._mail_rule_payload(
            base,
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
        if not payload:
            raise ValueError("Provide at least one rule field to update")
        rule = await self._request(
            f"{self._mail_rules_path(base)}/{quote(ruleId, safe='')}",
            method="PATCH",
            json_body=payload,
        )
        return MailRuleResult(
            mailbox=normalized_mailbox or "me",
            rule=self._map_rule(rule),
        )

    async def delete_mail_rule(
        self,
        *,
        mailbox: str | None = None,
        ruleId: str,
    ) -> MailRuleResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        await self._request(
            f"{self._mail_rules_path(base)}/{quote(ruleId, safe='')}",
            method="DELETE",
        )
        return MailRuleResult(
            mailbox=normalized_mailbox or "me",
            ruleId=ruleId,
            deleted=True,
        )

    async def list_attachments(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
        includeInline: bool = False,
    ) -> MailListAttachmentsResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        result = await self._request(
            f"{base}/messages/{quote(messageId, safe='')}/attachments"
        )
        attachments = [
            self._map_attachment(attachment)
            for attachment in result.get("value", [])
        ]
        if not includeInline:
            attachments = [
                attachment for attachment in attachments if not attachment.isInline
            ]
        return MailListAttachmentsResult(
            mailbox=normalized_mailbox or "me",
            messageId=messageId,
            attachments=attachments,
        )

    async def get_attachment_content(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
        attachmentId: str,
        maxBytes: int = 1_000_000,
        maxChars: int = 100_000,
    ) -> MailAttachmentContentResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        metadata = await self._request(
            f"{base}/messages/{quote(messageId, safe='')}/attachments/{quote(attachmentId, safe='')}"
        )
        attachment = self._map_attachment(metadata)
        unsupported = self._attachment_unsupported_reason(
            attachment,
            maxBytes=maxBytes,
        )
        if unsupported:
            return MailAttachmentContentResult(
                mailbox=normalized_mailbox or "me",
                messageId=messageId,
                attachment=attachment,
                unsupportedReason=unsupported,
            )

        content_bytes = await self._attachment_bytes(
            base,
            messageId=messageId,
            attachment=metadata,
        )

        if len(content_bytes) > maxBytes:
            return MailAttachmentContentResult(
                mailbox=normalized_mailbox or "me",
                messageId=messageId,
                attachment=attachment,
                truncated=True,
                unsupportedReason=f"Attachment content exceeds maxBytes={maxBytes}",
            )

        if self._is_pdf_attachment(attachment):
            content = self._extract_pdf_text(content_bytes)
            if not content.strip():
                return MailAttachmentContentResult(
                    mailbox=normalized_mailbox or "me",
                    messageId=messageId,
                    attachment=attachment,
                    encoding="pdf-text",
                    unsupportedReason=(
                        "PDF did not contain extractable text; scanned PDFs need OCR"
                    ),
                )
            truncated = len(content) > maxChars
            return MailAttachmentContentResult(
                mailbox=normalized_mailbox or "me",
                messageId=messageId,
                attachment=attachment,
                content=self._pid_policy.redact_text(content[:maxChars]),
                encoding="pdf-text",
                truncated=truncated,
                unsupportedReason=(
                    f"Extracted PDF text exceeds maxChars={maxChars}"
                    if truncated
                    else None
                ),
            )

        return MailAttachmentContentResult(
            mailbox=normalized_mailbox or "me",
            messageId=messageId,
            attachment=attachment,
            content=self._pid_policy.redact_text(
                content_bytes.decode("utf-8", errors="replace")
            ),
            encoding="utf-8",
        )

    async def get_attachment_image(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
        attachmentId: str,
        maxBytes: int = DEFAULT_IMAGE_MAX_BYTES,
    ) -> MailAttachmentImageResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        metadata = await self._request(
            f"{base}/messages/{quote(messageId, safe='')}/attachments/{quote(attachmentId, safe='')}"
        )
        attachment = self._map_attachment(metadata)
        unsupported = self._attachment_image_unsupported_reason(
            attachment,
            maxBytes=maxBytes,
        )
        if unsupported:
            return MailAttachmentImageResult(
                mailbox=normalized_mailbox or "me",
                messageId=messageId,
                attachment=attachment,
                unsupportedReason=unsupported,
            )

        content_bytes = await self._attachment_bytes(
            base,
            messageId=messageId,
            attachment=metadata,
        )
        if len(content_bytes) > maxBytes:
            return MailAttachmentImageResult(
                mailbox=normalized_mailbox or "me",
                messageId=messageId,
                attachment=attachment,
                unsupportedReason=f"Image content exceeds maxBytes={maxBytes}",
            )

        return MailAttachmentImageResult(
            mailbox=normalized_mailbox or "me",
            messageId=messageId,
            attachment=attachment,
            image=self._build_attachment_image(attachment, content_bytes),
        )

    async def get_attachment_pdf_pages(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
        attachmentId: str,
        firstPage: int = 1,
        maxPages: int = DEFAULT_MAX_PDF_PAGES,
        longEdge: int = DEFAULT_PDF_PAGE_LONG_EDGE,
        maxBytes: int = DEFAULT_PDF_MAX_BYTES,
        maxTotalBytes: int = DEFAULT_MAX_TOTAL_IMAGE_BYTES,
    ) -> MailAttachmentPdfResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        metadata = await self._request(
            f"{base}/messages/{quote(messageId, safe='')}/attachments/{quote(attachmentId, safe='')}"
        )
        attachment = self._map_attachment(metadata)

        def unsupported(reason: str) -> MailAttachmentPdfResult:
            return MailAttachmentPdfResult(
                mailbox=normalized_mailbox or "me",
                messageId=messageId,
                attachment=attachment,
                unsupportedReason=reason,
            )

        if attachment.attachmentType not in (None, "#microsoft.graph.fileAttachment"):
            return unsupported(
                "Only file attachments can be rendered as pages in this MCP server"
            )
        if not self._is_pdf_attachment(attachment):
            return unsupported("Attachment is not a PDF")
        if pypdfium2 is None or PILImage is None:
            return unsupported(
                "PDF page rendering requires the pypdfium2 and pillow packages"
            )
        if attachment.size is not None and attachment.size > maxBytes:
            return unsupported(f"PDF size exceeds maxBytes={maxBytes}")

        content_bytes = await self._attachment_bytes(
            base,
            messageId=messageId,
            attachment=metadata,
        )
        if len(content_bytes) > maxBytes:
            return unsupported(f"PDF size exceeds maxBytes={maxBytes}")

        try:
            page_count, pages, truncated = self._render_pdf_pages(
                content_bytes,
                firstPage=firstPage,
                maxPages=maxPages,
                longEdge=longEdge,
                maxTotalBytes=maxTotalBytes,
            )
        except _PdfRenderError as error:
            return unsupported(str(error))

        return MailAttachmentPdfResult(
            mailbox=normalized_mailbox or "me",
            messageId=messageId,
            attachment=attachment,
            pageCount=page_count,
            pages=pages,
            truncated=truncated,
            unsupportedReason=(
                f"PDF has no page {firstPage}; it has {page_count} pages"
                if not pages and page_count
                else None
            ),
        )

    async def get_inline_images(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
        maxBytes: int = DEFAULT_IMAGE_MAX_BYTES,
        maxTotalBytes: int = DEFAULT_MAX_TOTAL_IMAGE_BYTES,
        maxImages: int = DEFAULT_MAX_INLINE_IMAGES,
        includeNonInline: bool = False,
    ) -> MailInlineImagesResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        listed = await self._request(
            f"{base}/messages/{quote(messageId, safe='')}/attachments"
        )

        images: list[AttachmentImage] = []
        skipped: list[SkippedAttachment] = []
        truncated = False
        total_bytes = 0

        for raw in listed.get("value", []):
            attachment = self._map_attachment(raw)
            if not self._is_image_attachment(attachment):
                continue
            if not includeNonInline and not self._is_inline_attachment(attachment):
                continue
            if len(images) >= maxImages:
                truncated = True
                break

            unsupported = self._attachment_image_unsupported_reason(
                attachment,
                maxBytes=maxBytes,
            )
            if unsupported:
                skipped.append(
                    SkippedAttachment(attachment=attachment, reason=unsupported)
                )
                continue

            content_bytes = await self._attachment_bytes(
                base,
                messageId=messageId,
                attachment=raw,
            )
            if len(content_bytes) > maxBytes:
                skipped.append(
                    SkippedAttachment(
                        attachment=attachment,
                        reason=f"Image content exceeds maxBytes={maxBytes}",
                    )
                )
                continue

            if total_bytes + len(content_bytes) > maxTotalBytes:
                truncated = True
                break

            total_bytes += len(content_bytes)
            images.append(self._build_attachment_image(attachment, content_bytes))

        return MailInlineImagesResult(
            mailbox=normalized_mailbox or "me",
            messageId=messageId,
            images=images,
            skipped=skipped,
            truncated=truncated,
        )

    async def get_thread(
        self,
        *,
        mailbox: str | None = None,
        messageId: str | None = None,
        conversationId: str | None = None,
        top: int = 50,
    ) -> MailThreadResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        resolved_conversation_id = conversationId
        if not resolved_conversation_id:
            if not messageId:
                raise ValueError("Provide messageId or conversationId")
            message = await self._request(
                f"{base}/messages/{quote(messageId, safe='')}?$select=conversationId"
            )
            resolved_conversation_id = str(message["conversationId"])

        params = httpx.QueryParams(
            {
                "$top": str(min(top, 100)),
                "$select": MESSAGE_SUMMARY_SELECT,
                "$filter": (
                    "conversationId eq "
                    f"'{self._escape_odata_string(resolved_conversation_id)}'"
                ),
            }
        )
        result = await self._request(f"{base}/messages?{params}")
        messages = [
            self._map_message_summary(message) for message in result["value"]
        ]
        messages.sort(key=lambda message: message.receivedDateTime or "")
        return MailThreadResult(
            mailbox=normalized_mailbox or "me",
            conversationId=resolved_conversation_id,
            messages=messages,
        )

    async def create_reply_draft(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
        comment: str,
        replyAll: bool = False,
        bodyType: str = "html",
    ) -> MailCreateDraftResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        action = "createReplyAll" if replyAll else "createReply"
        message = await self._request(
            f"{base}/messages/{quote(messageId, safe='')}/{action}",
            method="POST",
            json_body={
                "message": {
                    "body": {
                        "contentType": self._graph_body_content_type(bodyType),
                        "content": comment,
                    }
                }
            },
        )
        return MailCreateDraftResult(
            mailbox=normalized_mailbox or "me",
            draft=self._map_message_summary(message),
        )

    async def send_reply(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
        comment: str,
        replyAll: bool = False,
        bodyType: str = "html",
    ) -> MailSendResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        action = "replyAll" if replyAll else "reply"
        await self._request(
            f"{base}/messages/{quote(messageId, safe='')}/{action}",
            method="POST",
            json_body={
                "message": {
                    "body": {
                        "contentType": self._graph_body_content_type(bodyType),
                        "content": comment,
                    }
                }
            },
        )
        return MailSendResult(
            mailbox=normalized_mailbox or "me",
            messageId=messageId,
            replyAll=replyAll,
            sent=True,
        )

    async def list_categories(
        self,
        *,
        mailbox: str | None = None,
    ) -> MailListCategoriesResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        result = await self._request(f"{base}/outlook/masterCategories")
        return MailListCategoriesResult(
            mailbox=normalized_mailbox or "me",
            categories=[
                self._map_category(category) for category in result.get("value", [])
            ],
        )

    async def create_category(
        self,
        *,
        mailbox: str | None = None,
        displayName: str,
        color: str = "preset0",
    ) -> MailCategoryResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        category = await self._request(
            f"{base}/outlook/masterCategories",
            method="POST",
            json_body={"displayName": displayName, "color": color},
        )
        return MailCategoryResult(
            mailbox=normalized_mailbox or "me",
            category=self._map_category(category),
        )

    async def update_category(
        self,
        *,
        mailbox: str | None = None,
        categoryId: str,
        displayName: str | None = None,
        color: str | None = None,
    ) -> MailCategoryResult:
        if displayName is not None:
            raise ValueError(
                "Microsoft Graph does not support renaming existing Outlook "
                "master categories. Create a new category and delete the old "
                "one after confirming it is safe to do so."
            )
        payload = self._omit_none({"color": color})
        if not payload:
            raise ValueError("Provide color to update an Outlook master category.")

        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        category = await self._request(
            f"{base}/outlook/masterCategories/{quote(categoryId, safe='')}",
            method="PATCH",
            json_body=payload,
        )
        return MailCategoryResult(
            mailbox=normalized_mailbox or "me",
            category=self._map_category(category),
        )

    async def delete_category(
        self,
        *,
        mailbox: str | None = None,
        categoryId: str,
    ) -> MailCategoryResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        await self._request(
            f"{base}/outlook/masterCategories/{quote(categoryId, safe='')}",
            method="DELETE",
        )
        return MailCategoryResult(
            mailbox=normalized_mailbox or "me",
            categoryId=categoryId,
            deleted=True,
        )

    async def set_message_categories(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
        categories: list[str],
    ) -> MailUpdateMessageResult:
        return await self._patch_message(
            mailbox=mailbox,
            messageId=messageId,
            payload={"categories": categories},
        )

    async def add_message_categories(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
        categories: list[str],
    ) -> MailUpdateMessageResult:
        current = await self.get_message(mailbox=mailbox, messageId=messageId)
        existing = list(current.message.categories)
        for category in categories:
            if category not in existing:
                existing.append(category)
        return await self.set_message_categories(
            mailbox=mailbox,
            messageId=messageId,
            categories=existing,
        )

    async def remove_message_categories(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
        categories: list[str],
    ) -> MailUpdateMessageResult:
        current = await self.get_message(mailbox=mailbox, messageId=messageId)
        remove = set(categories)
        return await self.set_message_categories(
            mailbox=mailbox,
            messageId=messageId,
            categories=[
                category for category in current.message.categories if category not in remove
            ],
        )

    async def clear_message_categories(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
    ) -> MailUpdateMessageResult:
        return await self.set_message_categories(
            mailbox=mailbox,
            messageId=messageId,
            categories=[],
        )

    async def mark_message_read(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
        isRead: bool = True,
    ) -> MailUpdateMessageResult:
        return await self._patch_message(
            mailbox=mailbox,
            messageId=messageId,
            payload={"isRead": isRead},
        )

    async def set_message_flag(
        self,
        *,
        mailbox: str | None = None,
        messageId: str,
        flagStatus: str,
        startDateTime: str | None = None,
        dueDateTime: str | None = None,
    ) -> MailUpdateMessageResult:
        flag: dict[str, Any] = {"flagStatus": flagStatus}
        if startDateTime:
            flag["startDateTime"] = {"dateTime": startDateTime, "timeZone": "UTC"}
        if dueDateTime:
            flag["dueDateTime"] = {"dateTime": dueDateTime, "timeZone": "UTC"}
        return await self._patch_message(
            mailbox=mailbox,
            messageId=messageId,
            payload={"flag": flag},
        )

    async def list_contacts(
        self,
        *,
        mailbox: str | None = None,
        folderId: str | None = None,
        top: int = 25,
    ) -> ContactsListResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        query_string = _build_contact_read_query(top=min(top, 100))
        result = await self._request(
            f"{self._contacts_path(normalized_mailbox, folderId)}?{query_string}"
        )
        return ContactsListResult(
            mailbox=normalized_mailbox or "me",
            folderId=folderId,
            contacts=[self._map_contact(contact) for contact in result.get("value", [])],
        )

    async def search_contacts(
        self,
        *,
        mailbox: str | None = None,
        query: str,
        folderId: str | None = None,
        top: int = 25,
        maxPages: int = 5,
    ) -> ContactsSearchResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        query_text = query.strip()
        if "@" in query_text and " " not in query_text:
            base_query = _build_contact_read_query(top=min(max(top, 1), 100))
            email_filter = (
                "emailAddresses/any(a:a/address eq "
                f"'{self._escape_odata_string(query_text)}')"
            )
            email_filter_escaped = quote(email_filter, safe="")
            url = (
                f"{self._contacts_path(normalized_mailbox, folderId)}"
                f"?{base_query}&$filter={email_filter_escaped}"
            )
            result = await self._request(url)
            contacts = [self._map_contact(contact) for contact in result.get("value", [])]
        else:
            contacts = []
            page_url: str | None = (
                f"{self._contacts_path(normalized_mailbox, folderId)}"
                f"?{_build_contact_read_query(top=100)}"
            )
            pages_remaining = max(1, min(maxPages, 10))
            while page_url and pages_remaining > 0 and len(contacts) < top:
                result = await self._request(page_url)
                contacts.extend(
                    contact
                    for contact in [
                        self._map_contact(raw) for raw in result.get("value", [])
                    ]
                    if self._contact_matches_query(contact, query_text)
                )
                page_url = result.get("@odata.nextLink")
                pages_remaining -= 1
            contacts = contacts[:top]

        return ContactsSearchResult(
            mailbox=normalized_mailbox or "me",
            query=query,
            contacts=contacts,
        )

    async def get_contact(
        self,
        *,
        mailbox: str | None = None,
        contactId: str,
        folderId: str | None = None,
    ) -> ContactGetResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        path = (
            f"{self._contacts_path(normalized_mailbox, folderId)}/{quote(contactId, safe='')}"
        )
        contact = await self._request(f"{path}?{_build_contact_read_query()}")
        return ContactGetResult(
            mailbox=normalized_mailbox or "me",
            contact=self._map_contact(contact),
        )

    async def create_contact(
        self,
        *,
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
        businessAddress: ContactAddress | dict[str, Any] | None = None,
        homeAddress: ContactAddress | dict[str, Any] | None = None,
        otherAddress: ContactAddress | dict[str, Any] | None = None,
    ) -> ContactMutationResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        created = await self._request(
            self._contacts_path(normalized_mailbox, folderId),
            method="POST",
            json_body=self._contact_payload(
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
            ),
        )
        # Re-fetch with $expand so the response surfaces personalHomePage
        # (stored as an extended property and not returned by POST by default).
        created_id = created.get("id") if isinstance(created, dict) else None
        if created_id:
            refreshed = await self._request(
                f"{self._contacts_path(normalized_mailbox, folderId)}/{quote(str(created_id), safe='')}"
                f"?{_build_contact_read_query()}"
            )
            contact = refreshed
        else:
            contact = created
        return ContactMutationResult(
            mailbox=normalized_mailbox or "me",
            contact=self._map_contact(contact),
        )

    async def update_contact(
        self,
        *,
        mailbox: str | None = None,
        contactId: str,
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
        businessAddress: ContactAddress | dict[str, Any] | None = None,
        homeAddress: ContactAddress | dict[str, Any] | None = None,
        otherAddress: ContactAddress | dict[str, Any] | None = None,
    ) -> ContactMutationResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        contact_path = (
            f"{self._contacts_path(normalized_mailbox, folderId)}/{quote(contactId, safe='')}"
        )
        await self._request(
            contact_path,
            method="PATCH",
            json_body=self._contact_payload(
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
            ),
        )
        # Re-fetch with $expand so the response surfaces personalHomePage
        # (stored as an extended property and not returned by PATCH by default).
        contact = await self._request(f"{contact_path}?{_build_contact_read_query()}")
        return ContactMutationResult(
            mailbox=normalized_mailbox or "me",
            contact=self._map_contact(contact),
        )

    async def delete_contact(
        self,
        *,
        mailbox: str | None = None,
        contactId: str,
        folderId: str | None = None,
    ) -> ContactMutationResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        await self._request(
            f"{self._contacts_path(normalized_mailbox, folderId)}/{quote(contactId, safe='')}",
            method="DELETE",
        )
        return ContactMutationResult(
            mailbox=normalized_mailbox or "me",
            contactId=contactId,
            deleted=True,
        )

    async def set_contact_categories(
        self,
        *,
        mailbox: str | None = None,
        contactId: str,
        categories: list[str],
        folderId: str | None = None,
    ) -> ContactMutationResult:
        return await self.update_contact(
            mailbox=mailbox,
            contactId=contactId,
            folderId=folderId,
            categories=categories,
        )

    async def add_contact_categories(
        self,
        *,
        mailbox: str | None = None,
        contactId: str,
        categories: list[str],
        folderId: str | None = None,
    ) -> ContactMutationResult:
        current = await self.get_contact(
            mailbox=mailbox,
            contactId=contactId,
            folderId=folderId,
        )
        existing = list(current.contact.categories)
        for category in categories:
            if category not in existing:
                existing.append(category)
        return await self.set_contact_categories(
            mailbox=mailbox,
            contactId=contactId,
            folderId=folderId,
            categories=existing,
        )

    async def remove_contact_categories(
        self,
        *,
        mailbox: str | None = None,
        contactId: str,
        categories: list[str],
        folderId: str | None = None,
    ) -> ContactMutationResult:
        current = await self.get_contact(
            mailbox=mailbox,
            contactId=contactId,
            folderId=folderId,
        )
        remove = set(categories)
        return await self.set_contact_categories(
            mailbox=mailbox,
            contactId=contactId,
            folderId=folderId,
            categories=[
                category for category in current.contact.categories if category not in remove
            ],
        )

    async def clear_contact_categories(
        self,
        *,
        mailbox: str | None = None,
        contactId: str,
        folderId: str | None = None,
    ) -> ContactMutationResult:
        return await self.set_contact_categories(
            mailbox=mailbox,
            contactId=contactId,
            folderId=folderId,
            categories=[],
        )

    async def list_contact_folders(
        self,
        *,
        mailbox: str | None = None,
        parentFolderId: str | None = None,
        top: int = 100,
    ) -> ContactFoldersResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        path = (
            f"{base}/contactFolders/{quote(parentFolderId, safe='')}/childFolders"
            if parentFolderId
            else f"{base}/contactFolders"
        )
        params = httpx.QueryParams(
            {"$top": str(min(top, 100)), "$select": CONTACT_FOLDER_SELECT}
        )
        result = await self._request(f"{path}?{params}")
        return ContactFoldersResult(
            mailbox=normalized_mailbox or "me",
            parentFolderId=parentFolderId,
            folders=[
                self._map_contact_folder(folder) for folder in result.get("value", [])
            ],
        )

    async def list_events(
        self,
        *,
        mailbox: str | None = None,
        start: str | None = None,
        end: str | None = None,
        top: int = 25,
    ) -> CalendarListEventsResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        start_value = start or _utc_now_iso()
        end_value = end or (
            datetime.now(UTC) + timedelta(days=7)
        ).isoformat().replace("+00:00", "Z")
        base = self._base_path(normalized_mailbox)
        params = httpx.QueryParams(
            {
                "startDateTime": start_value,
                "endDateTime": end_value,
                "$top": str(min(top, 100)),
                "$orderby": "start/dateTime",
                "$select": "id,subject,webLink,start,end,location,attendees,bodyPreview,body",
            }
        )
        result = await self._request(f"{base}/calendarView?{params}")
        return CalendarListEventsResult(
            mailbox=normalized_mailbox or "me",
            window=CalendarWindow(start=start_value, end=end_value),
            events=[self._map_event(event) for event in result["value"]],
        )

    async def create_event(
        self,
        *,
        mailbox: str | None = None,
        subject: str,
        start: str,
        end: str,
        timeZone: str = "UTC",
        attendees: list[str] | None = None,
        body: str | None = None,
        bodyType: str = "text",
        location: str | None = None,
    ) -> CalendarCreateEventResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        path = (
            f"/users/{quote(normalized_mailbox, safe='')}/calendar/events"
            if normalized_mailbox
            else "/me/calendar/events"
        )
        event = await self._request(
            path,
            method="POST",
            json_body={
                "subject": subject,
                "start": {"dateTime": start, "timeZone": timeZone},
                "end": {"dateTime": end, "timeZone": timeZone},
                "attendees": [
                    {"emailAddress": {"address": address}, "type": "required"}
                    for address in (attendees or [])
                ],
                "body": (
                    {
                        "contentType": self._graph_body_content_type(bodyType),
                        "content": body,
                    }
                    if body is not None
                    else None
                ),
                "location": (
                    {"displayName": location} if location is not None else None
                ),
            },
        )
        return CalendarCreateEventResult(
            mailbox=normalized_mailbox or "me",
            event=self._map_event(event),
        )

    async def update_event(
        self,
        *,
        mailbox: str | None = None,
        eventId: str,
        subject: str | None = None,
        start: str | None = None,
        end: str | None = None,
        timeZone: str = "UTC",
        attendees: list[str] | None = None,
        body: str | None = None,
        bodyType: str = "text",
        location: str | None = None,
    ) -> CalendarUpdateEventResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        payload = self._omit_none(
            {
                "subject": subject,
                "start": (
                    {"dateTime": start, "timeZone": timeZone}
                    if start is not None
                    else None
                ),
                "end": (
                    {"dateTime": end, "timeZone": timeZone}
                    if end is not None
                    else None
                ),
                "attendees": (
                    [
                        {
                            "emailAddress": {"address": address},
                            "type": "required",
                        }
                        for address in attendees
                    ]
                    if attendees is not None
                    else None
                ),
                "body": (
                    {
                        "contentType": self._graph_body_content_type(bodyType),
                        "content": body,
                    }
                    if body is not None
                    else None
                ),
                "location": (
                    {"displayName": location} if location is not None else None
                ),
            }
        )
        if not payload:
            raise ValueError("Provide at least one event field to update")

        event = await self._request(
            self._calendar_event_path(normalized_mailbox, eventId),
            method="PATCH",
            json_body=payload,
        )
        return CalendarUpdateEventResult(
            mailbox=normalized_mailbox or "me",
            event=self._map_event(event),
        )

    async def delete_event(
        self,
        *,
        mailbox: str | None = None,
        eventId: str,
    ) -> CalendarDeleteEventResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        await self._request(
            self._calendar_event_path(normalized_mailbox, eventId),
            method="DELETE",
        )
        return CalendarDeleteEventResult(
            mailbox=normalized_mailbox or "me",
            eventId=eventId,
            deleted=True,
        )

    async def _resolve_folder_id(self, base_path: str, folder_name: str) -> str:
        folder = await self._request(
            f"{base_path}/mailFolders('{quote(folder_name, safe='')}')?$select=id,displayName"
        )
        return str(folder["id"])

    async def _patch_message(
        self,
        *,
        mailbox: str | None,
        messageId: str,
        payload: dict[str, Any],
    ) -> MailUpdateMessageResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        message = await self._request(
            f"{base}/messages/{quote(messageId, safe='')}",
            method="PATCH",
            json_body=payload,
        )
        return MailUpdateMessageResult(
            mailbox=normalized_mailbox or "me",
            messageId=messageId,
            message=self._map_message_summary(message),
        )

    async def _get_mail_folder(self, base_path: str, folder_id: str) -> MailFolderInfo:
        folder = await self._request(
            f"{base_path}/mailFolders/{quote(folder_id, safe='')}?$select={MAIL_FOLDER_SELECT}"
        )
        return self._map_mail_folder(folder)

    async def _list_mail_folder_infos(
        self,
        base_path: str,
        *,
        parentFolderId: str | None,
        top: int = 100,
        currentPath: str | None = None,
    ) -> list[MailFolderInfo]:
        path = (
            f"{base_path}/mailFolders/{quote(parentFolderId, safe='')}/childFolders"
            if parentFolderId
            else f"{base_path}/mailFolders"
        )
        params = httpx.QueryParams(
            {
                "$top": str(min(top, 100)),
                "$select": MAIL_FOLDER_SELECT,
            }
        )
        result = await self._request(f"{path}?{params}")
        folders = [
            self._map_mail_folder(folder)
            for folder in result.get("value", [])
        ]
        if currentPath is not None:
            for folder in folders:
                folder.path = f"{currentPath}/{folder.displayName}"
        return folders

    async def _list_mail_folder_tree(
        self,
        base_path: str,
        *,
        parentFolderId: str | None,
        maxDepth: int,
        currentPath: str | None,
    ) -> list[MailFolderTreeNode]:
        folders = await self._list_mail_folder_infos(
            base_path,
            parentFolderId=parentFolderId,
            currentPath=currentPath,
        )
        nodes: list[MailFolderTreeNode] = []
        for folder in folders:
            path = folder.path or folder.displayName
            children = (
                await self._list_mail_folder_tree(
                    base_path,
                    parentFolderId=folder.id,
                    maxDepth=maxDepth - 1,
                    currentPath=path,
                )
                if maxDepth > 1 and folder.childFolderCount > 0
                else []
            )
            node_data = folder.model_dump(mode="python")
            node_data["path"] = path
            nodes.append(MailFolderTreeNode(**node_data, childFolders=children))
        return nodes

    async def _find_mail_folder_child(
        self,
        base_path: str,
        *,
        parentFolderId: str | None,
        displayName: str,
    ) -> MailFolderInfo:
        folders = await self._list_mail_folder_infos(
            base_path,
            parentFolderId=parentFolderId,
        )
        matches = [
            folder
            for folder in folders
            if folder.displayName.lower() == displayName.lower() or folder.id == displayName
        ]
        if not matches:
            raise RuntimeError(f"Mail folder was not found: {displayName}")
        if len(matches) > 1:
            raise RuntimeError(
                f"Multiple mail folders named {displayName!r}; use parentFolderId or folderId"
            )
        return matches[0]

    async def _resolve_mail_folder_by_path(
        self,
        base_path: str,
        folder_path: str,
    ) -> MailFolderInfo:
        parts = [part.strip() for part in folder_path.split("/") if part.strip()]
        if not parts:
            raise RuntimeError("folderPath must contain at least one folder name")
        parent_id: str | None = None
        resolved: MailFolderInfo | None = None
        path_parts: list[str] = []
        for part in parts:
            resolved = await self._find_mail_folder_child(
                base_path,
                parentFolderId=parent_id,
                displayName=part,
            )
            path_parts.append(resolved.displayName)
            resolved.path = "/".join(path_parts)
            parent_id = resolved.id
        if resolved is None:  # pragma: no cover - parts guard above
            raise RuntimeError("folderPath was invalid")
        return resolved

    async def _resolve_mail_folder_identifier(
        self,
        base_path: str,
        *,
        folderId: str | None,
        folderPath: str | None,
    ) -> str:
        if folderId:
            return folderId
        if folderPath:
            return (await self._resolve_mail_folder_by_path(base_path, folderPath)).id
        raise ValueError("Provide folderId or folderPath")

    def _mail_folder_messages_path(
        self,
        base_path: str,
        folder: str,
        folder_id: str | None,
    ) -> str:
        if folder_id:
            return f"{base_path}/mailFolders/{quote(folder_id, safe='')}/messages"
        return f"{base_path}/mailFolders('{quote(folder, safe='')}')/messages"

    def _mail_rules_path(self, base_path: str) -> str:
        return f"{base_path}/mailFolders/inbox/messageRules"

    def _base_path(self, mailbox: str | None) -> str:
        return f"/users/{quote(mailbox, safe='')}" if mailbox else "/me"

    def _normalize_mailbox(self, mailbox: str | None) -> str | None:
        value = (mailbox or "").strip()
        normalized = value or None
        self._pid_policy.require_mailbox(normalized)
        return normalized

    async def _request(
        self,
        path: str,
        *,
        method: str = "GET",
        headers: dict[str, str] | None = None,
        json_body: dict[str, Any] | None = None,
    ) -> Any:
        access_token = await self._auth_service.get_access_token()
        request_headers = {
            "Accept": "application/json",
            "Authorization": f"Bearer {access_token}",
            "Prefer": 'IdType="ImmutableId"',
            **({"Content-Type": "application/json"} if json_body is not None else {}),
            **(headers or {}),
        }

        async with self._client() as client:
            url = (
                path
                if path.startswith("https://")
                else f"https://graph.microsoft.com/v1.0{path}"
            )
            response = await client.request(
                method,
                url,
                headers=request_headers,
                json=self._omit_none(json_body) if json_body is not None else None,
            )

        if response.status_code == 204:
            return None

        data = response.json() if response.text else None
        if not response.is_success:
            error_message = (
                data.get("error", {}).get("message")
                if isinstance(data, dict)
                else None
            ) or (
                data.get("error_description")
                if isinstance(data, dict)
                else None
            ) or response.reason_phrase
            error_code = (
                data.get("error", {}).get("code")
                if isinstance(data, dict)
                else None
            )
            detail = f"{error_code}: {error_message}" if error_code else error_message
            raise RuntimeError(
                f"Microsoft Graph request failed ({response.status_code}): {detail}"
            )

        return data

    async def _request_bytes(self, path: str) -> bytes:
        access_token = await self._auth_service.get_access_token()
        async with self._client() as client:
            response = await client.request(
                "GET",
                f"https://graph.microsoft.com/v1.0{path}",
                headers={
                    "Accept": "*/*",
                    "Authorization": f"Bearer {access_token}",
                    "Prefer": 'IdType="ImmutableId"',
                },
            )

        if not response.is_success:
            detail = response.reason_phrase
            if response.text:
                try:
                    data = response.json()
                    if isinstance(data, dict):
                        detail = (
                            data.get("error", {}).get("message")
                            or data.get("error_description")
                            or detail
                        )
                except Exception:
                    detail = response.text
            raise RuntimeError(
                f"Microsoft Graph request failed ({response.status_code}): {detail}"
            )

        return response.content

    def _to_recipients(self, addresses: list[str] | None) -> list[dict[str, Any]] | None:
        cleaned = [address for address in (addresses or []) if address]
        if not cleaned:
            return None
        return [{"emailAddress": {"address": address}} for address in cleaned]

    async def _mail_rule_payload(
        self,
        base_path: str,
        *,
        displayName: str | None,
        sequence: int | None,
        isEnabled: bool | None,
        conditions: dict[str, Any] | None,
        actions: dict[str, Any] | None,
        exceptions: dict[str, Any] | None,
        fromAddresses: list[str] | None,
        senderContains: list[str] | None,
        subjectContains: list[str] | None,
        bodyContains: list[str] | None,
        sentToAddresses: list[str] | None,
        moveToFolderId: str | None,
        moveToFolderPath: str | None,
        markAsRead: bool | None,
        assignCategories: list[str] | None,
        stopProcessingRules: bool | None,
    ) -> dict[str, Any]:
        rule_conditions = dict(conditions or {})
        rule_actions = dict(actions or {})
        rule_exceptions = dict(exceptions or {})

        if fromAddresses is not None:
            rule_conditions["fromAddresses"] = self._to_recipients(fromAddresses)
        if senderContains is not None:
            rule_conditions["senderContains"] = senderContains
        if subjectContains is not None:
            rule_conditions["subjectContains"] = subjectContains
        if bodyContains is not None:
            rule_conditions["bodyContains"] = bodyContains
        if sentToAddresses is not None:
            rule_conditions["sentToAddresses"] = self._to_recipients(sentToAddresses)
        if moveToFolderPath and not moveToFolderId:
            moveToFolderId = (
                await self._resolve_mail_folder_by_path(base_path, moveToFolderPath)
            ).id
        if moveToFolderId is not None:
            rule_actions["moveToFolder"] = moveToFolderId
        if markAsRead is not None:
            rule_actions["markAsRead"] = markAsRead
        if assignCategories is not None:
            rule_actions["assignCategories"] = assignCategories
        if stopProcessingRules is not None:
            rule_actions["stopProcessingRules"] = stopProcessingRules

        return self._omit_none(
            {
                "displayName": displayName,
                "sequence": sequence,
                "isEnabled": isEnabled,
                "conditions": rule_conditions or None,
                "actions": rule_actions or None,
                "exceptions": rule_exceptions or None,
            }
        )

    def _message_filters(
        self,
        *,
        isRead: bool | None,
        hasAttachments: bool | None,
        importance: str | None,
        categories: list[str] | None,
        flagStatus: str | None,
        inferenceClassification: str | None,
    ) -> list[str]:
        filters: list[str] = []
        if isRead is not None:
            filters.append(f"isRead eq {'true' if isRead else 'false'}")
        if hasAttachments is not None:
            filters.append(
                f"hasAttachments eq {'true' if hasAttachments else 'false'}"
            )
        if importance:
            filters.append(f"importance eq '{self._escape_odata_string(importance)}'")
        for category in categories or []:
            filters.append(
                "categories/any(c:c eq "
                f"'{self._escape_odata_string(category)}')"
            )
        if flagStatus:
            filters.append(
                f"flag/flagStatus eq '{self._escape_odata_string(flagStatus)}'"
            )
        if inferenceClassification:
            filters.append(
                "inferenceClassification eq "
                f"'{self._escape_odata_string(inferenceClassification)}'"
            )
        return filters

    def _escape_odata_string(self, value: str) -> str:
        return value.replace("'", "''")

    def _contacts_path(self, mailbox: str | None, folder_id: str | None) -> str:
        base = self._base_path(mailbox)
        if folder_id:
            return f"{base}/contactFolders/{quote(folder_id, safe='')}/contacts"
        return f"{base}/contacts"

    def _calendar_event_path(self, mailbox: str | None, event_id: str) -> str:
        return f"{self._base_path(mailbox)}/calendar/events/{quote(event_id, safe='')}"

    def _contact_payload(
        self,
        *,
        displayName: str | None,
        givenName: str | None,
        surname: str | None,
        emailAddresses: list[str] | None,
        companyName: str | None,
        jobTitle: str | None,
        personalHomePage: str | None,
        personalNotes: str | None,
        businessPhones: list[str] | None,
        mobilePhone: str | None,
        categories: list[str] | None,
        businessAddress: ContactAddress | dict[str, Any] | None,
        homeAddress: ContactAddress | dict[str, Any] | None,
        otherAddress: ContactAddress | dict[str, Any] | None,
    ) -> dict[str, Any]:
        return {
            "displayName": displayName,
            "givenName": givenName,
            "surname": surname,
            "companyName": companyName,
            "jobTitle": jobTitle,
            "personalNotes": personalNotes,
            "singleValueExtendedProperties": self._personal_home_page_payload(
                personalHomePage
            ),
            "businessPhones": businessPhones,
            "mobilePhone": mobilePhone,
            "emailAddresses": [
                {"address": address, "name": displayName or address}
                for address in (emailAddresses or [])
                if address
            ]
            if emailAddresses is not None
            else None,
            "categories": categories,
            "businessAddress": self._contact_address_payload(businessAddress),
            "homeAddress": self._contact_address_payload(homeAddress),
            "otherAddress": self._contact_address_payload(otherAddress),
        }

    def _personal_home_page_payload(
        self,
        personal_home_page: str | None,
    ) -> list[dict[str, str]] | None:
        if personal_home_page is None:
            return None
        return [
            {
                "id": PERSONAL_HOME_PAGE_EXTENDED_PROPERTY_ID,
                "value": personal_home_page,
            }
        ]

    def _contact_address_payload(
        self,
        address: ContactAddress | dict[str, Any] | None,
    ) -> dict[str, Any] | None:
        if address is None:
            return None
        if isinstance(address, ContactAddress):
            return address.model_dump(mode="json")
        return {
            "street": address.get("street"),
            "city": address.get("city"),
            "state": address.get("state"),
            "countryOrRegion": address.get("countryOrRegion"),
            "postalCode": address.get("postalCode"),
        }

    def _contact_matches_query(self, contact: ContactInfo, query: str) -> bool:
        needle = query.lower()
        haystack = [
            contact.displayName,
            contact.givenName,
            contact.surname,
            contact.companyName,
            contact.jobTitle,
            contact.personalHomePage,
            contact.personalNotes,
            contact.mobilePhone,
            *contact.businessPhones,
            *contact.emailAddresses,
            *contact.categories,
            *self._contact_address_values(contact.businessAddress),
            *self._contact_address_values(contact.homeAddress),
            *self._contact_address_values(contact.otherAddress),
        ]
        return any(needle in value.lower() for value in haystack if value)

    def _contact_address_values(self, address: ContactAddress | None) -> list[str]:
        if address is None:
            return []
        return [
            value
            for value in [
                address.street,
                address.city,
                address.state,
                address.countryOrRegion,
                address.postalCode,
            ]
            if value
        ]

    def _attachment_unsupported_reason(
        self,
        attachment: AttachmentInfo,
        *,
        maxBytes: int,
    ) -> str | None:
        if attachment.attachmentType not in (None, "#microsoft.graph.fileAttachment"):
            return "Only file attachments can be read as content in this MCP server"
        if attachment.size is not None and attachment.size > maxBytes:
            return f"Attachment size exceeds maxBytes={maxBytes}"
        if self._is_pdf_attachment(attachment):
            if PdfReader is None:
                return "PDF text extraction requires the pypdf package"
            return None
        content_type = (attachment.contentType or "").split(";")[0].strip().lower()
        name = (attachment.name or "").lower()
        is_safe_type = content_type.startswith("text/") or content_type in SAFE_ATTACHMENT_CONTENT_TYPES
        is_safe_extension = any(name.endswith(extension) for extension in SAFE_ATTACHMENT_EXTENSIONS)
        if not is_safe_type and not is_safe_extension:
            if self._is_image_attachment(attachment):
                return (
                    "Attachment is an image; use mail_get_attachment_image or "
                    "mail_get_inline_images to view it"
                )
            return "Attachment content type is not text-like and was not returned"
        return None

    async def _attachment_bytes(
        self,
        base_path: str,
        *,
        messageId: str,
        attachment: dict[str, Any],
    ) -> bytes:
        """Return raw attachment bytes, preferring the inline ``contentBytes``.

        Microsoft Graph returns ``contentBytes`` on file attachments read either
        singly or through the collection, so the extra ``/$value`` round trip is
        only needed when Graph omitted it.
        """

        content_bytes_value = attachment.get("contentBytes")
        if isinstance(content_bytes_value, str):
            return base64.b64decode(content_bytes_value)
        attachment_id = str(attachment["id"])
        return await self._request_bytes(
            f"{base_path}/messages/{quote(messageId, safe='')}"
            f"/attachments/{quote(attachment_id, safe='')}/$value"
        )

    def _image_mime_type(self, attachment: AttachmentInfo) -> str | None:
        content_type = (attachment.contentType or "").split(";")[0].strip().lower()
        content_type = IMAGE_CONTENT_TYPE_ALIASES.get(content_type, content_type)
        if content_type in SUPPORTED_IMAGE_CONTENT_TYPES:
            return content_type
        name = (attachment.name or "").lower()
        for extension, mime_type in IMAGE_EXTENSION_CONTENT_TYPES.items():
            if name.endswith(extension):
                return mime_type
        return None

    def _is_image_attachment(self, attachment: AttachmentInfo) -> bool:
        content_type = (attachment.contentType or "").split(";")[0].strip().lower()
        return content_type.startswith("image/") or (
            self._image_mime_type(attachment) is not None
        )

    def _is_inline_attachment(self, attachment: AttachmentInfo) -> bool:
        return attachment.isInline or bool(attachment.contentId)

    def _attachment_image_unsupported_reason(
        self,
        attachment: AttachmentInfo,
        *,
        maxBytes: int,
    ) -> str | None:
        if attachment.attachmentType not in (None, "#microsoft.graph.fileAttachment"):
            return "Only file attachments can be read as images in this MCP server"
        if self._image_mime_type(attachment) is None:
            return (
                "Attachment is not a PNG, JPEG, GIF, or WEBP image and cannot be "
                "returned as image content"
            )
        if attachment.size is not None and attachment.size > maxBytes:
            return f"Image size exceeds maxBytes={maxBytes}"
        return None

    def _build_attachment_image(
        self,
        attachment: AttachmentInfo,
        content_bytes: bytes,
    ) -> AttachmentImage:
        mime_type = self._image_mime_type(attachment) or "image/png"
        return AttachmentImage(
            attachment=attachment,
            mimeType=mime_type,
            dataBase64=base64.b64encode(content_bytes).decode("ascii"),
            byteSize=len(content_bytes),
        )

    def _is_pdf_attachment(self, attachment: AttachmentInfo) -> bool:
        content_type = (attachment.contentType or "").split(";")[0].strip().lower()
        name = (attachment.name or "").lower()
        return content_type == "application/pdf" or name.endswith(".pdf")

    def _render_pdf_pages(
        self,
        content_bytes: bytes,
        *,
        firstPage: int,
        maxPages: int,
        longEdge: int,
        maxTotalBytes: int,
    ) -> tuple[int, list[PdfPageImage], bool]:
        """Rasterize a slice of a PDF to JPEG page images.

        Returns the document's total page count, the rendered pages, and
        whether rendering stopped early because of ``maxPages`` or the
        ``maxTotalBytes`` budget.
        """

        try:
            document = pypdfium2.PdfDocument(content_bytes)
        except Exception as error:
            raise _PdfRenderError(
                f"Could not open PDF (it may be password-protected or damaged): {error}"
            ) from error

        try:
            page_count = len(document)
            start = max(firstPage, 1) - 1
            end = min(start + max(maxPages, 0), page_count)

            pages: list[PdfPageImage] = []
            truncated = end < page_count
            total_bytes = 0

            for index in range(start, end):
                page = document[index]
                width, height = page.get_size()
                longest = max(width, height) or 1
                scale = min(longEdge / longest, MAX_PDF_RENDER_SCALE)
                image = page.render(scale=scale).to_pil()
                buffer = io.BytesIO()
                image.convert("RGB").save(
                    buffer,
                    format="JPEG",
                    quality=PDF_PAGE_JPEG_QUALITY,
                    optimize=True,
                )
                rendered = buffer.getvalue()

                if pages and total_bytes + len(rendered) > maxTotalBytes:
                    truncated = True
                    break

                total_bytes += len(rendered)
                pages.append(
                    PdfPageImage(
                        pageNumber=index + 1,
                        mimeType="image/jpeg",
                        dataBase64=base64.b64encode(rendered).decode("ascii"),
                        byteSize=len(rendered),
                        widthPx=image.width,
                        heightPx=image.height,
                    )
                )

            return page_count, pages, truncated
        finally:
            document.close()

    def _extract_pdf_text(self, content_bytes: bytes) -> str:
        if PdfReader is None:
            raise RuntimeError("PDF text extraction requires the pypdf package")
        try:
            reader = PdfReader(io.BytesIO(content_bytes))
        except Exception as error:  # pragma: no cover - defensive around parser internals
            raise RuntimeError(f"Could not read PDF attachment: {error}") from error

        pages: list[str] = []
        for index, page in enumerate(reader.pages, start=1):
            text = page.extract_text() or ""
            if text.strip():
                pages.append(f"--- Page {index} ---\n{text.strip()}")
        return "\n\n".join(pages)

    def _graph_body_content_type(self, body_type: str) -> str:
        match body_type.lower():
            case "html":
                return "HTML"
            case "text":
                return "Text"
            case _:
                return body_type

    def _omit_none(self, value: Any) -> Any:
        if isinstance(value, dict):
            return {
                key: self._omit_none(nested_value)
                for key, nested_value in value.items()
                if nested_value is not None
            }
        if isinstance(value, list):
            return [self._omit_none(item) for item in value]
        return value

    def _map_mail_folder(self, folder: dict[str, Any]) -> MailFolderInfo:
        return MailFolderInfo(
            id=str(folder["id"]),
            displayName=str(folder.get("displayName") or ""),
            parentFolderId=self._nullable_string(folder.get("parentFolderId")),
            childFolderCount=int(folder.get("childFolderCount") or 0),
            totalItemCount=int(folder.get("totalItemCount") or 0),
            unreadItemCount=int(folder.get("unreadItemCount") or 0),
            isHidden=(
                bool(folder["isHidden"])
                if folder.get("isHidden") is not None
                else None
            ),
        )

    def _map_attachment(self, attachment: dict[str, Any]) -> AttachmentInfo:
        return AttachmentInfo(
            id=str(attachment["id"]),
            name=self._nullable_string(attachment.get("name")),
            contentType=self._nullable_string(attachment.get("contentType")),
            size=int(attachment.get("size") or 0)
            if attachment.get("size") is not None
            else None,
            isInline=bool(attachment.get("isInline", False)),
            contentId=self._nullable_string(attachment.get("contentId")),
            lastModifiedDateTime=self._nullable_string(
                attachment.get("lastModifiedDateTime")
            ),
            attachmentType=self._nullable_string(attachment.get("@odata.type")),
        )

    def _map_category(self, category: dict[str, Any]) -> MailCategoryInfo:
        return MailCategoryInfo(
            id=self._nullable_string(category.get("id")),
            displayName=str(category.get("displayName") or ""),
            color=self._nullable_string(category.get("color")),
        )

    def _map_rule(self, rule: dict[str, Any]) -> MailRuleInfo:
        return MailRuleInfo(
            id=str(rule["id"]),
            displayName=str(rule.get("displayName") or ""),
            sequence=(
                int(rule["sequence"])
                if rule.get("sequence") is not None
                else None
            ),
            isEnabled=(
                bool(rule["isEnabled"])
                if rule.get("isEnabled") is not None
                else None
            ),
            hasError=(
                bool(rule["hasError"])
                if rule.get("hasError") is not None
                else None
            ),
            isReadOnly=(
                bool(rule["isReadOnly"])
                if rule.get("isReadOnly") is not None
                else None
            ),
            conditions=rule.get("conditions") or {},
            actions=rule.get("actions") or {},
            exceptions=rule.get("exceptions") or {},
        )

    def _map_contact_folder(self, folder: dict[str, Any]) -> ContactFolderInfo:
        return ContactFolderInfo(
            id=str(folder["id"]),
            displayName=str(folder.get("displayName") or ""),
            parentFolderId=self._nullable_string(folder.get("parentFolderId")),
            childFolderCount=int(folder.get("childFolderCount") or 0),
        )

    def _map_contact_address(self, address: Any) -> ContactAddress | None:
        if not isinstance(address, dict):
            return None
        mapped = ContactAddress(
            street=self._nullable_string(address.get("street")),
            city=self._nullable_string(address.get("city")),
            state=self._nullable_string(address.get("state")),
            countryOrRegion=self._nullable_string(address.get("countryOrRegion")),
            postalCode=self._nullable_string(address.get("postalCode")),
        )
        if not any(
            [
                mapped.street,
                mapped.city,
                mapped.state,
                mapped.countryOrRegion,
                mapped.postalCode,
            ]
        ):
            return None
        return mapped

    def _map_contact(self, contact: dict[str, Any]) -> ContactInfo:
        return ContactInfo(
            id=str(contact["id"]),
            displayName=self._nullable_string(contact.get("displayName")),
            givenName=self._nullable_string(contact.get("givenName")),
            surname=self._nullable_string(contact.get("surname")),
            companyName=self._nullable_string(contact.get("companyName")),
            jobTitle=self._nullable_string(contact.get("jobTitle")),
            personalHomePage=self._map_personal_home_page(contact),
            personalNotes=self._nullable_string(contact.get("personalNotes")),
            businessPhones=[
                str(phone) for phone in (contact.get("businessPhones") or []) if phone
            ],
            mobilePhone=self._nullable_string(contact.get("mobilePhone")),
            emailAddresses=[
                address
                for address in [
                    self._nullable_string(
                        email.get("address") if isinstance(email, dict) else email
                    )
                    for email in (contact.get("emailAddresses") or [])
                ]
                if address
            ],
            categories=[
                str(category)
                for category in (contact.get("categories") or [])
                if category
            ],
            businessAddress=self._map_contact_address(contact.get("businessAddress")),
            homeAddress=self._map_contact_address(contact.get("homeAddress")),
            otherAddress=self._map_contact_address(contact.get("otherAddress")),
            parentFolderId=self._nullable_string(contact.get("parentFolderId")),
        )

    def _map_personal_home_page(self, contact: dict[str, Any]) -> str | None:
        raw_props = contact.get("singleValueExtendedProperties")
        for property_value in raw_props or []:
            if not isinstance(property_value, dict):
                continue
            if property_value.get("id") == PERSONAL_HOME_PAGE_EXTENDED_PROPERTY_ID:
                return self._nullable_string(property_value.get("value"))
        return None

    def _map_message_summary(self, message: dict[str, Any]) -> MessageSummary:
        return MessageSummary(
            id=str(message["id"]),
            subject=self._pid_policy.redact_text(str(message.get("subject") or "")),
            from_=self._map_email_address(message.get("from")),
            sender=self._map_email_address(message.get("sender")),
            replyTo=self._map_recipients(message.get("replyTo")),
            receivedDateTime=self._nullable_string(message.get("receivedDateTime")),
            sentDateTime=self._nullable_string(message.get("sentDateTime")),
            bodyPreview=self._pid_policy.redact_text(
                str(message.get("bodyPreview") or "")
            ),
            webLink=self._nullable_string(message.get("webLink")),
            isDraft=bool(message.get("isDraft", False)),
            isRead=(
                bool(message["isRead"])
                if message.get("isRead") is not None
                else None
            ),
            hasAttachments=(
                bool(message["hasAttachments"])
                if message.get("hasAttachments") is not None
                else None
            ),
            importance=self._nullable_string(message.get("importance")),
            categories=[
                str(category) for category in (message.get("categories") or [])
            ],
            flagStatus=self._nullable_string((message.get("flag") or {}).get("flagStatus")),
            inferenceClassification=self._nullable_string(
                message.get("inferenceClassification")
            ),
            parentFolderId=self._nullable_string(message.get("parentFolderId")),
            internetMessageId=self._nullable_string(message.get("internetMessageId")),
            conversationId=self._nullable_string(message.get("conversationId")),
        )

    def _map_full_message(self, message: dict[str, Any]) -> FullMessage:
        body = message.get("body") or {}
        return FullMessage(
            id=str(message["id"]),
            subject=self._pid_policy.redact_text(str(message.get("subject") or "")),
            from_=self._map_email_address(message.get("from")),
            sender=self._map_email_address(message.get("sender")),
            replyTo=self._map_recipients(message.get("replyTo")),
            to=self._map_recipients(message.get("toRecipients")),
            cc=self._map_recipients(message.get("ccRecipients")),
            bcc=self._map_recipients(message.get("bccRecipients")),
            receivedDateTime=self._nullable_string(message.get("receivedDateTime")),
            sentDateTime=self._nullable_string(message.get("sentDateTime")),
            bodyPreview=self._pid_policy.redact_text(
                str(message.get("bodyPreview") or "")
            ),
            body=MessageBody(
                contentType=str(body.get("contentType") or "text"),
                content=self._pid_policy.redact_text(str(body.get("content") or "")),
            ),
            webLink=self._nullable_string(message.get("webLink")),
            isDraft=bool(message.get("isDraft", False)),
            isRead=(
                bool(message["isRead"])
                if message.get("isRead") is not None
                else None
            ),
            hasAttachments=(
                bool(message["hasAttachments"])
                if message.get("hasAttachments") is not None
                else None
            ),
            importance=self._nullable_string(message.get("importance")),
            categories=[
                str(category) for category in (message.get("categories") or [])
            ],
            flagStatus=self._nullable_string((message.get("flag") or {}).get("flagStatus")),
            inferenceClassification=self._nullable_string(
                message.get("inferenceClassification")
            ),
            parentFolderId=self._nullable_string(message.get("parentFolderId")),
            internetMessageId=self._nullable_string(message.get("internetMessageId")),
            conversationId=self._nullable_string(message.get("conversationId")),
        )

    def _map_event(self, event: dict[str, Any]) -> CalendarEvent:
        body = event.get("body") or {}
        attendees = event.get("attendees") or []
        return CalendarEvent(
            id=str(event["id"]),
            subject=str(event.get("subject") or ""),
            webLink=self._nullable_string(event.get("webLink")),
            start=CalendarDateTime(
                dateTime=self._nullable_string((event.get("start") or {}).get("dateTime")),
                timeZone=self._nullable_string((event.get("start") or {}).get("timeZone")),
            ),
            end=CalendarDateTime(
                dateTime=self._nullable_string((event.get("end") or {}).get("dateTime")),
                timeZone=self._nullable_string((event.get("end") or {}).get("timeZone")),
            ),
            location=self._nullable_string((event.get("location") or {}).get("displayName")),
            attendees=[
                CalendarAttendee(
                    address=self._nullable_string((attendee.get("emailAddress") or {}).get("address")),
                    name=self._nullable_string((attendee.get("emailAddress") or {}).get("name")),
                    type=self._nullable_string(attendee.get("type")),
                    response=self._nullable_string((attendee.get("status") or {}).get("response")),
                )
                for attendee in attendees
            ],
            bodyPreview=self._pid_policy.redact_text(
                str(event.get("bodyPreview") or "")
            ),
            body=MessageBody(
                contentType=str(body.get("contentType") or "text"),
                content=self._pid_policy.redact_text(str(body.get("content") or "")),
            ),
        )

    def _map_email_address(self, recipient: dict[str, Any] | None) -> str | None:
        email_address = recipient.get("emailAddress") if isinstance(recipient, dict) else None
        return self._nullable_string(email_address.get("address") if isinstance(email_address, dict) else None)

    def _map_recipients(self, recipients: list[dict[str, Any]] | None) -> list[str]:
        addresses = []
        for recipient in recipients or []:
            email = self._map_email_address(recipient)
            if email:
                addresses.append(email)
        return addresses

    def _nullable_string(self, value: Any) -> str | None:
        if value is None:
            return None
        text = str(value)
        return text if text else None
