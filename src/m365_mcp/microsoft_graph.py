from __future__ import annotations

from contextlib import asynccontextmanager
from datetime import UTC, datetime, timedelta
from typing import Any
from urllib.parse import quote

import httpx

from .models import (
    CalendarAttendee,
    CalendarCreateEventResult,
    CalendarDateTime,
    CalendarEvent,
    CalendarListEventsResult,
    CalendarWindow,
    FullMessage,
    MailCreateDraftResult,
    MailGetResult,
    MailListDraftsResult,
    MailListResult,
    MailMoveResult,
    MailSearchResult,
    MailSendDraftResult,
    MessageBody,
    MessageSummary,
)
from .microsoft_auth import MicrosoftAuthService


def _utc_now_iso() -> str:
    return datetime.now(UTC).isoformat().replace("+00:00", "Z")


class MicrosoftGraphClient:
    def __init__(
        self,
        auth_service: MicrosoftAuthService,
        http_client: httpx.AsyncClient | None = None,
    ) -> None:
        self._auth_service = auth_service
        self._http_client = http_client

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
        top: int = 25,
    ) -> MailListResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        query = httpx.QueryParams(
            {
                "$top": str(min(top, 100)),
                "$select": (
                    "id,subject,from,receivedDateTime,sentDateTime,"
                    "bodyPreview,webLink,isDraft,conversationId"
                ),
            }
        )
        result = await self._request(
            f"{base}/mailFolders('{quote(folder, safe='')}')/messages?{query}"
        )

        return MailListResult(
            mailbox=normalized_mailbox or "me",
            folder=folder,
            messages=[self._map_message_summary(message) for message in result["value"]],
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
        params = httpx.QueryParams(
            {
                "$top": str(min(top, 50)),
                "$search": f"\"{query.replace('\"', '\\\"')}\"",
                "$select": (
                    "id,subject,from,receivedDateTime,sentDateTime,"
                    "bodyPreview,webLink,isDraft,conversationId"
                ),
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
                "$select": (
                    "id,subject,from,toRecipients,ccRecipients,bccRecipients,"
                    "receivedDateTime,sentDateTime,bodyPreview,body,webLink,"
                    "isDraft,importance,conversationId"
                ),
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
                "$select": (
                    "id,subject,from,receivedDateTime,sentDateTime,"
                    "bodyPreview,webLink,isDraft,conversationId"
                ),
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
                    "contentType": bodyType,
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
    ) -> MailMoveResult:
        normalized_mailbox = self._normalize_mailbox(mailbox)
        base = self._base_path(normalized_mailbox)
        destination_id = (
            destinationFolder
            if destinationFolderIsId
            else await self._resolve_folder_id(base, destinationFolder)
        )
        moved = await self._request(
            f"{base}/messages/{quote(messageId, safe='')}/move",
            method="POST",
            json_body={"destinationId": destination_id},
        )

        return MailMoveResult(
            mailbox=normalized_mailbox or "me",
            destinationFolder=destinationFolder,
            movedMessage=self._map_message_summary(moved),
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
                    {"contentType": bodyType, "content": body}
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

    async def _resolve_folder_id(self, base_path: str, folder_name: str) -> str:
        folder = await self._request(
            f"{base_path}/mailFolders('{quote(folder_name, safe='')}')?$select=id,displayName"
        )
        return str(folder["id"])

    def _base_path(self, mailbox: str | None) -> str:
        return f"/users/{quote(mailbox, safe='')}" if mailbox else "/me"

    def _normalize_mailbox(self, mailbox: str | None) -> str | None:
        value = (mailbox or "").strip()
        return value or None

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
            **({"Content-Type": "application/json"} if json_body is not None else {}),
            **(headers or {}),
        }

        async with self._client() as client:
            response = await client.request(
                method,
                f"https://graph.microsoft.com/v1.0{path}",
                headers=request_headers,
                json=json_body,
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

    def _to_recipients(self, addresses: list[str] | None) -> list[dict[str, Any]] | None:
        cleaned = [address for address in (addresses or []) if address]
        if not cleaned:
            return None
        return [{"emailAddress": {"address": address}} for address in cleaned]

    def _map_message_summary(self, message: dict[str, Any]) -> MessageSummary:
        return MessageSummary(
            id=str(message["id"]),
            subject=str(message.get("subject") or ""),
            from_=self._map_email_address(message.get("from")),
            receivedDateTime=self._nullable_string(message.get("receivedDateTime")),
            sentDateTime=self._nullable_string(message.get("sentDateTime")),
            bodyPreview=str(message.get("bodyPreview") or ""),
            webLink=self._nullable_string(message.get("webLink")),
            isDraft=bool(message.get("isDraft", False)),
            conversationId=self._nullable_string(message.get("conversationId")),
        )

    def _map_full_message(self, message: dict[str, Any]) -> FullMessage:
        body = message.get("body") or {}
        return FullMessage(
            id=str(message["id"]),
            subject=str(message.get("subject") or ""),
            from_=self._map_email_address(message.get("from")),
            to=self._map_recipients(message.get("toRecipients")),
            cc=self._map_recipients(message.get("ccRecipients")),
            bcc=self._map_recipients(message.get("bccRecipients")),
            receivedDateTime=self._nullable_string(message.get("receivedDateTime")),
            sentDateTime=self._nullable_string(message.get("sentDateTime")),
            bodyPreview=str(message.get("bodyPreview") or ""),
            body=MessageBody(
                contentType=str(body.get("contentType") or "text"),
                content=str(body.get("content") or ""),
            ),
            webLink=self._nullable_string(message.get("webLink")),
            isDraft=bool(message.get("isDraft", False)),
            importance=self._nullable_string(message.get("importance")),
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
            bodyPreview=str(event.get("bodyPreview") or ""),
            body=MessageBody(
                contentType=str(body.get("contentType") or "text"),
                content=str(body.get("content") or ""),
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
