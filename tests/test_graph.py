from __future__ import annotations

import json

import httpx
import pytest

from m365_mcp.microsoft_graph import MicrosoftGraphClient


class StaticAuthService:
    async def get_access_token(self) -> str:
        return "access-token"


@pytest.mark.anyio
async def test_list_messages_and_search_messages() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        assert request.headers["authorization"] == "Bearer access-token"

        if request.url.path.endswith("/mailFolders('Inbox')/messages"):
            assert request.url.params["$top"] == "100"
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "m1",
                            "subject": "Hello",
                            "from": {"emailAddress": {"address": "sender@example.com"}},
                            "receivedDateTime": "2026-04-21T12:00:00Z",
                            "sentDateTime": "2026-04-21T12:00:00Z",
                            "bodyPreview": "Preview",
                            "webLink": "https://outlook.example/messages/m1",
                            "isDraft": False,
                            "conversationId": "conv-1",
                        }
                    ]
                },
            )

        if request.url.path.endswith("/messages"):
            assert request.headers["consistencylevel"] == "eventual"
            assert request.url.params["$top"] == "50"
            assert request.url.params["$search"] == "\"from:\\\"boss\\\"\""
            return httpx.Response(
                200,
                json={"value": [{"id": "m2", "subject": "Search hit", "bodyPreview": ""}]},
            )

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    listed = await graph.list_messages(mailbox=None, folder="Inbox", top=999)
    assert listed.mailbox == "me"
    assert listed.messages[0].from_ == "sender@example.com"

    searched = await graph.search_messages(query='from:"boss"', top=99)
    assert searched.mailbox == "me"
    assert searched.messages[0].id == "m2"

    await client.aclose()


@pytest.mark.anyio
async def test_get_message_and_list_drafts() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path.endswith("/messages/msg-123"):
            return httpx.Response(
                200,
                json={
                    "id": "msg-123",
                    "subject": "Full message",
                    "from": {"emailAddress": {"address": "sender@example.com"}},
                    "toRecipients": [{"emailAddress": {"address": "to@example.com"}}],
                    "ccRecipients": [{"emailAddress": {"address": "cc@example.com"}}],
                    "bccRecipients": [{"emailAddress": {"address": "bcc@example.com"}}],
                    "bodyPreview": "Preview",
                    "body": {"contentType": "html", "content": "<p>Hello</p>"},
                    "isDraft": False,
                },
            )

        if request.url.path.endswith("/mailFolders('Drafts')/messages"):
            return httpx.Response(
                200,
                json={"value": [{"id": "draft-1", "subject": "Draft", "isDraft": True}]},
            )

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    message = await graph.get_message(mailbox="shared@example.com", messageId="msg-123")
    assert message.mailbox == "shared@example.com"
    assert message.message.to == ["to@example.com"]
    assert message.message.body.contentType == "html"

    drafts = await graph.list_drafts(mailbox="shared@example.com", top=10)
    assert drafts.drafts[0].isDraft is True

    await client.aclose()


@pytest.mark.anyio
async def test_create_send_and_move_message() -> None:
    requests: list[tuple[str, str, dict[str, object] | None]] = []

    def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content.decode("utf-8")) if request.content else None
        requests.append((request.method, request.url.path, body))

        if request.method == "POST" and request.url.path.endswith("/messages"):
            return httpx.Response(
                200,
                json={
                    "id": "draft-99",
                    "subject": "Created",
                    "from": {"emailAddress": {"address": "delegated@example.com"}},
                    "bodyPreview": "Draft preview",
                    "isDraft": True,
                },
            )

        if request.method == "POST" and request.url.path.endswith("/messages/draft-99/send"):
            return httpx.Response(204)

        if request.url.path.endswith("/mailFolders('Archive')"):
            return httpx.Response(200, json={"id": "folder-archive", "displayName": "Archive"})

        if request.method == "POST" and request.url.path.endswith("/messages/mail-1/move"):
            return httpx.Response(
                200,
                json={"id": "mail-1", "subject": "Moved", "bodyPreview": "Moved preview"},
            )

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    draft = await graph.create_draft(
        mailbox="shared@example.com",
        subject="Created",
        to=["a@example.com"],
        cc=["b@example.com"],
        bcc=["c@example.com"],
        body="Hello",
        bodyType="html",
        from_="delegated@example.com",
    )
    assert draft.draft.id == "draft-99"

    sent = await graph.send_draft(mailbox="shared@example.com", messageId="draft-99")
    assert sent.sent is True

    moved = await graph.move_message(
        mailbox="shared@example.com",
        messageId="mail-1",
        destinationFolder="Archive",
    )
    assert moved.destinationFolder == "Archive"

    create_body = requests[0][2]
    assert create_body is not None
    assert create_body["from"] == {"emailAddress": {"address": "delegated@example.com"}}
    assert create_body["body"] == {"contentType": "HTML", "content": "Hello"}
    assert create_body["toRecipients"] == [
        {"emailAddress": {"address": "a@example.com"}}
    ]
    assert create_body["ccRecipients"] == [
        {"emailAddress": {"address": "b@example.com"}}
    ]
    assert create_body["bccRecipients"] == [
        {"emailAddress": {"address": "c@example.com"}}
    ]

    move_body = requests[-1][2]
    assert move_body == {"destinationId": "folder-archive"}

    await client.aclose()


@pytest.mark.anyio
async def test_create_draft_omits_empty_optional_fields() -> None:
    captured_body: dict[str, object] | None = None

    def handler(request: httpx.Request) -> httpx.Response:
        nonlocal captured_body

        if request.method == "POST" and request.url.path.endswith("/messages"):
            captured_body = json.loads(request.content.decode("utf-8"))
            return httpx.Response(
                200,
                json={
                    "id": "draft-minimal",
                    "subject": "Minimal",
                    "bodyPreview": "Draft preview",
                    "isDraft": True,
                },
            )

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    draft = await graph.create_draft(
        subject="Minimal",
        to=[],
        cc=None,
        bcc=None,
        body="Hello",
        bodyType="text",
        from_=None,
    )

    assert draft.draft.id == "draft-minimal"
    assert captured_body == {
        "subject": "Minimal",
        "body": {"contentType": "Text", "content": "Hello"},
    }

    await client.aclose()


@pytest.mark.anyio
async def test_list_and_create_events_and_graph_errors() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content.decode("utf-8")) if request.content else None

        if request.url.path.endswith("/calendarView"):
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "event-1",
                            "subject": "Planning",
                            "start": {"dateTime": "2026-04-22T16:00:00", "timeZone": "UTC"},
                            "end": {"dateTime": "2026-04-22T17:00:00", "timeZone": "UTC"},
                            "location": {"displayName": "Room 1"},
                            "attendees": [
                                {
                                    "emailAddress": {"address": "attendee@example.com", "name": "Attendee"},
                                    "type": "required",
                                    "status": {"response": "accepted"},
                                }
                            ],
                            "bodyPreview": "Preview",
                            "body": {"contentType": "text", "content": "Agenda"},
                        }
                    ]
                },
            )

        if request.url.path.endswith("/calendar/events"):
            assert body == {
                "subject": "Created event",
                "start": {"dateTime": "2026-04-23T16:00:00", "timeZone": "UTC"},
                "end": {"dateTime": "2026-04-23T17:00:00", "timeZone": "UTC"},
                "attendees": [],
                "body": {"contentType": "HTML", "content": "Agenda"},
            }
            return httpx.Response(
                200,
                json={
                    "id": "event-2",
                    "subject": "Created event",
                    "start": {"dateTime": "2026-04-23T16:00:00", "timeZone": "UTC"},
                    "end": {"dateTime": "2026-04-23T17:00:00", "timeZone": "UTC"},
                    "attendees": [],
                    "bodyPreview": "",
                    "body": {"contentType": "text", "content": ""},
                },
            )

        if request.url.path.endswith("/messages/bad-id"):
            return httpx.Response(
                404,
                json={"error": {"code": "ErrorItemNotFound", "message": "No such message"}},
            )

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    events = await graph.list_events(mailbox="shared@example.com", start="2026-04-22T00:00:00Z", end="2026-04-23T00:00:00Z")
    assert events.events[0].attendees[0].response == "accepted"

    created = await graph.create_event(
        subject="Created event",
        start="2026-04-23T16:00:00",
        end="2026-04-23T17:00:00",
        mailbox="shared@example.com",
        body="Agenda",
        bodyType="html",
    )
    assert created.event.id == "event-2"

    with pytest.raises(RuntimeError, match="ErrorItemNotFound: No such message"):
        await graph.get_message(messageId="bad-id")

    await client.aclose()
