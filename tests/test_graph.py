from __future__ import annotations

import base64
import json

import httpx
import pytest

from m365_mcp import microsoft_graph as graph_module
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
async def test_direct_send_and_reply() -> None:
    requests: list[tuple[str, str, dict[str, object] | None]] = []

    def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content.decode("utf-8")) if request.content else None
        requests.append((request.method, request.url.path, body))

        if request.method == "POST" and request.url.path.endswith("/sendMail"):
            assert body == {
                "message": {
                    "subject": "Quick note",
                    "body": {"contentType": "Text", "content": "Hello"},
                    "toRecipients": [
                        {"emailAddress": {"address": "a@example.com"}}
                    ],
                },
                "saveToSentItems": True,
            }
            return httpx.Response(202)

        if request.method == "POST" and request.url.path.endswith("/messages/msg-1/replyAll"):
            assert body == {
                "message": {
                    "body": {"contentType": "HTML", "content": "<p>Thanks</p>"}
                }
            }
            return httpx.Response(202)

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    sent = await graph.send_mail(
        subject="Quick note",
        to=["a@example.com"],
        body="Hello",
    )
    assert sent.sent is True
    assert sent.subject == "Quick note"

    replied = await graph.send_reply(
        messageId="msg-1",
        comment="<p>Thanks</p>",
        replyAll=True,
    )
    assert replied.sent is True
    assert replied.messageId == "msg-1"
    assert len(requests) == 2

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
async def test_folder_navigation_inbox_filters_and_nested_move() -> None:
    requests: list[tuple[str, str, dict[str, str], dict[str, object] | None]] = []

    def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content.decode("utf-8")) if request.content else None
        requests.append((request.method, request.url.path, dict(request.url.params), body))

        if request.url.path.endswith("/mailFolders") and request.method == "GET":
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "inbox-id",
                            "displayName": "Inbox",
                            "childFolderCount": 1,
                            "totalItemCount": 10,
                            "unreadItemCount": 2,
                        }
                    ]
                },
            )

        if request.url.path.endswith("/mailFolders/inbox-id/childFolders"):
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "clients-id",
                            "displayName": "Clients",
                            "parentFolderId": "inbox-id",
                            "childFolderCount": 1,
                        }
                    ]
                },
            )

        if request.url.path.endswith("/mailFolders/clients-id/childFolders"):
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "acme-id",
                            "displayName": "Acme",
                            "parentFolderId": "clients-id",
                            "childFolderCount": 0,
                        }
                    ]
                },
            )

        if request.url.path.endswith("/mailFolders/acme-id/messages") and request.method == "GET":
            assert request.url.params["$filter"] == (
                "isRead eq false and hasAttachments eq true and "
                "importance eq 'high' and categories/any(c:c eq 'Client') and "
                "flag/flagStatus eq 'flagged'"
            )
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "m-nav",
                            "subject": "Needs review",
                            "from": {"emailAddress": {"address": "client@example.com"}},
                            "sender": {"emailAddress": {"address": "assistant@example.com"}},
                            "replyTo": [{"emailAddress": {"address": "reply@example.com"}}],
                            "bodyPreview": "Please review",
                            "isDraft": False,
                            "isRead": False,
                            "hasAttachments": True,
                            "importance": "high",
                            "categories": ["Client"],
                            "flag": {"flagStatus": "flagged"},
                            "parentFolderId": "acme-id",
                            "internetMessageId": "<message@example.com>",
                            "conversationId": "conv-nav",
                        }
                    ]
                },
            )

        if request.url.path.endswith("/messages/m-nav/move"):
            return httpx.Response(
                200,
                json={"id": "m-nav", "subject": "Moved", "bodyPreview": "", "isDraft": False},
            )

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    tree = await graph.mail_folder_tree(maxDepth=3)
    assert tree.folders[0].childFolders[0].childFolders[0].path == "Inbox/Clients/Acme"

    resolved = await graph.resolve_mail_folder(folderPath="Inbox/Clients/Acme")
    assert resolved.folder.id == "acme-id"

    listed = await graph.list_messages(
        folderPath="Inbox/Clients/Acme",
        isRead=False,
        hasAttachments=True,
        importance="high",
        categories=["Client"],
        flagStatus="flagged",
    )
    message = listed.messages[0]
    assert message.sender == "assistant@example.com"
    assert message.replyTo == ["reply@example.com"]
    assert message.isRead is False
    assert message.hasAttachments is True
    assert message.categories == ["Client"]
    assert message.flagStatus == "flagged"
    assert message.parentFolderId == "acme-id"

    moved = await graph.move_message(
        messageId="m-nav",
        destinationFolder="Inbox/Clients/Acme",
    )
    assert moved.destinationFolderId == "acme-id"
    assert requests[-1][3] == {"destinationId": "acme-id"}

    await client.aclose()


@pytest.mark.anyio
async def test_folder_mutations_and_rules() -> None:
    requests: list[tuple[str, str, dict[str, object] | None]] = []

    def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content.decode("utf-8")) if request.content else None
        requests.append((request.method, request.url.path, body))

        if request.url.path.endswith("/mailFolders") and request.method == "GET":
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "inbox-id",
                            "displayName": "Inbox",
                            "childFolderCount": 1,
                        }
                    ]
                },
            )

        if request.url.path.endswith("/mailFolders/inbox-id/childFolders") and request.method == "GET":
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "clients-id",
                            "displayName": "Clients",
                            "parentFolderId": "inbox-id",
                            "childFolderCount": 0,
                        }
                    ]
                },
            )

        if request.url.path.endswith("/mailFolders/clients-id/childFolders") and request.method == "POST":
            assert body == {"displayName": "Acme"}
            return httpx.Response(
                201,
                json={
                    "id": "acme-id",
                    "displayName": "Acme",
                    "parentFolderId": "clients-id",
                    "childFolderCount": 0,
                },
            )

        if request.url.path.endswith("/mailFolders/acme-id") and request.method == "PATCH":
            assert body == {"displayName": "Acme Corp"}
            return httpx.Response(
                200,
                json={
                    "id": "acme-id",
                    "displayName": "Acme Corp",
                    "parentFolderId": "clients-id",
                    "childFolderCount": 0,
                },
            )

        if request.url.path.endswith("/mailFolders/acme-id") and request.method == "DELETE":
            return httpx.Response(204)

        if request.url.path.endswith("/mailFolders/inbox/messageRules") and request.method == "GET":
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "rule-1",
                            "displayName": "Clients",
                            "sequence": 1,
                            "isEnabled": True,
                            "actions": {"markAsRead": True},
                            "conditions": {"senderContains": ["client"]},
                        }
                    ]
                },
            )

        if request.url.path.endswith("/mailFolders/inbox/messageRules") and request.method == "POST":
            assert body == {
                "displayName": "Move Acme",
                "sequence": 2,
                "isEnabled": True,
                "conditions": {
                    "senderContains": ["acme"],
                    "subjectContains": ["Invoice"],
                },
                "actions": {
                    "moveToFolder": "clients-id",
                    "markAsRead": True,
                    "assignCategories": ["Client"],
                },
            }
            return httpx.Response(
                201,
                json={
                    "id": "rule-2",
                    "displayName": "Move Acme",
                    "sequence": 2,
                    "isEnabled": True,
                    "actions": body["actions"],
                    "conditions": body["conditions"],
                },
            )

        if request.url.path.endswith("/mailFolders/inbox/messageRules/rule-2") and request.method == "PATCH":
            assert body == {"displayName": "Move Acme invoices", "isEnabled": False}
            return httpx.Response(
                200,
                json={
                    "id": "rule-2",
                    "displayName": "Move Acme invoices",
                    "sequence": 2,
                    "isEnabled": False,
                },
            )

        if request.url.path.endswith("/mailFolders/inbox/messageRules/rule-2") and request.method == "DELETE":
            return httpx.Response(204)

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    created_folder = await graph.create_mail_folder(
        displayName="Acme",
        parentFolderPath="Inbox/Clients",
    )
    assert created_folder.folder.id == "acme-id"

    renamed_folder = await graph.rename_mail_folder(
        folderId="acme-id",
        displayName="Acme Corp",
    )
    assert renamed_folder.folder.displayName == "Acme Corp"

    deleted_folder = await graph.delete_mail_folder(folderId="acme-id")
    assert deleted_folder.deleted is True

    rules = await graph.list_mail_rules()
    assert rules.rules[0].displayName == "Clients"

    created_rule = await graph.create_mail_rule(
        displayName="Move Acme",
        sequence=2,
        senderContains=["acme"],
        subjectContains=["Invoice"],
        moveToFolderPath="Inbox/Clients",
        markAsRead=True,
        assignCategories=["Client"],
    )
    assert created_rule.rule.id == "rule-2"

    updated_rule = await graph.update_mail_rule(
        ruleId="rule-2",
        displayName="Move Acme invoices",
        isEnabled=False,
    )
    assert updated_rule.rule.isEnabled is False

    deleted_rule = await graph.delete_mail_rule(ruleId="rule-2")
    assert deleted_rule.deleted is True

    await client.aclose()


@pytest.mark.anyio
async def test_attachments_threads_and_categories() -> None:
    text_payload = base64.b64encode(b"hello attachment").decode("ascii")
    requests: list[tuple[str, str, dict[str, object] | None]] = []

    def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content.decode("utf-8")) if request.content else None
        requests.append((request.method, request.url.path, body))

        if request.url.path.endswith("/messages/msg-1/attachments") and request.method == "GET":
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "id": "att-1",
                            "name": "notes.txt",
                            "contentType": "text/plain",
                            "size": 16,
                            "isInline": False,
                        },
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "id": "att-inline",
                            "name": "pixel.png",
                            "contentType": "image/png",
                            "size": 10,
                            "isInline": True,
                        },
                    ]
                },
            )

        if request.url.path.endswith("/messages/msg-1/attachments/att-1"):
            return httpx.Response(
                200,
                json={
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "id": "att-1",
                    "name": "notes.txt",
                    "contentType": "text/plain",
                    "size": 16,
                    "contentBytes": text_payload,
                },
            )

        if request.url.path.endswith("/messages/msg-1/attachments/bin-1"):
            return httpx.Response(
                200,
                json={
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "id": "bin-1",
                    "name": "image.png",
                    "contentType": "image/png",
                    "size": 16,
                },
            )

        if request.url.path.endswith("/messages/msg-1/createReplyAll"):
            assert body == {
                "message": {
                    "body": {"contentType": "HTML", "content": "<p>Thanks</p>"}
                }
            }
            return httpx.Response(
                201,
                json={
                    "id": "reply-draft",
                    "subject": "Re: Hello",
                    "bodyPreview": "Thanks",
                    "isDraft": True,
                    "conversationId": "conv-1",
                },
            )

        if request.url.path.endswith("/messages") and "conversationId" in request.url.params.get("$filter", ""):
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "msg-1",
                            "subject": "Hello",
                            "bodyPreview": "First",
                            "isDraft": False,
                            "conversationId": "conv-1",
                        }
                    ]
                },
            )

        if request.url.path.endswith("/outlook/masterCategories") and request.method == "GET":
            return httpx.Response(
                200,
                json={"value": [{"id": "cat-1", "displayName": "Client", "color": "preset1"}]},
            )

        if request.url.path.endswith("/outlook/masterCategories") and request.method == "POST":
            return httpx.Response(
                201,
                json={"id": "cat-2", "displayName": body["displayName"], "color": body["color"]},
            )

        if request.url.path.endswith("/outlook/masterCategories/cat-2") and request.method == "PATCH":
            return httpx.Response(
                200,
                json={"id": "cat-2", "displayName": body["displayName"], "color": body["color"]},
            )

        if request.url.path.endswith("/outlook/masterCategories/cat-2") and request.method == "DELETE":
            return httpx.Response(204)

        if request.url.path.endswith("/messages/msg-1") and request.method == "PATCH":
            return httpx.Response(
                200,
                json={
                    "id": "msg-1",
                    "subject": "Hello",
                    "bodyPreview": "",
                    "isDraft": False,
                    "categories": body.get("categories", []),
                    "flag": body.get("flag", {}),
                    "isRead": body.get("isRead"),
                },
            )

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    attachments = await graph.list_attachments(messageId="msg-1")
    assert [attachment.id for attachment in attachments.attachments] == ["att-1"]

    content = await graph.get_attachment_content(messageId="msg-1", attachmentId="att-1")
    assert content.content == "hello attachment"

    binary = await graph.get_attachment_content(messageId="msg-1", attachmentId="bin-1")
    assert binary.content is None
    assert binary.unsupportedReason is not None

    thread = await graph.get_thread(conversationId="conv-1")
    assert thread.messages[0].id == "msg-1"

    reply = await graph.create_reply_draft(
        messageId="msg-1",
        comment="<p>Thanks</p>",
        replyAll=True,
    )
    assert reply.draft.id == "reply-draft"

    categories = await graph.list_categories()
    assert categories.categories[0].displayName == "Client"

    created = await graph.create_category(displayName="Prospect", color="preset2")
    assert created.category.displayName == "Prospect"

    updated = await graph.update_category(
        categoryId="cat-2",
        displayName="Customer",
        color="preset3",
    )
    assert updated.category.color == "preset3"

    deleted = await graph.delete_category(categoryId="cat-2")
    assert deleted.deleted is True

    categorized = await graph.set_message_categories(
        messageId="msg-1",
        categories=["Client"],
    )
    assert categorized.message.categories == ["Client"]

    read = await graph.mark_message_read(messageId="msg-1", isRead=True)
    assert read.message.isRead is True

    flagged = await graph.set_message_flag(messageId="msg-1", flagStatus="flagged")
    assert flagged.message.flagStatus == "flagged"

    await client.aclose()


@pytest.mark.anyio
async def test_pdf_attachment_text_extraction(monkeypatch: pytest.MonkeyPatch) -> None:
    class FakePage:
        def __init__(self, text: str) -> None:
            self._text = text

        def extract_text(self) -> str:
            return self._text

    class FakePdfReader:
        def __init__(self, stream: object) -> None:
            self.pages = [FakePage("First page"), FakePage("Second page")]

    monkeypatch.setattr(graph_module, "PdfReader", FakePdfReader)
    pdf_payload = base64.b64encode(b"%PDF fake content").decode("ascii")

    def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path.endswith("/messages/msg-1/attachments/pdf-1"):
            return httpx.Response(
                200,
                json={
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "id": "pdf-1",
                    "name": "brief.pdf",
                    "contentType": "application/pdf",
                    "size": 128,
                    "contentBytes": pdf_payload,
                },
            )

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    content = await graph.get_attachment_content(
        messageId="msg-1",
        attachmentId="pdf-1",
        maxChars=20,
    )

    assert content.encoding == "pdf-text"
    assert content.content == "--- Page 1 ---\nFirst"
    assert content.truncated is True
    assert "maxChars=20" in content.unsupportedReason

    await client.aclose()


@pytest.mark.anyio
async def test_contacts_crud_search_and_folders() -> None:
    requests: list[tuple[str, str, dict[str, str], dict[str, object] | None]] = []
    contact_categories = ["Client"]

    def contact(contact_id: str = "contact-1", name: str = "Ada Lovelace") -> dict[str, object]:
        return {
            "id": contact_id,
            "displayName": name,
            "givenName": "Ada",
            "surname": "Lovelace",
            "companyName": "Analytical Engines",
            "jobTitle": "Mathematician",
            "businessPhones": ["555-0100"],
            "mobilePhone": "555-0101",
            "emailAddresses": [{"address": "ada@example.com", "name": name}],
            "categories": contact_categories,
            "parentFolderId": "contacts-folder",
            "businessAddress": {
                "street": "1 Analytical Way",
                "city": "London",
                "state": "",
                "countryOrRegion": "UK",
                "postalCode": "NW1",
            },
            "homeAddress": {},
            "otherAddress": {
                "street": "PO Box 1",
                "city": "Seattle",
                "state": "WA",
                "countryOrRegion": "US",
                "postalCode": "98101",
            },
        }

    def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content.decode("utf-8")) if request.content else None
        requests.append((request.method, request.url.path, dict(request.url.params), body))

        if request.url.path.endswith("/contactFolders") and request.method == "GET":
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "contacts-folder",
                            "displayName": "VIP",
                            "childFolderCount": 0,
                        }
                    ]
                },
            )

        if request.url.path.endswith("/contacts") and request.method == "GET":
            assert "categories" in request.url.params["$select"]
            assert "businessAddress" in request.url.params["$select"]
            if "$filter" in request.url.params:
                assert "emailAddresses/any" in request.url.params["$filter"]
            return httpx.Response(200, json={"value": [contact()]})

        if request.url.path.endswith("/contacts/contact-1") and request.method == "GET":
            return httpx.Response(200, json=contact())

        if request.url.path.endswith("/contacts") and request.method == "POST":
            assert body["emailAddresses"] == [
                {"address": "ada@example.com", "name": "Ada Lovelace"}
            ]
            assert body["categories"] == ["Client", "VIP"]
            assert body["businessAddress"] == {
                "street": "1 Analytical Way",
                "city": "London",
                "countryOrRegion": "UK",
                "postalCode": "NW1",
            }
            return httpx.Response(201, json=contact("contact-new"))

        if request.url.path.endswith("/contacts/contact-1") and request.method == "PATCH":
            if "categories" in body:
                contact_categories[:] = body["categories"]
            return httpx.Response(
                200,
                json=contact("contact-1", body.get("displayName", "Ada Lovelace")),
            )

        if request.url.path.endswith("/contacts/contact-1") and request.method == "DELETE":
            return httpx.Response(204)

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    folders = await graph.list_contact_folders(mailbox="shared@example.com")
    assert folders.folders[0].displayName == "VIP"

    listed = await graph.list_contacts(mailbox="shared@example.com")
    assert listed.contacts[0].emailAddresses == ["ada@example.com"]
    assert listed.contacts[0].categories == ["Client"]
    assert listed.contacts[0].parentFolderId == "contacts-folder"
    assert listed.contacts[0].businessAddress.city == "London"
    assert listed.contacts[0].homeAddress is None
    assert listed.contacts[0].otherAddress.state == "WA"

    searched = await graph.search_contacts(
        mailbox="shared@example.com",
        query="ada@example.com",
    )
    assert searched.contacts[0].displayName == "Ada Lovelace"

    got = await graph.get_contact(
        mailbox="shared@example.com",
        contactId="contact-1",
    )
    assert got.contact.companyName == "Analytical Engines"

    created = await graph.create_contact(
        mailbox="shared@example.com",
        displayName="Ada Lovelace",
        emailAddresses=["ada@example.com"],
        categories=["Client", "VIP"],
        businessAddress={
            "street": "1 Analytical Way",
            "city": "London",
            "state": None,
            "countryOrRegion": "UK",
            "postalCode": "NW1",
        },
    )
    assert created.contact.id == "contact-new"

    updated = await graph.update_contact(
        mailbox="shared@example.com",
        contactId="contact-1",
        displayName="Ada Byron",
        homeAddress={"city": "Oxford", "countryOrRegion": "UK"},
        otherAddress={"street": "PO Box 1", "city": "Seattle"},
    )
    assert updated.contact.displayName == "Ada Byron"

    categorized = await graph.set_contact_categories(
        mailbox="shared@example.com",
        contactId="contact-1",
        categories=["Client", "VIP"],
    )
    assert categorized.contact.categories == ["Client", "VIP"]

    added = await graph.add_contact_categories(
        mailbox="shared@example.com",
        contactId="contact-1",
        categories=["Prospect", "Client"],
    )
    assert added.contact.categories == ["Client", "VIP", "Prospect"]

    removed = await graph.remove_contact_categories(
        mailbox="shared@example.com",
        contactId="contact-1",
        categories=["VIP"],
    )
    assert removed.contact.categories == ["Client", "Prospect"]

    cleared = await graph.clear_contact_categories(
        mailbox="shared@example.com",
        contactId="contact-1",
    )
    assert cleared.contact.categories == []

    deleted = await graph.delete_contact(
        mailbox="shared@example.com",
        contactId="contact-1",
        folderId="contacts-folder",
    )
    assert deleted.deleted is True

    assert any("/users/shared@example.com/" in path for _, path, _, _ in requests)
    assert any(
        path.endswith("/contactFolders/contacts-folder/contacts/contact-1")
        and method == "DELETE"
        for method, path, _, _ in requests
    )

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

        if request.url.path.endswith("/calendar/events/event-2") and request.method == "PATCH":
            assert body == {
                "subject": "Updated event",
                "start": {"dateTime": "2026-04-23T18:00:00", "timeZone": "UTC"},
                "end": {"dateTime": "2026-04-23T19:00:00", "timeZone": "UTC"},
                "attendees": [
                    {
                        "emailAddress": {"address": "new@example.com"},
                        "type": "required",
                    }
                ],
                "location": {"displayName": "Room 2"},
            }
            return httpx.Response(
                200,
                json={
                    "id": "event-2",
                    "subject": "Updated event",
                    "start": {"dateTime": "2026-04-23T18:00:00", "timeZone": "UTC"},
                    "end": {"dateTime": "2026-04-23T19:00:00", "timeZone": "UTC"},
                    "attendees": [],
                    "bodyPreview": "",
                    "body": {"contentType": "text", "content": ""},
                },
            )

        if request.url.path.endswith("/calendar/events/event-2") and request.method == "DELETE":
            return httpx.Response(204)

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

    updated = await graph.update_event(
        eventId="event-2",
        mailbox="shared@example.com",
        subject="Updated event",
        start="2026-04-23T18:00:00",
        end="2026-04-23T19:00:00",
        attendees=["new@example.com"],
        location="Room 2",
    )
    assert updated.event.subject == "Updated event"

    deleted = await graph.delete_event(
        eventId="event-2",
        mailbox="shared@example.com",
    )
    assert deleted.deleted is True

    with pytest.raises(RuntimeError, match="ErrorItemNotFound: No such message"):
        await graph.get_message(messageId="bad-id")

    await client.aclose()
