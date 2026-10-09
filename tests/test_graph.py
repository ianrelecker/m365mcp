from __future__ import annotations

import base64
import json

import httpx
import pytest

from m365_mcp import microsoft_graph as graph_module
from m365_mcp import workbook_reader
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
            assert "inferenceClassification" in request.url.params["$select"]
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
                            "inferenceClassification": "focused",
                            "conversationId": "conv-1",
                        }
                    ]
                },
            )

        if request.url.path.endswith("/messages"):
            assert request.headers["consistencylevel"] == "eventual"
            assert request.url.params["$top"] == "50"
            assert request.url.params["$search"] == "\"from:\\\"boss\\\"\""
            assert "inferenceClassification" in request.url.params["$select"]
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "m2",
                            "subject": "Search hit",
                            "bodyPreview": "",
                            "inferenceClassification": "other",
                        }
                    ]
                },
            )

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    listed = await graph.list_messages(mailbox=None, folder="Inbox", top=999)
    assert listed.mailbox == "me"
    assert listed.messages[0].from_ == "sender@example.com"
    assert listed.messages[0].inferenceClassification == "focused"

    searched = await graph.search_messages(query='from:"boss"', top=99)
    assert searched.mailbox == "me"
    assert searched.messages[0].id == "m2"
    assert searched.messages[0].inferenceClassification == "other"

    await client.aclose()


@pytest.mark.anyio
async def test_get_message_and_list_drafts() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path.endswith("/messages/msg-123"):
            assert "inferenceClassification" in request.url.params["$select"]
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
                    "inferenceClassification": "focused",
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
    assert message.message.inferenceClassification == "focused"

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
        requests.append(
            (request.method, request.url.path, dict(request.url.params), body)
        )

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
            assert "inferenceClassification" in request.url.params["$select"]
            assert request.url.params["$filter"] == (
                "isRead eq false and hasAttachments eq true and "
                "importance eq 'high' and categories/any(c:c eq 'Client') and "
                "flag/flagStatus eq 'flagged' and "
                "inferenceClassification eq 'focused'"
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
                            "inferenceClassification": "focused",
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
        inferenceClassification="focused",
    )
    message = listed.messages[0]
    assert message.sender == "assistant@example.com"
    assert message.replyTo == ["reply@example.com"]
    assert message.isRead is False
    assert message.hasAttachments is True
    assert message.categories == ["Client"]
    assert message.flagStatus == "flagged"
    assert message.inferenceClassification == "focused"
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
    requests: list[tuple[str, str, dict[str, str], dict[str, object] | None]] = []

    def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content.decode("utf-8")) if request.content else None
        requests.append((request.method, request.url.path, dict(request.url.params), body))

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

        if request.url.path.endswith("/messages/msg-1") and request.method == "GET":
            assert request.url.params["$select"] == "conversationId"
            return httpx.Response(200, json={"conversationId": "conv-1"})

        if request.url.path.endswith(
            "/messages"
        ) and "conversationId" in request.url.params.get("$filter", ""):
            assert "$orderby" not in request.url.params
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "msg-2",
                            "subject": "Follow-up",
                            "bodyPreview": "Second",
                            "isDraft": False,
                            "conversationId": "conv-1",
                            "receivedDateTime": "2026-05-13T12:05:00Z",
                        },
                        {
                            "id": "msg-1",
                            "subject": "Hello",
                            "bodyPreview": "First",
                            "isDraft": False,
                            "conversationId": "conv-1",
                            "receivedDateTime": "2026-05-13T12:00:00Z",
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
            assert body == {"color": "preset3"}
            return httpx.Response(
                200,
                json={"id": "cat-2", "displayName": "Prospect", "color": body["color"]},
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

    thread_by_message = await graph.get_thread(messageId="msg-1")
    assert thread_by_message.conversationId == "conv-1"
    assert [message.id for message in thread_by_message.messages] == ["msg-1", "msg-2"]

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

    updated = await graph.update_category(categoryId="cat-2", color="preset3")
    assert updated.category.color == "preset3"

    with pytest.raises(ValueError, match="does not support renaming"):
        await graph.update_category(categoryId="cat-2", displayName="Customer")

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
async def test_attachment_images_and_inline_pictures() -> None:
    png_payload = base64.b64encode(b"\x89PNG\r\n\x1a\nlogo").decode("ascii")
    jpeg_bytes = b"\xff\xd8\xff\xe0jpeg-body"
    requested_paths: list[str] = []

    def handler(request: httpx.Request) -> httpx.Response:
        requested_paths.append(request.url.path)

        if request.url.path.endswith("/messages/msg-1/attachments"):
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "id": "inline-png",
                            "name": "logo.png",
                            "contentType": "image/png",
                            "size": 12,
                            "isInline": True,
                            "contentId": "logo@01D9",
                            "contentBytes": png_payload,
                        },
                        {
                            # Inline by content ID only, and Graph omitted
                            # contentBytes so the bytes come from /$value.
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "id": "inline-jpg",
                            "name": "signature.JPG",
                            "contentType": "application/octet-stream",
                            "size": len(jpeg_bytes),
                            "isInline": False,
                            "contentId": "sig@01D9",
                        },
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "id": "inline-huge",
                            "name": "banner.png",
                            "contentType": "image/png",
                            "size": 9_000_000,
                            "isInline": True,
                            "contentId": "banner@01D9",
                        },
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "id": "attached-png",
                            "name": "chart.png",
                            "contentType": "image/png",
                            "size": 12,
                            "isInline": False,
                            "contentBytes": png_payload,
                        },
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "id": "notes",
                            "name": "notes.txt",
                            "contentType": "text/plain",
                            "size": 4,
                            "isInline": False,
                        },
                    ]
                },
            )

        if request.url.path.endswith("/messages/msg-1/attachments/inline-jpg/$value"):
            return httpx.Response(200, content=jpeg_bytes)

        if request.url.path.endswith("/messages/msg-1/attachments/inline-png"):
            return httpx.Response(
                200,
                json={
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "id": "inline-png",
                    "name": "logo.png",
                    "contentType": "image/png",
                    "size": 12,
                    "isInline": True,
                    "contentId": "logo@01D9",
                    "contentBytes": png_payload,
                },
            )

        if request.url.path.endswith("/messages/msg-1/attachments/notes"):
            return httpx.Response(
                200,
                json={
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "id": "notes",
                    "name": "notes.txt",
                    "contentType": "text/plain",
                    "size": 4,
                },
            )

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    single = await graph.get_attachment_image(
        messageId="msg-1",
        attachmentId="inline-png",
    )
    assert single.unsupportedReason is None
    assert single.image is not None
    assert single.image.mimeType == "image/png"
    assert single.image.dataBase64 == png_payload
    assert single.image.byteSize == 12
    assert single.image.attachment.contentId == "logo@01D9"

    not_an_image = await graph.get_attachment_image(
        messageId="msg-1",
        attachmentId="notes",
    )
    assert not_an_image.image is None
    assert "not a PNG" in not_an_image.unsupportedReason

    inline = await graph.get_inline_images(messageId="msg-1")
    assert [image.attachment.id for image in inline.images] == [
        "inline-png",
        "inline-jpg",
    ]
    # Content type was generic, so the mime type comes from the file extension.
    assert inline.images[1].mimeType == "image/jpeg"
    assert base64.b64decode(inline.images[1].dataBase64) == jpeg_bytes
    assert [skipped.attachment.id for skipped in inline.skipped] == ["inline-huge"]
    assert "maxBytes" in inline.skipped[0].reason
    assert inline.truncated is False

    with_attachments = await graph.get_inline_images(
        messageId="msg-1",
        includeNonInline=True,
    )
    assert [image.attachment.id for image in with_attachments.images] == [
        "inline-png",
        "inline-jpg",
        "attached-png",
    ]

    capped = await graph.get_inline_images(messageId="msg-1", maxImages=1)
    assert [image.attachment.id for image in capped.images] == ["inline-png"]
    assert capped.truncated is True

    # The batch budget stops before a picture that would overflow it, even
    # though each picture on its own is under maxBytes.
    budgeted = await graph.get_inline_images(messageId="msg-1", maxTotalBytes=14)
    assert [image.attachment.id for image in budgeted.images] == ["inline-png"]
    assert budgeted.truncated is True

    # Text attachment reads point at the image tools instead of failing silently.
    image_via_text_tool = await graph.get_attachment_content(
        messageId="msg-1",
        attachmentId="inline-png",
    )
    assert image_via_text_tool.content is None
    assert "mail_get_attachment_image" in image_via_text_tool.unsupportedReason

    assert "/messages/msg-1/attachments/inline-huge/$value" not in requested_paths

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


def _single_page_pdf(text: bytes = b"INVOICE 12345", pages: int = 2) -> bytes:
    """Build a small valid PDF so rendering runs against real pdfium."""

    objects: list[bytes] = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"",  # placeholder for the page tree, filled in below
    ]
    kids: list[bytes] = []
    for page_index in range(pages):
        page_obj = len(objects) + 1
        content_obj = page_obj + 1
        kids.append(b"%d 0 R" % page_obj)
        objects.append(
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
            b"/Resources << /Font << /F1 %d 0 R >> >> /Contents %d 0 R >>"
            % (2 + pages * 2 + 1, content_obj)
        )
        stream = b"BT /F1 36 Tf 72 700 Td (%s p%d) Tj ET" % (text, page_index + 1)
        objects.append(
            b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream"
        )
    objects[1] = b"<< /Type /Pages /Kids [%s] /Count %d >>" % (b" ".join(kids), pages)
    objects.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")

    out = bytearray(b"%PDF-1.4\n")
    offsets: list[int] = []
    for number, body in enumerate(objects, start=1):
        offsets.append(len(out))
        out += b"%d 0 obj\n" % number + body + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objects) + 1)
    for offset in offsets:
        out += b"%010d 00000 n \n" % offset
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (
        len(objects) + 1,
        xref,
    )
    return bytes(out)


@pytest.mark.anyio
async def test_pdf_attachment_page_rendering() -> None:
    pdf_bytes = _single_page_pdf(pages=3)
    payload = base64.b64encode(pdf_bytes).decode("ascii")

    def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path.endswith("/messages/msg-1/attachments/pdf-1"):
            return httpx.Response(
                200,
                json={
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "id": "pdf-1",
                    "name": "invoice.pdf",
                    "contentType": "application/pdf",
                    "size": len(pdf_bytes),
                    "contentBytes": payload,
                },
            )

        if request.url.path.endswith("/messages/msg-1/attachments/notes"):
            return httpx.Response(
                200,
                json={
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "id": "notes",
                    "name": "notes.txt",
                    "contentType": "text/plain",
                    "size": 4,
                },
            )

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    rendered = await graph.get_attachment_pdf_pages(
        messageId="msg-1",
        attachmentId="pdf-1",
        maxPages=2,
    )
    assert rendered.unsupportedReason is None
    assert rendered.pageCount == 3
    assert [page.pageNumber for page in rendered.pages] == [1, 2]
    assert rendered.truncated is True
    first = rendered.pages[0]
    assert first.mimeType == "image/jpeg"
    assert first.byteSize > 0
    # Long edge honours the requested target, aspect ratio preserved.
    assert max(first.widthPx, first.heightPx) == 1600
    assert base64.b64decode(first.dataBase64)[:2] == b"\xff\xd8"

    paged = await graph.get_attachment_pdf_pages(
        messageId="msg-1",
        attachmentId="pdf-1",
        firstPage=3,
    )
    assert [page.pageNumber for page in paged.pages] == [3]
    assert paged.truncated is False

    smaller = await graph.get_attachment_pdf_pages(
        messageId="msg-1",
        attachmentId="pdf-1",
        maxPages=1,
        longEdge=400,
    )
    assert max(smaller.pages[0].widthPx, smaller.pages[0].heightPx) == 400
    assert smaller.pages[0].byteSize < first.byteSize

    # The byte budget still yields the first page rather than nothing.
    budgeted = await graph.get_attachment_pdf_pages(
        messageId="msg-1",
        attachmentId="pdf-1",
        maxTotalBytes=1,
    )
    assert [page.pageNumber for page in budgeted.pages] == [1]
    assert budgeted.truncated is True

    beyond_end = await graph.get_attachment_pdf_pages(
        messageId="msg-1",
        attachmentId="pdf-1",
        firstPage=99,
    )
    assert beyond_end.pages == []
    assert "has 3 pages" in beyond_end.unsupportedReason

    not_pdf = await graph.get_attachment_pdf_pages(
        messageId="msg-1",
        attachmentId="notes",
    )
    assert not_pdf.pages == []
    assert not_pdf.unsupportedReason == "Attachment is not a PDF"

    await client.aclose()


@pytest.mark.anyio
async def test_pdf_page_rendering_reports_damaged_file() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(
            200,
            json={
                "@odata.type": "#microsoft.graph.fileAttachment",
                "id": "pdf-bad",
                "name": "broken.pdf",
                "contentType": "application/pdf",
                "size": 9,
                "contentBytes": base64.b64encode(b"not a pdf").decode("ascii"),
            },
        )

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    result = await graph.get_attachment_pdf_pages(
        messageId="msg-1",
        attachmentId="pdf-bad",
    )
    assert result.pages == []
    assert "Could not open PDF" in result.unsupportedReason

    await client.aclose()


@pytest.mark.anyio
async def test_pdf_page_rendering_without_renderer(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(graph_module, "pypdfium2", None)

    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(
            200,
            json={
                "@odata.type": "#microsoft.graph.fileAttachment",
                "id": "pdf-1",
                "name": "invoice.pdf",
                "contentType": "application/pdf",
                "size": 128,
            },
        )

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    result = await graph.get_attachment_pdf_pages(
        messageId="msg-1",
        attachmentId="pdf-1",
    )
    assert result.pages == []
    assert "pypdfium2" in result.unsupportedReason

    await client.aclose()


@pytest.mark.anyio
async def test_contacts_crud_search_and_folders() -> None:
    requests: list[tuple[str, str, dict[str, str], dict[str, object] | None]] = []
    contact_categories = ["Client"]
    contact_display_name = ["Ada Lovelace"]

    def contact(contact_id: str = "contact-1", name: str = "Ada Lovelace") -> dict[str, object]:
        return {
            "id": contact_id,
            "displayName": name,
            "givenName": "Ada",
            "surname": "Lovelace",
            "companyName": "Analytical Engines",
            "jobTitle": "Mathematician",
            "personalNotes": "Analytical notes for client follow-up",
            "singleValueExtendedProperties": [
                {"id": "String 0x3A50", "value": "https://analytical.example"}
            ],
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
            assert request.url.params["$select"] == "id,displayName,parentFolderId"
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "contacts-folder",
                            "displayName": "VIP",
                        }
                    ]
                },
            )

        if request.url.path.endswith("/contacts") and request.method == "GET":
            assert "categories" in request.url.params["$select"]
            assert "businessAddress" in request.url.params["$select"]
            assert "businessHomePage" not in request.url.params["$select"]
            assert "personalNotes" in request.url.params["$select"]
            assert request.url.params["$expand"] == (
                "singleValueExtendedProperties($filter=id eq 'String 0x3A50')"
            )
            if "$filter" in request.url.params:
                assert "emailAddresses/any" in request.url.params["$filter"]
            return httpx.Response(200, json={"value": [contact()]})

        if request.url.path.endswith("/contacts/contact-1") and request.method == "GET":
            assert request.url.params["$expand"] == (
                "singleValueExtendedProperties($filter=id eq 'String 0x3A50')"
            )
            return httpx.Response(
                200, json=contact("contact-1", contact_display_name[0])
            )

        if request.url.path.endswith("/contacts/contact-new") and request.method == "GET":
            assert request.url.params["$expand"] == (
                "singleValueExtendedProperties($filter=id eq 'String 0x3A50')"
            )
            return httpx.Response(200, json=contact("contact-new"))

        if request.url.path.endswith("/contacts") and request.method == "POST":
            assert body["emailAddresses"] == [
                {"address": "ada@example.com", "name": "Ada Lovelace"}
            ]
            assert body["categories"] == ["Client", "VIP"]
            assert body["personalNotes"] == "Met at the analytics conference"
            assert body["singleValueExtendedProperties"] == [
                {"id": "String 0x3A50", "value": "https://ada.example"}
            ]
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
            if "displayName" in body:
                assert body["personalNotes"] == "Updated relationship notes"
                assert body["singleValueExtendedProperties"] == [
                    {"id": "String 0x3A50", "value": "https://byron.example"}
                ]
                contact_display_name[0] = body["displayName"]
            return httpx.Response(
                200,
                json=contact("contact-1", contact_display_name[0]),
            )

        if request.url.path.endswith("/contacts/contact-1") and request.method == "DELETE":
            return httpx.Response(204)

        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    folders = await graph.list_contact_folders(mailbox="shared@example.com")
    assert folders.folders[0].displayName == "VIP"
    assert folders.folders[0].childFolderCount == 0

    listed = await graph.list_contacts(mailbox="shared@example.com")
    assert listed.contacts[0].emailAddresses == ["ada@example.com"]
    assert listed.contacts[0].categories == ["Client"]
    assert listed.contacts[0].parentFolderId == "contacts-folder"
    assert listed.contacts[0].personalHomePage == "https://analytical.example"
    assert listed.contacts[0].personalNotes == "Analytical notes for client follow-up"
    assert listed.contacts[0].businessAddress.city == "London"
    assert listed.contacts[0].homeAddress is None
    assert listed.contacts[0].otherAddress.state == "WA"

    searched = await graph.search_contacts(
        mailbox="shared@example.com",
        query="ada@example.com",
    )
    assert searched.contacts[0].displayName == "Ada Lovelace"

    searched_note = await graph.search_contacts(
        mailbox="shared@example.com",
        query="follow-up",
    )
    assert searched_note.contacts[0].personalNotes == "Analytical notes for client follow-up"

    searched_home_page = await graph.search_contacts(
        mailbox="shared@example.com",
        query="analytical.example",
    )
    assert searched_home_page.contacts[0].personalHomePage == "https://analytical.example"

    got = await graph.get_contact(
        mailbox="shared@example.com",
        contactId="contact-1",
    )
    assert got.contact.companyName == "Analytical Engines"
    assert got.contact.personalHomePage == "https://analytical.example"

    created = await graph.create_contact(
        mailbox="shared@example.com",
        displayName="Ada Lovelace",
        emailAddresses=["ada@example.com"],
        categories=["Client", "VIP"],
        personalHomePage="https://ada.example",
        personalNotes="Met at the analytics conference",
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
        personalHomePage="https://byron.example",
        personalNotes="Updated relationship notes",
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


def _xlsx_bytes(
    build,
    *,
    cached: dict[str, str] | None = None,
    dimension: str | None = None,
) -> bytes:
    """Build an .xlsx with openpyxl, then patch the first sheet's XML.

    openpyxl never calculates, so ``cached`` maps a formula's text to the value
    Excel would have saved beside it; ``dimension`` overwrites the recorded
    used range to mimic a writer that left it stale.
    """

    import io
    import re
    import zipfile

    from openpyxl import Workbook

    workbook = Workbook()
    build(workbook)
    raw = io.BytesIO()
    workbook.save(raw)

    patched = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(raw.getvalue())) as source, zipfile.ZipFile(
        patched, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                xml = data.decode("utf-8")
                for formula, value in (cached or {}).items():
                    # openpyxl writes the empty cached value as <v /> or, when
                    # lxml is installed, as <v></v>.
                    xml = re.sub(
                        rf"<f>{re.escape(formula)}</f><v(?: />|></v>)",
                        f"<f>{formula}</f><v>{value}</v>",
                        xml,
                    )
                if dimension is not None:
                    xml = re.sub(r'<dimension ref="[^"]*" />', f'<dimension ref="{dimension}" />', xml)
                data = xml.encode("utf-8")
            target.writestr(info, data)
    return patched.getvalue()


def _build_deal_model(workbook) -> None:
    from datetime import date

    from openpyxl.workbook.defined_name import DefinedName

    sheet = workbook.active
    sheet.title = "Unit Mix"
    sheet["B2"] = "Units"
    sheet["C2"] = 120
    sheet["B3"] = "Cap Rate"
    sheet["C3"] = 0.055
    sheet["C3"].number_format = "0.0%"
    sheet["B4"] = "Doubled"
    sheet["C4"] = "=C2*2"
    sheet["B5"] = "Never calculated"
    sheet["C5"] = "=C2*3"
    sheet["B6"] = "Close"
    sheet["C6"] = date(2026, 1, 31)

    assumptions = workbook.create_sheet("Assumptions")
    assumptions["A1"] = "Purchase Price"
    assumptions["B1"] = 25_000_000

    hidden = workbook.create_sheet("Calc")
    hidden.sheet_state = "hidden"
    hidden["A1"] = "secret"

    workbook.defined_names["PurchasePrice"] = DefinedName(
        "PurchasePrice", attr_text="Assumptions!$B$1"
    )
    workbook.defined_names["IRR"] = DefinedName(
        "IRR", attr_text="'Unit Mix'!$C$3"
    )
    workbook.defined_names["Broken"] = DefinedName("Broken", attr_text="#REF!")


def _workbook_graph(
    payload: bytes,
    *,
    name: str = "model.xlsx",
    content_type: str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    odata_type: str = "#microsoft.graph.fileAttachment",
    size: int | None = None,
    downloads: list[int] | None = None,
) -> tuple[httpx.AsyncClient, MicrosoftGraphClient]:
    """Mock Graph: attachment metadata (``$select``, no contentBytes), then the
    raw bytes from ``/$value``. ``downloads`` records each download's size."""

    def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path.endswith("/messages/msg-1/attachments/xl-1"):
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
        if request.url.path.endswith("/messages/msg-1/attachments/xl-1/$value"):
            if downloads is not None:
                downloads.append(len(payload))
            return httpx.Response(200, content=payload)
        raise AssertionError(f"Unexpected request: {request.method} {request.url}")

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    return client, MicrosoftGraphClient(StaticAuthService(), client)


@pytest.mark.anyio
async def test_workbook_attachment_outline_lists_sheets_and_names() -> None:
    client, graph = _workbook_graph(_xlsx_bytes(_build_deal_model))

    result = await graph.get_attachment_workbook(messageId="msg-1", attachmentId="xl-1")

    assert result.unsupportedReason is None
    assert result.ranges == []
    assert [
        (sheet.name, sheet.visibility, sheet.dimensions) for sheet in result.sheets
    ] == [
        ("Unit Mix", "visible", "B2:C6"),
        ("Assumptions", "visible", "A1:B1"),
        ("Calc", "hidden", "A1"),
    ]
    assert result.sheets[0].rowCount == 5
    assert result.sheets[0].columnCount == 2
    # The outline counts names but lists none; Broken points at deleted cells.
    assert result.definedNames == []
    assert result.definedNameCount == 2
    assert result.definedNamesSkipped == 1

    named = await graph.get_attachment_workbook(
        messageId="msg-1", attachmentId="xl-1", includeDefinedNames=True
    )
    assert {(name.name, name.value) for name in named.definedNames} == {
        ("PurchasePrice", "Assumptions!$B$1"),
        ("IRR", "'Unit Mix'!$C$3"),
    }
    assert len(named.sheets) == 3

    await client.aclose()


@pytest.mark.anyio
async def test_workbook_attachment_reads_ranges_names_and_formulas() -> None:
    client, graph = _workbook_graph(
        _xlsx_bytes(_build_deal_model, cached={"C2*2": "240"})
    )

    result = await graph.get_attachment_workbook(
        messageId="msg-1",
        attachmentId="xl-1",
        ranges=[
            "'unit mix'!B2:C6",
            "PurchasePrice",
            "irr",
            "Calc!A:A",
            "Missing!A1",
            "B2",
        ],
        includeFormulas=True,
        includeNumberFormat=True,
    )

    unit_mix, price, irr, calc, missing, unqualified = result.ranges
    assert result.truncated is False

    assert unit_mix.worksheet == "Unit Mix"
    assert unit_mix.address == "B2:C6"
    assert (unit_mix.rowCount, unit_mix.columnCount) == (5, 2)
    assert unit_mix.values == [
        ["Units", 120],
        ["Cap Rate", 0.055],
        ["Doubled", 240],
        # Excel never saved a result for this formula, so there is none to read.
        ["Never calculated", None],
        ["Close", "2026-01-31T00:00:00"],
    ]
    assert unit_mix.formulas[2] == ["Doubled", "=C2*2"]
    assert unit_mix.formulas[3] == ["Never calculated", "=C2*3"]
    assert unit_mix.numberFormat[1][1] == "0.0%"

    assert (price.worksheet, price.address, price.values) == (
        "Assumptions",
        "B1",
        [[25_000_000]],
    )
    # A bare "IRR" is a defined name, not column IRR.
    assert (irr.worksheet, irr.address, irr.values) == ("Unit Mix", "C3", [[0.055]])
    # Whole columns clamp to the used range, and hidden sheets are readable.
    assert (calc.address, calc.values) == ("A1", [["secret"]])

    assert missing.error == "Workbook has no sheet named 'Missing'"
    assert missing.values is None
    assert "qualify the address with a sheet name" in unqualified.error

    await client.aclose()


@pytest.mark.anyio
async def test_workbook_attachment_sheet_mode_and_stale_dimension() -> None:
    def build(workbook) -> None:
        sheet = workbook.active
        sheet.title = "Rent Roll"
        for row in range(1, 6):
            sheet.cell(row=row, column=1, value=f"Unit {row}")
            sheet.cell(row=row, column=3, value=row * 1000)

    # The writer claims only A1 is used; the real data runs to C5.
    client, graph = _workbook_graph(_xlsx_bytes(build, dimension="A1"))

    result = await graph.get_attachment_workbook(
        messageId="msg-1",
        attachmentId="xl-1",
        sheet="Rent Roll",
        includeLayout=True,
    )

    assert result.sheets[0].dimensions == "A1:C5"
    (whole,) = result.ranges
    assert whole.address == "A1:C5"
    assert whole.values[4] == ["Unit 5", None, 5000]
    assert whole.formulas is None
    assert whole.numberFormat is None

    single = await graph.get_attachment_workbook(
        messageId="msg-1",
        attachmentId="xl-1",
        ranges=["C2:C3"],
    )
    # One sheet, so an unqualified address needs no sheet name.
    assert single.ranges[0].values == [[2000], [3000]]

    await client.aclose()


@pytest.mark.anyio
async def test_workbook_attachment_range_reads_skip_layout_unless_asked() -> None:
    def build(workbook) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        first = workbook.active
        first.title = "Summary"
        first["A1"] = "Price"
        first["B1"] = 1000
        for index in range(40):
            extra = workbook.create_sheet(f"Tab {index}")
            extra["A1"] = index
            workbook.defined_names[f"Input{index}"] = DefinedName(
                f"Input{index}", attr_text=f"'Tab {index}'!$A$1"
            )

    client, graph = _workbook_graph(_xlsx_bytes(build))

    narrow = await graph.get_attachment_workbook(
        messageId="msg-1", attachmentId="xl-1", ranges=["Summary!A1:B1"]
    )
    assert narrow.ranges[0].values == [["Price", 1000]]
    assert narrow.sheets == []
    assert narrow.definedNames == []
    assert len(narrow.model_dump_json()) < 1500

    by_sheet = await graph.get_attachment_workbook(
        messageId="msg-1", attachmentId="xl-1", sheet="Summary"
    )
    assert by_sheet.ranges[0].values == [["Price", 1000]]
    assert by_sheet.sheets == [] and by_sheet.definedNames == []

    with_layout = await graph.get_attachment_workbook(
        messageId="msg-1",
        attachmentId="xl-1",
        ranges=["Summary!A1:B1"],
        includeLayout=True,
    )
    assert len(with_layout.sheets) == 41
    assert with_layout.definedNameCount == 40
    assert with_layout.definedNames == []

    # The discovery call (no ranges, no sheet) still returns the layout.
    layout = await graph.get_attachment_workbook(messageId="msg-1", attachmentId="xl-1")
    assert len(layout.sheets) == 41 and layout.ranges == []

    await client.aclose()


@pytest.mark.anyio
async def test_workbook_attachment_names_drop_template_leftovers() -> None:
    def build(workbook) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        sheet = workbook.active
        sheet.title = "Assump"
        sheet["B2"] = 25_000_000
        sheet["B3"] = 0.055
        names = workbook.defined_names
        names["PurchasePrice"] = DefinedName("PurchasePrice", attr_text="Assump!$B$2")
        names["ExitCap"] = DefinedName("ExitCap", attr_text="Assump!$B$3")
        names["TaxRate"] = DefinedName("TaxRate", attr_text="0.012")
        # Leftovers broker templates carry by the hundred.
        for index in range(300):
            junk = "_" * 41 + f"a{index}"
            names[junk] = DefinedName(
                junk, attr_text='{"Assump",#N/A,TRUE,"Proforma";"Assump",#N/A,TRUE,"Proforma"}'
            )
        for index in range(150):
            names[f"OldInput{index}"] = DefinedName(f"OldInput{index}", attr_text="#N/A")
        names["Deleted"] = DefinedName("Deleted", attr_text="Assump!#REF!")
        names["Linked"] = DefinedName("Linked", attr_text="[1]Proforma!$C$4")
        names["RentColumn"] = DefinedName("RentColumn", attr_text="Rents[Rent]")
        names["Hidden"] = DefinedName("Hidden", attr_text="Assump!$B$2", hidden=True)
        names["wrn.Print_All."] = DefinedName(
            "wrn.Print_All.", attr_text='{#N/A,#N/A,FALSE,"Proforma"}'
        )
        sheet.print_area = "A1:B3"

    client, graph = _workbook_graph(_xlsx_bytes(build))

    outline = await graph.get_attachment_workbook(messageId="msg-1", attachmentId="xl-1")
    assert outline.definedNames == []
    assert outline.definedNameCount == 4
    # 300 + 150 leftovers plus Deleted, Linked, Hidden, and wrn.Print_All.;
    # openpyxl lifts the print area onto the sheet, so it is never a name.
    assert outline.definedNamesSkipped == 454
    assert len(outline.model_dump_json()) < 1500

    named = await graph.get_attachment_workbook(
        messageId="msg-1", attachmentId="xl-1", includeDefinedNames=True
    )
    assert [name.name for name in named.definedNames] == [
        "PurchasePrice",
        "ExitCap",
        "TaxRate",
        "RentColumn",
    ]
    assert len(named.model_dump_json()) < 2000

    # A name left off the list still resolves when asked for directly.
    with_cells = await graph.get_attachment_workbook(
        messageId="msg-1",
        attachmentId="xl-1",
        ranges=["Hidden", "PurchasePrice"],
        includeDefinedNames=True,
    )
    assert [r.values for r in with_cells.ranges] == [[[25_000_000]], [[25_000_000]]]
    assert len(with_cells.definedNames) == 4
    assert with_cells.sheets == []

    await client.aclose()


@pytest.mark.anyio
async def test_workbook_attachment_cell_budget_truncates() -> None:
    def build(workbook) -> None:
        sheet = workbook.active
        for row in range(1, 11):
            for column in range(1, 4):
                sheet.cell(row=row, column=column, value=row * column)

    client, graph = _workbook_graph(_xlsx_bytes(build))

    result = await graph.get_attachment_workbook(
        messageId="msg-1",
        attachmentId="xl-1",
        ranges=["A1:C10", "A1:B2"],
        maxCells=10,
    )

    first, second = result.ranges
    assert result.truncated is True
    assert first.truncated is True
    assert first.address == "A1:C3"
    assert first.rowCount == 3
    assert first.values[-1] == [3, 6, 9]
    # One cell of budget is left, which cannot hold a single row of the range.
    assert second.truncated is True
    assert "maxCells=10" in second.error

    await client.aclose()


@pytest.mark.anyio
@pytest.mark.parametrize(
    ("overrides", "reason"),
    [
        ({"name": "legacy.xls", "content_type": "application/vnd.ms-excel"}, "Legacy .xls"),
        ({"name": "binary.xlsb", "content_type": "application/octet-stream"}, "Legacy .xls"),
        ({"name": "notes.txt", "content_type": "text/plain"}, "not an .xlsx"),
        ({"size": 50_000_000}, "maxBytes=10000000"),
        (
            {"odata_type": "#microsoft.graph.itemAttachment"},
            "Only file attachments",
        ),
    ],
)
async def test_workbook_attachment_unsupported(overrides: dict, reason: str) -> None:
    client, graph = _workbook_graph(_xlsx_bytes(_build_deal_model), **overrides)

    result = await graph.get_attachment_workbook(messageId="msg-1", attachmentId="xl-1")

    assert reason in result.unsupportedReason
    assert result.sheets == []

    await client.aclose()


@pytest.mark.anyio
async def test_workbook_attachment_rejects_non_zip_payload() -> None:
    # A password-protected workbook is an OLE container, not a zip.
    client, graph = _workbook_graph(b"\xd0\xcf\x11\xe0 not a zip", name="locked.xlsx")

    result = await graph.get_attachment_workbook(messageId="msg-1", attachmentId="xl-1")

    assert "password-protected" in result.unsupportedReason

    await client.aclose()


@pytest.mark.anyio
async def test_workbook_attachment_without_openpyxl(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(workbook_reader, "openpyxl", None)
    downloads: list[int] = []
    client, graph = _workbook_graph(_xlsx_bytes(_build_deal_model), downloads=downloads)

    result = await graph.get_attachment_workbook(messageId="msg-1", attachmentId="xl-1")

    assert result.unsupportedReason == "Workbook reading requires the openpyxl package"
    assert downloads == []

    await client.aclose()


@pytest.mark.anyio
async def test_attachment_content_points_workbooks_to_workbook_tool() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(
            200,
            json={
                "@odata.type": "#microsoft.graph.fileAttachment",
                "id": "xl-1",
                "name": "model.xlsx",
                "contentType": "application/octet-stream",
                "size": 10,
            },
        )

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    graph = MicrosoftGraphClient(StaticAuthService(), client)

    result = await graph.get_attachment_content(messageId="msg-1", attachmentId="xl-1")

    assert result.content is None
    assert "mail_get_attachment_workbook" in result.unsupportedReason

    await client.aclose()


@pytest.mark.anyio
async def test_workbook_attachment_counts_uncalculated_formulas_in_used_range() -> None:
    def build(workbook) -> None:
        sheet = workbook.active
        sheet["A1"] = "Heading"
        sheet["A2"] = "=1+1"

    # Excel never calculated A2, so it has no cached value.
    client, graph = _workbook_graph(_xlsx_bytes(build))

    result = await graph.get_attachment_workbook(
        messageId="msg-1",
        attachmentId="xl-1",
        sheet="Sheet",
        includeFormulas=True,
        includeLayout=True,
    )

    assert result.sheets[0].dimensions == "A1:A2"
    (whole,) = result.ranges
    assert whole.address == "A1:A2"
    assert whole.values == [["Heading"], [None]]
    assert whole.formulas == [["Heading"], ["=1+1"]]

    await client.aclose()


@pytest.mark.anyio
async def test_workbook_attachment_isolates_unreadable_names() -> None:
    def build(workbook) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        sheet = workbook.active
        sheet["A1"] = "Rent"
        workbook.defined_names["RentColumn"] = DefinedName(
            "RentColumn", attr_text="Rents[Rent]"
        )
        workbook.defined_names["TaxRate"] = DefinedName("TaxRate", attr_text="0.012")

    client, graph = _workbook_graph(_xlsx_bytes(build))

    result = await graph.get_attachment_workbook(
        messageId="msg-1",
        attachmentId="xl-1",
        ranges=["RentColumn", "A1", "TaxRate"],
    )

    assert result.unsupportedReason is None
    structured, cell, constant = result.ranges
    assert "Rents[Rent]" in structured.error
    assert cell.values == [["Rent"]]
    assert "not a single cell range" in constant.error

    await client.aclose()


@pytest.mark.anyio
async def test_workbook_attachment_resolves_sheet_qualified_names() -> None:
    def build(workbook) -> None:
        from openpyxl.workbook.defined_name import DefinedName

        _build_deal_model(workbook)
        unit_mix = workbook["Unit Mix"]
        unit_mix.defined_names["Units"] = DefinedName(
            "Units", attr_text="'Unit Mix'!$C$2"
        )

    client, graph = _workbook_graph(_xlsx_bytes(build))

    result = await graph.get_attachment_workbook(
        messageId="msg-1",
        attachmentId="xl-1",
        ranges=["'Unit Mix'!IRR", "'Unit Mix'!Units", "Assumptions!Units"],
    )

    irr, units, wrong_sheet = result.ranges
    # Not column IRR: the workbook name IRR.
    assert (irr.worksheet, irr.address, irr.values) == ("Unit Mix", "C3", [[0.055]])
    assert (units.address, units.values) == ("C2", [[120]])
    # A sheet-scoped name is only visible from its own sheet.
    assert "not an A1 address or defined name" in wrong_sheet.error

    await client.aclose()


@pytest.mark.anyio
async def test_workbook_attachment_pads_ranges_past_the_sheet_end() -> None:
    def build(workbook) -> None:
        sheet = workbook.active
        for row in range(1, 4):
            sheet.cell(row=row, column=1, value=row)

    client, graph = _workbook_graph(_xlsx_bytes(build))

    result = await graph.get_attachment_workbook(
        messageId="msg-1",
        attachmentId="xl-1",
        ranges=["A2:B5"],
        includeFormulas=True,
        includeNumberFormat=True,
    )

    (padded,) = result.ranges
    assert (padded.address, padded.rowCount, padded.columnCount) == ("A2:B5", 4, 2)
    assert padded.values == [[2, None], [3, None], [None, None], [None, None]]
    assert len(padded.formulas) == 4 and len(padded.numberFormat) == 4

    await client.aclose()


@pytest.mark.anyio
async def test_workbook_attachment_stops_download_past_max_bytes() -> None:
    # Graph reports a small size, but the download is larger than maxBytes.
    client, graph = _workbook_graph(b"x" * 5_000, size=100)

    result = await graph.get_attachment_workbook(
        messageId="msg-1", attachmentId="xl-1", maxBytes=1_000
    )

    assert result.unsupportedReason == "Workbook size exceeds maxBytes=1000"

    await client.aclose()


@pytest.mark.anyio
async def test_workbook_attachment_clamps_limits() -> None:
    client, graph = _workbook_graph(_xlsx_bytes(_build_deal_model), size=50_000_000)

    oversized = await graph.get_attachment_workbook(
        messageId="msg-1", attachmentId="xl-1", maxBytes=10**12
    )
    assert oversized.unsupportedReason == "Workbook size exceeds maxBytes=25000000"
    await client.aclose()

    client, graph = _workbook_graph(_xlsx_bytes(_build_deal_model))
    starved = await graph.get_attachment_workbook(
        messageId="msg-1", attachmentId="xl-1", ranges=["'Unit Mix'!B2:C2"], maxCells=0
    )
    assert "maxCells=1 is used up" in starved.ranges[0].error
    await client.aclose()


def _rezip(payload: bytes, replace: dict[str, bytes]) -> bytes:
    import io
    import zipfile

    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(payload)) as source, zipfile.ZipFile(
        out, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for info in source.infolist():
            target.writestr(info, replace.pop(info.filename, source.read(info.filename)))
        for filename, data in replace.items():
            target.writestr(filename, data)
    return out.getvalue()


@pytest.mark.anyio
async def test_workbook_attachment_refuses_dtds_and_zip_bombs() -> None:
    import io
    import zipfile

    base = _xlsx_bytes(_build_deal_model)
    with zipfile.ZipFile(io.BytesIO(base)) as archive:
        sheet_xml = archive.read("xl/worksheets/sheet1.xml")
    entity = b'<!DOCTYPE worksheet [<!ENTITY a "aaaaaaaa">]>' + sheet_xml

    client, graph = _workbook_graph(_rezip(base, {"xl/worksheets/sheet1.xml": entity}))
    result = await graph.get_attachment_workbook(messageId="msg-1", attachmentId="xl-1")
    assert "document type declaration" in result.unsupportedReason
    await client.aclose()

    # 20 MB of zeros deflates to ~20 KB: far past any real workbook's ratio.
    bomb = _rezip(base, {"xl/media/padding.bin": b"\0" * 20_000_000})
    client, graph = _workbook_graph(bomb)
    result = await graph.get_attachment_workbook(messageId="msg-1", attachmentId="xl-1")
    assert "expands to more than" in result.unsupportedReason
    await client.aclose()
