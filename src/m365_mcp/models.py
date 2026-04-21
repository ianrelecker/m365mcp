from __future__ import annotations

from pydantic import BaseModel, ConfigDict, Field


class AppModel(BaseModel):
    model_config = ConfigDict(populate_by_name=True)


class AccountInfo(AppModel):
    name: str | None = None
    preferredUsername: str | None = None
    oid: str | None = None
    tid: str | None = None


class MicrosoftConnectionStatus(AppModel):
    connected: bool
    account: AccountInfo | None = None
    expiresAt: int | None = None
    knownMailboxes: list[str]


class AuthStatusResult(MicrosoftConnectionStatus):
    localStatusUrl: str
    microsoftConnectUrl: str
    microsoftDisconnectUrl: str


class MessageSummary(AppModel):
    id: str
    subject: str
    from_: str | None = Field(default=None, alias="from")
    receivedDateTime: str | None = None
    sentDateTime: str | None = None
    bodyPreview: str
    webLink: str | None = None
    isDraft: bool
    conversationId: str | None = None


class MessageBody(AppModel):
    contentType: str
    content: str


class FullMessage(AppModel):
    id: str
    subject: str
    from_: str | None = Field(default=None, alias="from")
    to: list[str]
    cc: list[str]
    bcc: list[str]
    receivedDateTime: str | None = None
    sentDateTime: str | None = None
    bodyPreview: str
    body: MessageBody
    webLink: str | None = None
    isDraft: bool
    importance: str | None = None
    conversationId: str | None = None


class MailListResult(AppModel):
    mailbox: str
    folder: str
    messages: list[MessageSummary]


class MailSearchResult(AppModel):
    mailbox: str
    query: str
    messages: list[MessageSummary]


class MailGetResult(AppModel):
    mailbox: str
    message: FullMessage


class MailListDraftsResult(AppModel):
    mailbox: str
    drafts: list[MessageSummary]


class MailCreateDraftResult(AppModel):
    mailbox: str
    draft: MessageSummary


class MailSendDraftResult(AppModel):
    mailbox: str
    messageId: str
    sent: bool


class MailMoveResult(AppModel):
    mailbox: str
    destinationFolder: str
    movedMessage: MessageSummary


class CalendarDateTime(AppModel):
    dateTime: str | None = None
    timeZone: str | None = None


class CalendarAttendee(AppModel):
    address: str | None = None
    name: str | None = None
    type: str | None = None
    response: str | None = None


class CalendarEvent(AppModel):
    id: str
    subject: str
    webLink: str | None = None
    start: CalendarDateTime
    end: CalendarDateTime
    location: str | None = None
    attendees: list[CalendarAttendee]
    bodyPreview: str
    body: MessageBody


class CalendarWindow(AppModel):
    start: str
    end: str


class CalendarListEventsResult(AppModel):
    mailbox: str
    window: CalendarWindow
    events: list[CalendarEvent]


class CalendarCreateEventResult(AppModel):
    mailbox: str
    event: CalendarEvent


class StoredMicrosoftTokens(AppModel):
    accessToken: str
    refreshToken: str
    expiresAt: int
    scope: str
    idToken: str | None = None
    account: AccountInfo | None = None
    updatedAt: int


class EncryptedPayload(AppModel):
    iv: str
    tag: str
    ciphertext: str

