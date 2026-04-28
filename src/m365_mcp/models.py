from __future__ import annotations

from typing import Any

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
    requiredScopes: list[str] = Field(default_factory=list)
    grantedScopes: list[str] = Field(default_factory=list)
    missingScopes: list[str] = Field(default_factory=list)


class AuthStatusResult(MicrosoftConnectionStatus):
    localStatusUrl: str
    microsoftConnectUrl: str
    microsoftDisconnectUrl: str


class MessageSummary(AppModel):
    id: str
    subject: str
    from_: str | None = Field(default=None, alias="from")
    sender: str | None = None
    replyTo: list[str] = Field(default_factory=list)
    receivedDateTime: str | None = None
    sentDateTime: str | None = None
    bodyPreview: str
    webLink: str | None = None
    isDraft: bool
    isRead: bool | None = None
    hasAttachments: bool | None = None
    importance: str | None = None
    categories: list[str] = Field(default_factory=list)
    flagStatus: str | None = None
    parentFolderId: str | None = None
    internetMessageId: str | None = None
    conversationId: str | None = None


class MessageBody(AppModel):
    contentType: str
    content: str


class FullMessage(AppModel):
    id: str
    subject: str
    from_: str | None = Field(default=None, alias="from")
    sender: str | None = None
    replyTo: list[str] = Field(default_factory=list)
    to: list[str]
    cc: list[str]
    bcc: list[str]
    receivedDateTime: str | None = None
    sentDateTime: str | None = None
    bodyPreview: str
    body: MessageBody
    webLink: str | None = None
    isDraft: bool
    isRead: bool | None = None
    hasAttachments: bool | None = None
    importance: str | None = None
    categories: list[str] = Field(default_factory=list)
    flagStatus: str | None = None
    parentFolderId: str | None = None
    internetMessageId: str | None = None
    conversationId: str | None = None


class M365CapabilitiesResult(AppModel):
    content: str


class MailFolderInfo(AppModel):
    id: str
    displayName: str
    parentFolderId: str | None = None
    childFolderCount: int = 0
    totalItemCount: int = 0
    unreadItemCount: int = 0
    isHidden: bool | None = None
    path: str | None = None


class MailFolderTreeNode(MailFolderInfo):
    childFolders: list["MailFolderTreeNode"] = Field(default_factory=list)


class MailListResult(AppModel):
    mailbox: str
    folder: str
    folderId: str | None = None
    folderPath: str | None = None
    messages: list[MessageSummary]


class MailCheckInboxResult(AppModel):
    mailbox: str
    folder: MailFolderInfo
    messages: list[MessageSummary]


class MailListFoldersResult(AppModel):
    mailbox: str
    parentFolderId: str | None = None
    folders: list[MailFolderInfo]


class MailFolderTreeResult(AppModel):
    mailbox: str
    rootFolderId: str | None = None
    maxDepth: int
    folders: list[MailFolderTreeNode]


class MailResolveFolderResult(AppModel):
    mailbox: str
    folder: MailFolderInfo


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


class MailSendResult(AppModel):
    mailbox: str
    subject: str | None = None
    messageId: str | None = None
    replyAll: bool | None = None
    sent: bool


class MailMoveResult(AppModel):
    mailbox: str
    destinationFolder: str
    destinationFolderId: str | None = None
    destinationFolderPath: str | None = None
    movedMessage: MessageSummary


class MailFolderMutationResult(AppModel):
    mailbox: str
    folder: MailFolderInfo | None = None
    folderId: str | None = None
    deleted: bool = False


class MailRuleInfo(AppModel):
    id: str
    displayName: str
    sequence: int | None = None
    isEnabled: bool | None = None
    hasError: bool | None = None
    isReadOnly: bool | None = None
    conditions: dict[str, Any] = Field(default_factory=dict)
    actions: dict[str, Any] = Field(default_factory=dict)
    exceptions: dict[str, Any] = Field(default_factory=dict)


class MailListRulesResult(AppModel):
    mailbox: str
    rules: list[MailRuleInfo]


class MailRuleResult(AppModel):
    mailbox: str
    rule: MailRuleInfo | None = None
    ruleId: str | None = None
    deleted: bool = False


class AttachmentInfo(AppModel):
    id: str
    name: str | None = None
    contentType: str | None = None
    size: int | None = None
    isInline: bool = False
    lastModifiedDateTime: str | None = None
    attachmentType: str | None = None


class MailListAttachmentsResult(AppModel):
    mailbox: str
    messageId: str
    attachments: list[AttachmentInfo]


class MailAttachmentContentResult(AppModel):
    mailbox: str
    messageId: str
    attachment: AttachmentInfo
    content: str | None = None
    encoding: str | None = None
    truncated: bool = False
    unsupportedReason: str | None = None


class MailThreadResult(AppModel):
    mailbox: str
    conversationId: str
    messages: list[MessageSummary]


class MailCategoryInfo(AppModel):
    id: str | None = None
    displayName: str
    color: str | None = None


class MailListCategoriesResult(AppModel):
    mailbox: str
    categories: list[MailCategoryInfo]


class MailCategoryResult(AppModel):
    mailbox: str
    category: MailCategoryInfo | None = None
    categoryId: str | None = None
    deleted: bool = False


class MailUpdateMessageResult(AppModel):
    mailbox: str
    messageId: str
    message: MessageSummary


class ContactFolderInfo(AppModel):
    id: str
    displayName: str
    parentFolderId: str | None = None
    childFolderCount: int = 0


class ContactInfo(AppModel):
    id: str
    displayName: str | None = None
    givenName: str | None = None
    surname: str | None = None
    companyName: str | None = None
    jobTitle: str | None = None
    businessPhones: list[str] = Field(default_factory=list)
    mobilePhone: str | None = None
    emailAddresses: list[str] = Field(default_factory=list)


class ContactsListResult(AppModel):
    mailbox: str
    folderId: str | None = None
    contacts: list[ContactInfo]


class ContactsSearchResult(AppModel):
    mailbox: str
    query: str
    contacts: list[ContactInfo]


class ContactGetResult(AppModel):
    mailbox: str
    contact: ContactInfo


class ContactMutationResult(AppModel):
    mailbox: str
    contact: ContactInfo | None = None
    contactId: str | None = None
    deleted: bool = False


class ContactFoldersResult(AppModel):
    mailbox: str
    parentFolderId: str | None = None
    folders: list[ContactFolderInfo]


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


class CalendarUpdateEventResult(AppModel):
    mailbox: str
    event: CalendarEvent


class CalendarDeleteEventResult(AppModel):
    mailbox: str
    eventId: str
    deleted: bool


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
