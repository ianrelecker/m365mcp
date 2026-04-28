# M365 MCP Capabilities

This MCP server gives Claude local delegated access to one Microsoft 365 account through Microsoft Graph. Use `auth_status` first. If `missingScopes` is not empty, reconnect Microsoft from the local helper URL before using tools that need those scopes.

## Mailboxes

- Leave `mailbox` blank for the signed-in user's mailbox.
- Pass a shared mailbox address, such as `shared@company.com`, when the signed-in user has delegated access.
- Use `mail_check_inbox` for the fastest inbox triage path. It defaults to unread messages in `Inbox`.

## Folders And Subfolders

- Use `mail_folder_tree` or `mail_list_folders` before navigating unfamiliar mailboxes.
- Use folder IDs for precise operations. Folder display names can repeat under different parents.
- Use `mail_resolve_folder` for paths like `Inbox/Clients/Acme`.
- `mail_list` accepts `folderId`, `folderPath`, or legacy `folder`. If `folder` contains `/`, it is treated like a path.
- `mail_move` can move messages by `destinationFolderId`, `destinationFolderPath`, or legacy `destinationFolder`.
- Use `mail_create_folder` to create top-level folders or subfolders under `parentFolderId`/`parentFolderPath`.
- Use `mail_rename_folder` and `mail_delete_folder` by folder ID or resolved folder path.

## Message Triage

- Message summaries include read state, attachment presence, importance, categories, flag status, parent folder ID, sender, reply-to, internet message ID, and conversation ID.
- Use `mail_mark_read` to mark mail read or unread.
- Use `mail_set_flag` to set follow-up status.
- Use category tools to set, add, remove, clear, or manage Outlook categories.

## Sending

- Use `mail_create_draft` plus `mail_send_draft` when the user has not explicitly approved the exact final email.
- Use `mail_send` only when the user clearly asked to send a new message now.
- Use `mail_create_reply_draft` for safer thread replies, or `mail_send_reply` only when the user clearly approved sending the reply now.

## Rules

- Use `mail_list_rules` to inspect existing Inbox rules before changing them.
- Use `mail_create_rule`, `mail_update_rule`, and `mail_delete_rule` for Outlook Inbox rules.
- Rule tools accept raw Microsoft Graph `conditions`, `actions`, and `exceptions`, plus convenience fields like `senderContains`, `subjectContains`, `moveToFolderPath`, `markAsRead`, and `assignCategories`.
- Resolve or create destination folders before making move rules. Prefer `moveToFolderId` after discovery.

## Attachments

- Use `mail_list_attachments` before reading attachment content.
- `mail_get_attachment_content` returns content for small text-like files and extracts text from small PDFs.
- PDF extraction is text-only. Scanned/image-only PDFs need OCR and return `unsupportedReason`.
- Large, binary, item, or reference attachments return metadata with `unsupportedReason`.
- The server does not save attachments to disk.

## Threads And Replies

- Use `mail_get_thread` with either `messageId` or `conversationId` to inspect a conversation.
- Use `mail_create_reply_draft` to create reply or reply-all drafts in the thread.
- Use `mail_send_draft` only after the draft looks correct. Use `mail_send_reply` only when immediate sending is explicit.

## Contacts

- Use `contacts_search` to resolve recipients before drafting mail.
- Use `contacts_list_folders` when contacts are organized in folders.
- Contact tools can create, update, and delete Outlook contacts in mailboxes the signed-in user can access.

## Calendar

- Use `calendar_list_events` to inspect a time window.
- Use `calendar_create_event` to create calendar events in the signed-in or delegated mailbox calendar.
- Use `calendar_update_event` to edit subject, time, attendees, body, or location by event ID.
- Use `calendar_delete_event` to delete by event ID. Deleting organizer meetings may notify attendees.

## Safety Notes

- The server acts as the signed-in user and can access shared resources only where that user already has permission.
- Prefer draft-first workflows for mail that will be sent.
- Prefer folder IDs after discovery to avoid accidentally acting on the wrong subfolder.
