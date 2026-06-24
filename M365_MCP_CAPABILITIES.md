# M365 MCP Capabilities

This MCP server gives Claude local delegated access to one Microsoft 365 account through Microsoft Graph. Use `auth_status` first. If `missingScopes` is not empty, reconnect Microsoft from the local helper URL before using tools that need those scopes.

## Mailboxes

- Leave `mailbox` blank for the signed-in user's mailbox.
- Pass a shared mailbox address, such as `shared@company.com`, when the signed-in user has delegated access.
- Use `mail_check_inbox` for the fastest inbox triage path. It defaults to unread messages in `Inbox`.
- Use `inferenceClassification="focused"` or `"other"` with `mail_check_inbox` or `mail_list` to separate Focused Inbox mail from Other mail before reading bodies.

## Folders And Subfolders

- Use `mail_folder_tree` or `mail_list_folders` before navigating unfamiliar mailboxes.
- Use folder IDs for precise operations. Folder display names can repeat under different parents.
- Use `mail_resolve_folder` for paths like `Inbox/Clients/Acme`.
- `mail_list` accepts `folderId`, `folderPath`, or legacy `folder`. If `folder` contains `/`, it is treated like a path.
- `mail_move` can move messages by `destinationFolderId`, `destinationFolderPath`, or legacy `destinationFolder`.
- Use `mail_create_folder` to create top-level folders or subfolders under `parentFolderId`/`parentFolderPath`.
- Use `mail_rename_folder` and `mail_delete_folder` by folder ID or resolved folder path.

## Message Triage

- Message summaries include read state, attachment presence, importance, categories, flag status, Focused/Other inference classification, parent folder ID, sender, reply-to, internet message ID, and conversation ID.
- `inferenceClassification` is a Microsoft Graph message property for Focused Inbox, not a separate app-registration permission.
- Prefer `mail_check_inbox` or `mail_list` filters for quick triage. `mail_search` delegates to Microsoft Graph `$search`, which can be slower on large mailboxes.
- Use `mail_mark_read` to mark mail read or unread.
- Use `mail_set_flag` to set follow-up status.
- Use category tools to set, add, remove, clear, or manage Outlook categories. `mail_update_category` can update a master category color, but Microsoft Graph does not support renaming an existing master category.

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
- Thread messages are sorted locally by received time when available to avoid Microsoft Graph's inefficient filtered-sort query path.
- Use `mail_create_reply_draft` to create reply or reply-all drafts in the thread.
- Use `mail_send_draft` only after the draft looks correct. Use `mail_send_reply` only when immediate sending is explicit.

## Contacts

- Use `contacts_search` to resolve recipients before drafting mail.
- Use `contacts_list_folders` when contacts are organized in folders.
- Contact tools can create, update, and delete Outlook contacts in mailboxes the signed-in user can access.
- Contact results include category names, parent folder IDs, Outlook Website values, personal notes, and structured business/home/other addresses.
- Use `personalHomePage` for Outlook Contacts `Other -> Website`; it is stored through the MAPI `PR_PERSONAL_HOME_PAGE` extended property (`String 0x3A50`).
- Use `contacts_update` to edit contact names, email addresses, phones, categories, Outlook Website, personal notes, and physical addresses.
- Use `contacts_delete` only after confirming the exact `contactId`; pass `folderId` when the contact lives outside the default contacts folder.
- Use `contacts_set_categories`, `contacts_add_categories`, `contacts_remove_categories`, and `contacts_clear_categories` for contact category changes.
- Mail and contacts share the same Outlook master categories, so use `mail_list_categories` or `mail_create_category` when the category definition itself needs to be inspected or created.

## Calendar

- Use `calendar_list_events` to inspect a time window.
- Use `calendar_create_event` to create calendar events in the signed-in or delegated mailbox calendar.
- Use `calendar_update_event` to edit subject, time, attendees, body, or location by event ID.
- Use `calendar_delete_event` to delete by event ID. Deleting organizer meetings may notify attendees.

## SharePoint And OneDrive Files

- These tools browse files and folders anywhere the signed-in user has access, without mounting anything locally.
- Use `sharepoint_search_items` first for "find this file or folder anywhere" — it searches across all SharePoint sites and OneDrive.
- Use `sharepoint_search_sites` then `sharepoint_list_drives` to go from a site name to its document libraries (each library is a "drive").
- Use `sharepoint_get_site` when you already know the site hostname and path, for example hostname `contoso.sharepoint.com` and sitePath `Acquisitions`.
- Use `sharepoint_list_children` to browse a folder by `driveId` plus an `itemId` or a path relative to the drive root; filter with `extensions` (e.g. `["xlsx", "pdf"]`) or `foldersOnly`.
- Use `sharepoint_search_in_drive` to search by name inside one library.
- Use `sharepoint_get_item_by_url` to turn a SharePoint/OneDrive sharing or browser URL into a `driveId` + `itemId`.
- The typical flow is: locate a workbook with these tools to get its `driveId` + `itemId`, then hand those to the workbook tools to edit it in place.

## Excel Workbooks

- These tools edit `.xlsx` workbooks stored in OneDrive/SharePoint **in place** via the Microsoft Graph Workbook API, so Excel's own engine applies the change. Formulas, formatting, and data-validation dropdowns are preserved, the file is never re-uploaded, and every change is versioned by SharePoint.
- Resolve a workbook first with `workbook_resolve` (by shareUrl, `driveId`+`itemId`, or `driveId`+`itemPath`); reuse the returned `driveId` + `itemId` on later calls.
- Use `workbook_list_worksheets` and `workbook_list_tables` to discover structure before reading or writing.
- Use `workbook_get_used_range` to find the data extent, and `workbook_get_range` to read a fixed range like `A1:O5`. Range reads return `values`, display `text`, the cell `formulas`, `numberFormat`, and the resolved `address`.
- Use `workbook_update_range` to write `values`, `formulas`, and/or `numberFormat` into a fixed range; the shape must match the address dimensions. `formulas` cells may be literal values or formula strings like `='Unit Mix'!H11` (cross-sheet references are fine).
- Use `workbook_add_table_row` to append rows to an Excel table; each row must match the table's column count and order.
- For date cells, write an Excel serial date number together with a date `numberFormat` (e.g. `mm/dd/yy`) so Excel stores a real date rather than text.
- For several related edits, open a session with `workbook_create_session` (`persistChanges=true`), pass the `sessionId` to each call, then `workbook_close_session`. Without a session, each write still persists individually.
- Workbook edits write to the stored file. Confirm the target workbook, worksheet, and range before writing, since changes apply immediately.

### Batch read/write, recalc, and the match/fill flow

- Use `workbook_get_ranges` (input: a list of `{worksheet, address}`) and `workbook_update_ranges` (input: a list of `{worksheet, address, formulas?, values?, numberFormat?}`) to read or write many scattered cells in one call. Both bundle Graph `$batch` sub-requests, auto-chunk to ≤20 per batch, preserve input order, and surface a per-range `error` without failing the whole batch.
- Use `workbook_calculate` (`calculationType` = `Recalculate` | `Full` | `FullRebuild`) to force a recalculation before reading computed cells back. The intended flow inside one `persistChanges=true` session is: read inputs → `workbook_update_ranges` (write mapped formulas) → `workbook_calculate` → `workbook_get_ranges` (read computed outputs, including dynamic-array spill cells).
- Use `workbook_list_names` (workbook-scoped, or worksheet-scoped when `worksheet` is given) and `workbook_get_name_range` to read defined names and resolve a name to its address + values/formulas.
- Use `workbook_clear_range` (`applyTo` = `Contents` | `Formats` | `All`) to clear cells, and `workbook_copy_range` (`copyType` = `All` | `Formulas` | `Values` | `Formats`, with a possibly cross-sheet `sourceRange`) to copy.
- Use `workbook_insert_range` (`shift` = `Down` | `Right`) to insert blank cells (e.g. insert-at-top trackers).
- Use `workbook_list_worksheets` to enumerate the tabs in a workbook. (Adding, copying, renaming, and deleting worksheets are not supported — these tools operate on cell ranges and tables in existing sheets.)
- Every range result includes its resolved `address`. Pass `sessionId` to any of these tools to run them inside an open session.

## Safety Notes

- The server acts as the signed-in user and can access shared resources only where that user already has permission.
- Prefer draft-first workflows for mail that will be sent.
- Prefer folder IDs after discovery to avoid accidentally acting on the wrong subfolder.
- Microsoft sign-in uses PKCE to harden the authorization-code exchange.
- Local JSONL audit logs are written by default to `.audit/m365-mcp-audit.jsonl`.
- Audit logs record tool metadata and key IDs only (message, folder, event, contact, rule, drive, item, and worksheet IDs where present); they do not include tokens, secrets, email bodies, attachment content, cell values, or raw Microsoft Graph payloads.
