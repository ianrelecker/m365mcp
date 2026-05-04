# M365 MCP

This lets Claude Desktop work with your Microsoft 365 mailbox, calendar, contacts, folders, rules, categories, and small readable attachments through Microsoft Graph.

It is a local MCP server. That means Claude starts it on your computer when Claude Desktop opens. You normally do not need to run the server manually.

## What Claude Can Do

Claude can:

- Check and search mail, including shared mailboxes you can access.
- Read message bodies, threads, categories, flags, read/unread state, subfolders, and small text/PDF attachments.
- Draft or send mail, reply in threads, move messages, and mark messages read/unread.
- Create, rename, delete, and navigate mail folders and subfolders.
- List, create, update, and delete Outlook Inbox rules.
- Search, create, update, and delete contacts.
- List, create, update, and delete calendar events.

Claude can also read [M365_MCP_CAPABILITIES.md](M365_MCP_CAPABILITIES.md) through the `m365_capabilities` tool or the `m365://capabilities` MCP resource.

## Setup Checklist

You need four things:

- `uv` installed on the computer running Claude Desktop.
- A Microsoft Entra app registration.
- A local `.env` file with your Microsoft app values.
- A Claude Desktop config entry for this MCP server.

## 1. Install uv

`uv` is the Python runner Claude will use to start this MCP server.

macOS or Linux:

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

Windows PowerShell:

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Close and reopen your terminal, then check that it installed:

```bash
uv --version
```

If Claude Desktop later cannot find `uv`, use the full path to `uv` in `claude_desktop_config.json`. On Windows, run:

```powershell
where uv
```

## 2. Create The Microsoft App

In the Azure Portal, create a Microsoft Entra app registration for local use.

Use these settings:

- Platform: `Web`
- Redirect URI: `http://localhost:8787/auth/microsoft/callback`
- Supported account type: `Accounts in this organizational directory only`
- Advanced setting: `Allow public client flows = No`

Create a client secret under `Certificates & secrets`. Copy the secret `Value`, not the `Secret ID`.

Add these delegated Microsoft Graph permissions:

- `Mail.ReadWrite`
- `Mail.ReadWrite.Shared`
- `Mail.Send`
- `Mail.Send.Shared`
- `Calendars.ReadWrite.Shared`
- `Contacts.ReadWrite.Shared`
- `MailboxSettings.ReadWrite`
- `openid`
- `profile`
- `email`
- `offline_access`

If your organization requires admin approval, click `Grant admin consent`.

## 3. Create Your .env File

Copy `.env.example` to `.env`, then fill in the values.

Important fields:

- `MICROSOFT_TENANT_ID`: the Azure tenant/directory ID.
- `MICROSOFT_CLIENT_ID`: the Azure app/client ID.
- `MICROSOFT_CLIENT_SECRET`: the client secret `Value`.
- `TOKEN_ENCRYPTION_KEY`: a base64 32-byte key used to encrypt the local token cache.
- `KNOWN_MAILBOXES`: optional comma-separated shared mailboxes, such as `shared@company.com`.
- `M365_AUDIT_LOG_ENABLED`: optional. Defaults to `true`.
- `M365_AUDIT_LOG_FILE`: optional. Defaults to `.audit/m365-mcp-audit.jsonl`.

Generate `TOKEN_ENCRYPTION_KEY` with:

```bash
python3 -c "import os, base64; print(base64.b64encode(os.urandom(32)).decode())"
```

On Windows, if `python3` is not available, try:

```powershell
python -c "import os, base64; print(base64.b64encode(os.urandom(32)).decode())"
```

## 4. Add It To Claude Desktop

Use [claude_desktop_config.json](claude_desktop_config.json) as the starting point. It keeps Claude's default `preferences` block and adds the `m365` MCP server.

If you are adding this server inside Codex instead of Claude Desktop, use [CODEX_MCP_SETUP.md](CODEX_MCP_SETUP.md). Codex has separate fields for command and arguments, so the setup is a little different.

Replace this example path:

```text
C:\Users\YOUR_WINDOWS_USER\Documents\m365mcp
```

with the real full path to this repo on your computer.

The config should look like this:

```json
{
  "preferences": {
    "coworkScheduledTasksEnabled": false,
    "coworkWebSearchEnabled": true,
    "ccdScheduledTasksEnabled": false
  },
  "mcpServers": {
    "m365": {
      "command": "uv",
      "args": [
        "--directory",
        "C:\\Users\\YOUR_WINDOWS_USER\\Documents\\m365mcp",
        "run",
        "--env-file",
        ".env",
        "mcp",
        "run",
        "src/m365_mcp/server.py"
      ]
    }
  }
}
```

Keep Microsoft secrets in `.env`. Do not paste tenant IDs, client secrets, or token keys directly into Claude's config.

After saving the config, fully quit and reopen Claude Desktop.

## 5. Connect Microsoft

Do not run the MCP server manually for normal use. Let Claude Desktop start it.

After Claude Desktop reopens, the local auth site should be available here:

```text
http://localhost:8787/
```

If that page is not available yet, open a Claude chat and ask:

```text
Check my Microsoft auth status with the m365 MCP server.
```

Claude should start the MCP server and call `auth_status`. The result includes the Microsoft connect URL.

To sign in directly, open:

```text
http://localhost:8787/auth/microsoft/start
```

Sign in with the Microsoft 365 account Claude should use. After sign-in, tokens are stored locally at `.tokens/microsoft-graph-token.json`, encrypted with `TOKEN_ENCRYPTION_KEY`.

If Claude says it is not authenticated, or `auth_status` shows missing scopes, open the same local auth link again and reconnect:

```text
http://localhost:8787/auth/microsoft/start
```

If `http://localhost:8787/` does not load, Claude probably did not start the MCP server. Open Claude Desktop settings, find the `m365` MCP server, click `View Logs`, and check the troubleshooting section below.

## Everyday Use

Once authenticated, ask Claude things like:

- `Check my Microsoft inbox.`
- `Search my M365 mail for invoices from Microsoft.`
- `Read this thread and draft a reply.`
- `Create a subfolder under Inbox called Clients.`
- `Move this message to Inbox/Clients/Acme.`
- `Create a rule that moves Acme invoices to that folder.`
- `Read the PDF attachment on this email.`
- `Create a calendar event for tomorrow at 2 PM.`

For shared mailboxes, mention the mailbox address in your request, for example:

```text
Check the shared@company.com inbox.
```

## Troubleshooting

If Claude shows `Server disconnected`, click `View Logs`.

Common fixes:

- If the logs say `TOKEN_ENCRYPTION_KEY must be a base64-encoded 32-byte key`, regenerate `TOKEN_ENCRYPTION_KEY` and update `.env`.
- If the MCP details still show placeholder values like `MICROSOFT_TENANT_ID=your-tenant-id`, remove any old environment-variable block from Claude's config and use `--env-file .env`.
- If the logs include `WinError 10048` or say the port is already in use, something else is using port `8787`. Stop the other process, then restart Claude Desktop.
- If the local auth page does not open, make sure Claude Desktop is running and the `m365` MCP server is enabled.
- If Claude cannot find `uv`, replace `"command": "uv"` with the full path from `where uv` on Windows or `which uv` on macOS/Linux.
- If Microsoft sign-in fails, confirm the Azure redirect URI exactly matches `http://localhost:8787/auth/microsoft/callback`.
- If `auth_status` reports `missingScopes`, add the missing permissions in Azure, grant consent if needed, then reconnect Microsoft.

## For Developers

You do not need these commands for normal Claude Desktop use.

Install dependencies:

```bash
uv sync
```

Run tests:

```bash
uv run pytest
```

Optional manual smoke test:

```bash
uv run mcp run src/m365_mcp/server.py
```

Stop the manual smoke test before opening Claude Desktop. Two copies cannot both use the same localhost helper port.

## Tool Reference

Mail and folders:

- `auth_status`
- `m365_capabilities`
- `mail_check_inbox`
- `mail_list`
- `mail_search`
- `mail_get`
- `mail_list_drafts`
- `mail_create_draft`
- `mail_send`
- `mail_send_draft`
- `mail_move`
- `mail_list_folders`
- `mail_folder_tree`
- `mail_resolve_folder`
- `mail_create_folder`
- `mail_rename_folder`
- `mail_delete_folder`

Attachments, threads, categories, and rules:

- `mail_list_attachments`
- `mail_get_attachment_content`
- `mail_get_thread`
- `mail_create_reply_draft`
- `mail_send_reply`
- `mail_list_categories`
- `mail_set_categories`
- `mail_add_categories`
- `mail_remove_categories`
- `mail_clear_categories`
- `mail_create_category`
- `mail_update_category`
- `mail_delete_category`
- `mail_mark_read`
- `mail_set_flag`
- `mail_list_rules`
- `mail_create_rule`
- `mail_update_rule`
- `mail_delete_rule`

Contacts and calendar:

- `contacts_list`
- `contacts_search`
- `contacts_get`
- `contacts_create`
- `contacts_update`
- `contacts_delete`
- `contacts_set_categories`
- `contacts_add_categories`
- `contacts_remove_categories`
- `contacts_clear_categories`
- `contacts_list_folders`
- `calendar_list_events`
- `calendar_create_event`
- `calendar_update_event`
- `calendar_delete_event`

## Contacts

- Use `contacts_search` to resolve recipients before drafting mail.
- Use `contacts_list_folders` when contacts are organized in folders.
- Contact tools can create, update, and delete Outlook contacts in mailboxes the signed-in user can access.
- Contacts include category names, parent folder IDs, and structured business/home/other addresses.
- Use `contacts_set_categories`, `contacts_add_categories`, `contacts_remove_categories`, and `contacts_clear_categories` to manage contact categories. These use the same Outlook master categories as mail.
- Use `contacts_update` with `businessAddress`, `homeAddress`, or `otherAddress` objects for street, city, state, country or region, and postal code changes.

## Security Notes

- `.env`, `.env.local`, and `.tokens/` are local-only files and are ignored by git.
- `.audit/` is local-only and ignored by git. It stores JSONL tool-call audit records for incident review.
- `TOKEN_ENCRYPTION_KEY` encrypts the saved Microsoft token cache at rest. If you rotate or lose it, delete `.tokens/microsoft-graph-token.json` and reconnect Microsoft.
- This server uses a confidential-client `Web` app registration, so `MICROSOFT_CLIENT_SECRET` is required.
- Microsoft sign-in uses authorization-code flow with PKCE. PKCE hardens the login code exchange, but it does not reduce Microsoft Graph permissions or replace token protection.
- Audit records include timestamp, tool name, outcome, mailbox, operation category, and key IDs such as message/event/folder/rule IDs when present.
- Audit records do not include access tokens, refresh tokens, client secrets, encryption keys, email bodies, attachment content, draft body text, calendar body text, or raw Microsoft Graph payloads.
- Treat `MICROSOFT_CLIENT_SECRET` like any other local credential and do not place it in shared configs or screenshots.
- This is for local MCP clients, not `claude.ai` remote connectors.
- No public HTTPS endpoint, IIS, WSL, Linux server, or public web server is required.
