# M365 MCP

Local MCP server for one Microsoft 365 user. It lets Claude call Microsoft Graph with that user's delegated permissions, including shared or delegated mailboxes and calendars the user already has access to.

This Python implementation runs with `uv` and the official MCP Python SDK. It starts two things in one process:

- a stdio MCP server for Claude/Desktop-style clients
- a localhost-only helper web app for Microsoft OAuth and status checks

## What it exposes

- `m365_capabilities`
- `auth_status`
- `mail_check_inbox`
- `mail_list_folders`
- `mail_folder_tree`
- `mail_resolve_folder`
- `mail_list`
- `mail_search`
- `mail_get`
- `mail_list_drafts`
- `mail_create_draft`
- `mail_send_draft`
- `mail_move`
- `mail_list_attachments`
- `mail_get_attachment_content`
- `mail_get_thread`
- `mail_create_reply_draft`
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
- `contacts_list`
- `contacts_search`
- `contacts_get`
- `contacts_create`
- `contacts_update`
- `contacts_delete`
- `contacts_list_folders`
- `calendar_list_events`
- `calendar_create_event`

Every mail/calendar/contact tool accepts an optional `mailbox` argument. Leave it blank for the signed-in user's own mailbox. Set it to something like `shared@company.com` for a shared or delegated mailbox/calendar/contact folder.
The model can also read [M365_MCP_CAPABILITIES.md](M365_MCP_CAPABILITIES.md) through the `m365_capabilities` tool or the `m365://capabilities` MCP resource.

## 1. Microsoft Entra setup

Create a single-tenant app registration for local use.

Recommended settings for this implementation:

- Platform: `Web`
- Redirect URI: `http://localhost:8787/auth/microsoft/callback`
- Supported account type: `Accounts in this organizational directory only`
- Advanced setting: `Allow public client flows = No`

Create a client secret under `Certificates & secrets`, then copy the secret `Value` into `.env`. Do not use the `Secret ID`; that looks like a GUID and will fail during token exchange.

Add delegated Microsoft Graph permissions:

- `Mail.ReadWrite`
- `Mail.ReadWrite.Shared`
- `Mail.Send.Shared`
- `Mail.Send` (optional, but reasonable if you also want an explicit non-shared send scope)
- `Calendars.ReadWrite.Shared`
- `Contacts.ReadWrite.Shared`
- `MailboxSettings.ReadWrite`
- `openid`
- `profile`
- `email`
- `offline_access`

If your tenant restricts user consent, grant admin consent for the enterprise app.

If you are upgrading from an older version of this server, reconnect Microsoft after adding the new permissions. `auth_status` reports `requiredScopes`, `grantedScopes`, and `missingScopes` so Claude can tell when the local token needs a fresh consent.

## 2. Local config

Copy `.env.example` to `.env` and fill in the values.

Important values:

- `LOCAL_BASE_URL`
  Optional. Defaults to `http://localhost:8787`.
- `MICROSOFT_CLIENT_SECRET`
  Required. Use the Azure client secret `Value`, not the `Secret ID`.
- `TOKEN_ENCRYPTION_KEY`
  Base64-encoded 32 byte key used for encrypted Microsoft token storage.

Generate `TOKEN_ENCRYPTION_KEY` with:

```bash
python3 -c "import os, base64; print(base64.b64encode(os.urandom(32)).decode())"
```

## 3. Install and run

If `uv` is not installed, install it first.

macOS/Linux:

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

Windows PowerShell:

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Close and reopen the terminal after installation if `uv` is not found, then verify:

```bash
uv --version
```

Install dependencies with `uv`:

```bash
uv sync
```

For a quick manual smoke test, run the stdio MCP server:

```bash
uv run mcp run src/m365_mcp/server.py
```

You can also run it directly without the MCP CLI:

```bash
uv run python -m m365_mcp.server
```

When the server starts, the local helper page is available at:

```text
http://localhost:8787/
```

Stop this manual command before starting Claude Desktop. Claude launches its own MCP server process, and two copies cannot both own the local helper port.

## 4. Connect Microsoft

Open:

```text
http://localhost:8787/auth/microsoft/start
```

Sign in with the one Microsoft 365 user you want this server to act as.

After this, the server keeps that user's Microsoft refresh token in `.tokens/microsoft-graph-token.json`, encrypted with `TOKEN_ENCRYPTION_KEY`.

## 5. Install into Claude Desktop

Use the MCP CLI to install the server into Claude Desktop:

```bash
uv run mcp install src/m365_mcp/server.py -f .env --name "m365"
```

If you prefer to configure Claude Desktop manually, use [claude_desktop_config.json](claude_desktop_config.json) as a starting point. It preserves Claude's default `preferences` block and adds the `m365` MCP server.

In that file, replace `C:\\Users\\YOUR_WINDOWS_USER\\Documents\\m365mcp` with the absolute path to this repo. The `--directory` argument is important because Claude may launch `uv` from another working directory, and `uv` needs to find this repo's `pyproject.toml`.

If Claude cannot find `uv`, replace `"command": "uv"` with the full path from `where uv` on Windows.

Keep Microsoft credentials in `.env`; the sample config uses `uv run --env-file .env` so Claude does not need those secrets duplicated in `claude_desktop_config.json`.

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

If Claude shows `Server disconnected`, click `View Logs`. If the MCP details still show environment variables like `MICROSOFT_TENANT_ID=your-tenant-id` or `TOKEN_ENCRYPTION_KEY=your-base64-32-byte-key`, the config is still using the old placeholder env block. Remove that block, use `--env-file .env`, restart Claude Desktop, then open `http://localhost:8787/` to confirm the helper page is running.

If the logs include `WinError 10048` or say that only one usage of the socket address is permitted, another process is already using port `8787`. This usually means the manual `uv run mcp run ...` smoke test is still running. Stop the manual process, then restart Claude Desktop. Only change `PORT`, `LOCAL_BASE_URL`, and the Azure redirect URI if you intentionally want to use a different localhost port.

## 6. Run tests

Install dev dependencies and run the test suite:

```bash
uv sync --extra dev
uv run pytest
```

## Windows Notes

- This project works on Windows. If Claude Desktop is running on Windows, this is a fine target.
- Install `uv` from PowerShell with `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`, then reopen the terminal and run `uv --version`.
- `uv` can manage the required Python version for you.
- The Microsoft redirect URI stays the same on Windows: `http://localhost:8787/auth/microsoft/callback`.
- Keep the helper app bound to localhost. You do not need IIS, Linux, WSL, or a public web server for local MCP use.

## Security Notes

- `.env`, `.env.local`, and `.tokens/` are local-only files and are ignored by git.
- `TOKEN_ENCRYPTION_KEY` encrypts the saved Microsoft token cache at rest. If you rotate or lose it, delete `.tokens/microsoft-graph-token.json` and reconnect Microsoft.
- This server currently uses a confidential-client `Web` app registration, so `MICROSOFT_CLIENT_SECRET` is required.
- Treat `MICROSOFT_CLIENT_SECRET` like any other local credential and do not place it in shared configs or screenshots.

## Notes

- This is intentionally scoped for one user, not multi-tenant SaaS.
- This is for local MCP clients, not `claude.ai` remote connectors.
- Shared/delegated mailbox discovery is not automatic. Pass the mailbox address in the tool input when needed.
- No public HTTPS endpoint is required.
- The Python runtime requires Python 3.10 or newer. `uv` can download a compatible interpreter automatically when needed.
