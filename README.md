# M365 MCP

Local MCP server for one Microsoft 365 user. It lets Claude call Microsoft Graph with that user's delegated permissions, including shared or delegated mailboxes and calendars the user already has access to.

This Python port runs with `uv` and the official MCP Python SDK. It starts two things in one process:

- a stdio MCP server for Claude/Desktop-style clients
- a localhost-only helper web app for Microsoft OAuth and status checks

## What it exposes

- `auth_status`
- `mail_list`
- `mail_search`
- `mail_get`
- `mail_list_drafts`
- `mail_create_draft`
- `mail_send_draft`
- `mail_move`
- `calendar_list_events`
- `calendar_create_event`

Every mail/calendar tool accepts an optional `mailbox` argument. Leave it blank for the signed-in user's own mailbox. Set it to something like `shared@company.com` for a shared or delegated mailbox/calendar.

## 1. Microsoft Entra setup

Create a single-tenant app registration for local use.

Recommended settings:

- Platform: `Web`
- Redirect URI: `http://localhost:8787/auth/microsoft/callback`
- Supported account type: `Accounts in this organizational directory only`

Add delegated Microsoft Graph permissions:

- `Mail.ReadWrite.Shared`
- `Mail.Send.Shared`
- `Mail.Send` (optional, but reasonable if you also want an explicit non-shared send scope)
- `Calendars.ReadWrite.Shared`
- `openid`
- `profile`
- `email`
- `offline_access`

If your tenant restricts user consent, grant admin consent for the enterprise app.

## 2. Local config

Copy `.env.example` to `.env` and fill in the values.

Important values:

- `LOCAL_BASE_URL`
  Optional. Defaults to `http://localhost:8787`.
- `TOKEN_ENCRYPTION_KEY`
  Base64-encoded 32 byte key used for encrypted Microsoft token storage.

Generate `TOKEN_ENCRYPTION_KEY` with:

```bash
python3 -c "import os, base64; print(base64.b64encode(os.urandom(32)).decode())"
```

## 3. Install and run

Install dependencies with `uv`:

```bash
uv sync
```

Run the stdio MCP server:

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

If you prefer, you can point Claude Desktop at the same `uv run mcp run ...` entrypoint manually, but the install command above is the easiest path and keeps the server on the FastMCP-supported workflow.

## 6. Run tests

Install dev dependencies and run the test suite:

```bash
uv sync --extra dev
uv run pytest
```

## Windows Notes

- This project works on Windows. If Claude Desktop is running on Windows, this is a fine target.
- Install `uv`. It can manage the required Python version for you.
- The Microsoft redirect URI stays the same on Windows: `http://localhost:8787/auth/microsoft/callback`.
- Keep the helper app bound to localhost. You do not need IIS, Linux, WSL, or a public web server for local MCP use.

## Notes

- This is intentionally scoped for one user, not multi-tenant SaaS.
- This is for local MCP clients, not `claude.ai` remote connectors.
- Shared/delegated mailbox discovery is not automatic. Pass the mailbox address in the tool input when needed.
- No public HTTPS endpoint is required.
- The Python runtime requires Python 3.10 or newer. `uv` can download a compatible interpreter automatically when needed.
