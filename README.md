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

Recommended settings for a new local Claude Desktop install:

- Platform: `Mobile and desktop applications`
- Custom redirect URI: `http://localhost:8787/auth/microsoft/callback`
- Supported account type: `Accounts in this organizational directory only`
- Advanced setting: `Allow public client flows = Yes`

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

If you already have a `Web` platform registration with a client secret, this server still supports it. For a local-only Desktop install, though, the public-client PKCE setup above is the safer default because it avoids storing `MICROSOFT_CLIENT_SECRET` on the machine.

## 2. Local config

Copy `.env.example` to `.env` and fill in the values.

Important values:

- `LOCAL_BASE_URL`
  Optional. Defaults to `http://localhost:8787`.
- `MICROSOFT_CLIENT_SECRET`
  Optional. Leave this blank for the recommended public-client PKCE setup. Only set it if you intentionally keep a confidential-client `Web` app registration.
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

Then configure Claude Desktop to launch the local stdio server. Example config:

```json
{
  "mcpServers": {
    "m365": {
      "command": "node",
      "args": ["C:\\path\\to\\claude-m365-mcp\\dist\\index.js"],
      "env": {
        "PORT": "8787",
        "LOCAL_BASE_URL": "http://localhost:8787",
        "MICROSOFT_TENANT_ID": "your-tenant-id",
        "MICROSOFT_CLIENT_ID": "your-client-id",
        "TOKEN_ENCRYPTION_KEY": "your-base64-32-byte-key",
        "KNOWN_MAILBOXES": "shared@company.com"
      }
    }
  }
}
```

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

## Security Notes

- `.env`, `.env.local`, and `.tokens/` are local-only files and are ignored by git.
- `TOKEN_ENCRYPTION_KEY` encrypts the saved Microsoft token cache at rest. If you rotate or lose it, delete `.tokens/microsoft-graph-token.json` and reconnect Microsoft.
- The recommended local setup is now public-client OAuth with PKCE, which removes the need to store `MICROSOFT_CLIENT_SECRET`.
- If you keep using a confidential-client `Web` app registration, treat `MICROSOFT_CLIENT_SECRET` like any other local credential and do not place it in shared configs or screenshots.

## Notes

- This is intentionally scoped for one user, not multi-tenant SaaS.
- This is for local MCP clients, not `claude.ai` remote connectors.
- Shared/delegated mailbox discovery is not automatic. Pass the mailbox address in the tool input when needed.
- No public HTTPS endpoint is required.
- The Python runtime requires Python 3.10 or newer. `uv` can download a compatible interpreter automatically when needed.
