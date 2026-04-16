# Claude M365 MCP

Local Claude Desktop MCP server for one Microsoft 365 user. It lets Claude call Microsoft Graph with that user's delegated permissions, including any shared or delegated mailboxes and calendars the user already has access to.

This project only does one auth flow:

- This local server -> Microsoft 365: delegated OAuth against Microsoft Entra / Graph

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
- Redirect URI: `http://127.0.0.1:8787/auth/microsoft/callback`
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
  Optional. Defaults to `http://127.0.0.1:8787`.
- `TOKEN_ENCRYPTION_KEY`
  Base64-encoded 32 byte key used for encrypted Microsoft token storage.

Generate `TOKEN_ENCRYPTION_KEY` with:

```bash
node -e "console.log(require('crypto').randomBytes(32).toString('base64'))"
```

## 3. Run it

Install deps:

```bash
pnpm install
```

Development:

```bash
pnpm dev
```

Production build:

```bash
pnpm build
node dist/index.js
```

## 4. Connect Microsoft

Open:

```text
http://127.0.0.1:8787/auth/microsoft/start
```

Sign in with the one Microsoft 365 user you want this server to act as.

After this, the server keeps that user's Microsoft refresh token in `.tokens/microsoft-graph-token.json`, encrypted with `TOKEN_ENCRYPTION_KEY`.

## 5. Add it to Claude Desktop as a local MCP server

Build first:

```bash
pnpm build
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
        "LOCAL_BASE_URL": "http://127.0.0.1:8787",
        "MICROSOFT_TENANT_ID": "your-tenant-id",
        "MICROSOFT_CLIENT_ID": "your-client-id",
        "MICROSOFT_CLIENT_SECRET": "your-client-secret",
        "TOKEN_ENCRYPTION_KEY": "your-base64-32-byte-key",
        "KNOWN_MAILBOXES": "shared@company.com"
      }
    }
  }
}
```

After Claude Desktop launches the MCP server, the local helper page is available at:

```text
http://127.0.0.1:8787/
```

That page lets you connect or disconnect Microsoft and check status.

## Windows Notes

- This project works on Windows. For your case, Windows is a fine target if Claude Desktop is also running on Windows.
- Install Node.js 20 or newer and make sure `node` and `pnpm` are on your `PATH`.
- If Claude Desktop cannot find `node`, replace `"command": "node"` with the full path to `node.exe`, for example `C:\\Program Files\\nodejs\\node.exe`.
- The Microsoft redirect URI stays the same on Windows: `http://127.0.0.1:8787/auth/microsoft/callback`.
- Keep the helper app bound to localhost. You do not need IIS, Linux, WSL, or a public web server for Claude Desktop local MCP.

## Notes

- This is intentionally scoped for one user, not multi-tenant SaaS.
- This is for Claude Desktop local MCP, not `claude.ai` remote connectors.
- Shared/delegated mailbox discovery is not automatic. Pass the mailbox address in the tool input when needed.
- No public HTTPS endpoint is required.
