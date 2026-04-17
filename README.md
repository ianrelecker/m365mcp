# M365 MCP

Local Desktop MCP server for one Microsoft 365 user. It lets Claude call Microsoft Graph with that user's delegated permissions, including any shared or delegated mailboxes and calendars the user already has access to.

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
http://localhost:8787/auth/microsoft/start
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

After Claude Desktop launches the MCP server, the local helper page is available at:

```text
http://localhost:8787/
```

That page lets you connect or disconnect Microsoft and check status.

## Windows Notes

- This project works on Windows. For your case, Windows is a fine target if Claude Desktop is also running on Windows.
- Install Node.js 20 or newer and make sure `node` and `pnpm` are on your `PATH`.
- If Claude Desktop cannot find `node`, replace `"command": "node"` with the full path to `node.exe`, for example `C:\\Program Files\\nodejs\\node.exe`.
- The Microsoft redirect URI stays the same on Windows: `http://localhost:8787/auth/microsoft/callback`.
- Keep the helper app bound to localhost. You do not need IIS, Linux, WSL, or a public web server for Claude Desktop local MCP.

## Security Notes

- `.env`, `.env.local`, and `.tokens/` are local-only files and are ignored by git.
- `TOKEN_ENCRYPTION_KEY` encrypts the saved Microsoft token cache at rest. If you rotate or lose it, delete `.tokens/microsoft-graph-token.json` and reconnect Microsoft.
- The recommended local setup is now public-client OAuth with PKCE, which removes the need to store `MICROSOFT_CLIENT_SECRET`.
- If you keep using a confidential-client `Web` app registration, treat `MICROSOFT_CLIENT_SECRET` like any other local credential and do not place it in shared configs or screenshots.

## Notes

- This is intentionally scoped for one user, not multi-tenant SaaS.
- This is for Claude Desktop local MCP, not `claude.ai` remote connectors.
- Shared/delegated mailbox discovery is not automatic. Pass the mailbox address in the tool input when needed.
- No public HTTPS endpoint is required.
