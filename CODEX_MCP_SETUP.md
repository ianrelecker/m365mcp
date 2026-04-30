# Codex MCP Setup

Use this when adding this M365 MCP server in Codex with **Connect to a custom MCP**.

The most important rule: **Command to launch is only the executable.** Do not paste the whole command into that field.

## Windows

Use **STDIO**.

Name:

```text
m365
```

Command to launch:

```text
uv
```

If `uv` fails, use the full path from PowerShell:

```powershell
where uv
```

For example:

```text
C:\Users\Administrator\.local\bin\uv.exe
```

Arguments:

Add each value below as its own separate argument row.

```text
run
```

```text
--env-file
```

```text
.env
```

```text
mcp
```

```text
run
```

```text
src/m365_mcp/server.py
```

Environment variables:

```text
Leave blank
```

Environment variable passthrough:

```text
Leave blank
```

Working directory:

```text
C:\Users\Administrator\Desktop\m365mcp-main
```

Change that path if your repo is somewhere else. It must be the folder that contains `pyproject.toml` and `.env`.

The command Codex should effectively run is:

```powershell
uv run --env-file .env mcp run src/m365_mcp/server.py
```

But in the Codex UI, do **not** paste that full line into `Command to launch`.

## macOS Or Linux

Use **STDIO**.

Name:

```text
m365
```

Command to launch:

```text
uv
```

If `uv` fails, use the full path from:

```bash
which uv
```

Arguments:

Add each value below as its own separate argument row.

```text
run
```

```text
--env-file
```

```text
.env
```

```text
mcp
```

```text
run
```

```text
src/m365_mcp/server.py
```

Environment variables:

```text
Leave blank
```

Working directory:

```text
/path/to/m365mcp
```

Use the real path to the repo folder that contains `pyproject.toml` and `.env`.

## After Saving

After saving the MCP config, ask Codex to use the M365 MCP or check auth status.

Then open the local auth site:

```text
http://localhost:8787/auth/microsoft/start
```

If the browser cannot open that page, Codex probably did not start the server. Recheck the command, arguments, working directory, and `.env`.

## Common Mistake

Do not set `Command to launch` to this:

```text
uv run --env-file .env mcp run src/m365_mcp/server.py
```

That makes Codex look for one executable with that entire name, so the MCP server never starts and Codex will not see any tools.

Use this instead:

```text
Command to launch: uv
Arguments:
run
--env-file
.env
mcp
run
src/m365_mcp/server.py
```

## Troubleshooting

- If Codex says no M365 tools were found, the server likely did not start. Check the arguments are separate rows.
- If `uv` is not found, use the full `uv.exe` path from `where uv`.
- If the logs mention `.env`, confirm `.env` is in the working directory.
- If the logs mention `TOKEN_ENCRYPTION_KEY`, generate a real base64 32-byte key and update `.env`.
- If the logs mention port `8787`, another copy may already be running. Stop the other process and retry.
- If `auth_status` shows missing scopes, add the missing Microsoft Graph delegated permissions in Azure and reconnect Microsoft.
