# Codex MCP Setup

Use this when adding this M365 MCP server with **Connect to a custom MCP**.

The most important rule: **Command to launch is only the executable.** Do not paste the whole command into that field.

## One-Time Setup

Create the local env file first:

```bash
npx -y @ianrelecker/m365mcp init
```

Fill in the Microsoft values in the env file it creates:

```text
~/.m365mcp/.env
```

On Windows, that is usually:

```text
C:\Users\YOUR_WINDOWS_USER\.m365mcp\.env
```

You can check the setup with:

```bash
npx -y @ianrelecker/m365mcp doctor
```

## MCP Fields

Use **STDIO**.

Name:

```text
m365
```

Command to launch:

```text
npx
```

If `npx` is not found, install Node.js 20 or newer and restart the app.

Arguments:

Add each value below as its own separate argument row.

```text
-y
```

```text
@ianrelecker/m365mcp
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
Leave as the default
```

The server loads `~/.m365mcp/.env` automatically, so you do not need to point the MCP config at a cloned repo or manually run `uv`.

## Custom Env File

If you do not want to use `~/.m365mcp/.env`, add two more argument rows:

```text
--env-file
```

```text
C:\full\path\to\.env
```

On macOS or Linux, use a Unix path instead:

```text
/full/path/to/.env
```

## After Saving

After saving the MCP config, ask the model to use the M365 MCP or check auth status.

Then open the local auth site:

```text
http://localhost:8787/auth/microsoft/start
```

If the browser cannot open that page, the server probably did not start. Recheck the command and arguments, then run:

```bash
npx -y @ianrelecker/m365mcp doctor
```

## Common Mistake

Do not set `Command to launch` to this:

```text
npx -y @ianrelecker/m365mcp
```

That makes the app look for one executable with that entire name, so the MCP server never starts and no tools appear.

Use this instead:

```text
Command to launch: npx
Arguments:
-y
@ianrelecker/m365mcp
```

## Troubleshooting

- If no M365 tools were found, the server likely did not start. Check that the arguments are separate rows.
- If `npx` is not found, install Node.js 20 or newer and restart the app.
- If `uv` is not found, install `uv` or set `UV_PATH` to the full `uv` path.
- If the logs mention missing Microsoft env values, run `npx -y @ianrelecker/m365mcp init`, fill in `~/.m365mcp/.env`, and restart.
- If the logs mention `TOKEN_ENCRYPTION_KEY`, run `init` again or generate a real base64 32-byte key.
- If the logs mention port `8787`, another copy may already be running. Stop the other process and retry.
- If `auth_status` shows missing scopes, add the missing Microsoft Graph delegated permissions in Azure and reconnect Microsoft.
