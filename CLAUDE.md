# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A **local** MCP server (stdio) that gives an MCP client (Claude Desktop / Codex) delegated Microsoft Graph access to a single user's M365 mailbox, calendar, contacts, folders, rules, categories, and small attachments. It is not a remote connector — the client launches it on the user's machine via `uv`. End-user setup lives in `README.md`; this file is for working on the code.

## Commands

```bash
uv sync                                  # install deps (incl. dev extras)
uv run pytest                            # run all tests
uv run pytest tests/test_graph.py        # one test file
uv run pytest tests/test_graph.py::test_name   # single test
uv run mcp run src/m365_mcp/server.py    # manual smoke test (stop before opening Claude Desktop — both bind port 8787)
```

Python ≥3.10. No linter/formatter is configured. The packaged entry point is `m365-mcp` (`m365_mcp.server:main`).

## Architecture

Two servers run concurrently in one process, started by the FastMCP **lifespan** in `server.py`:
1. The **MCP stdio server** (`FastMCP`) that the client talks to.
2. An embedded **Starlette helper HTTP app** (`helper_app.py`) on `localhost:8787`, used only for the Microsoft OAuth browser flow (`/`, `/health`, `/auth/microsoft/{start,callback,disconnect}`). If the port is already bound, the helper is skipped with a stderr warning but the MCP server still runs.

### Layering (request flows top to bottom)

- **`server.py`** — defines every MCP tool. Tools are **thin wrappers**: they fetch the runtime and delegate straight to a `MicrosoftGraphClient` method. Put real logic in the graph client, not here. `RuntimeServices` is the DI container (config, auth, graph, http client, audit logger); `_RuntimeProvider` lazily builds a single instance and resets it on lifespan shutdown.
- **`microsoft_graph.py`** (`MicrosoftGraphClient`, ~2500 lines) — all Microsoft Graph REST calls. Central helpers: `_request` (JSON, always hits `https://graph.microsoft.com/v1.0`, sends `Prefer: IdType="ImmutableId"`), `_request_bytes` (attachment downloads), and `_base_path(mailbox)` which returns `/me` or `/users/{mailbox}`. **Mailbox routing is the core pattern**: nearly every tool takes an optional `mailbox` arg; absent → own mailbox (`/me`), present → that shared/delegated mailbox.
- **`microsoft_auth.py`** (`MicrosoftAuthService`) — OAuth authorization-code + PKCE flow, token exchange/refresh, scope/status reporting. `get_access_token()` auto-refreshes when within 60s of expiry. Requires a confidential-client secret (`Web` app), and `offline_access` for the refresh token.
- **`token_store.py` + `crypto.py`** — `EncryptedFileStore` persists tokens to `.tokens/microsoft-graph-token.json`, encrypted at rest with AES-256-GCM using `TOKEN_ENCRYPTION_KEY` (base64 32-byte key).
- **`config.py`** — `load_config()` reads env (via `python-dotenv`). The Graph **scope list is hardcoded** here (`MicrosoftConfig.scopes`); the test fixture in `tests/conftest.py` duplicates it, so update both together.
- **`models.py`** — pydantic models for all tool results and stored token/payload shapes.

### Auditing

`_AuditedFastMCP` (in `server.py`) overrides `call_tool` to record every tool invocation (success or error) via `LocalAuditLogger` (`audit.py`) to `.audit/m365-mcp-audit.jsonl`. `audit.py` classifies tools by category, captures key IDs (`ID_FIELDS`), and **redacts** `SENSITIVE_FIELDS` (bodies, subjects, addresses, etc.) from logged error messages. Never log token/secret/body content here.

## Conventions & gotchas

- **Tool/param names use camelCase** (e.g. `messageId`, `folderPath`, `inferenceClassification`) to mirror Graph, even though this is Python. The `from` mail field is aliased to `from_` via a pydantic `Field` because `from` is a Python keyword.
- **Prefer drafts over sending.** `mail_send`/`mail_send_reply` send immediately; tool descriptions instruct preferring `mail_create_draft`/`mail_create_reply_draft` unless the user explicitly approved sending. Preserve that guidance.
- **Folders** can be addressed three ways: well-known name (`Inbox`), `folderPath` (`Inbox/Clients/Acme`), or raw `folderId`. Resolution helpers live in the graph client (`_resolve_mail_folder_*`).
- **Categories** use Outlook *master categories*. Graph cannot rename an existing master category — `mail_update_category` only changes color; renaming means create-new + delete-old.
- **`M365_MCP_CAPABILITIES.md`** is the model-facing usage guide, served both as the `m365_capabilities` tool and the `m365://capabilities` MCP resource (loaded from disk at call time). Keep it in sync when tool behavior changes.
- `create_contact`/`update_contact` issue a follow-up GET (`_build_contact_read_query` with `$expand`) after the POST/PATCH, because Graph does not return the `personalHomePage` extended property on a write response. Tests that mock these must answer that re-fetch.

### SharePoint files and Excel workbooks

Two extra Graph clients live alongside `MicrosoftGraphClient`, each self-contained (own pydantic models, own `_request`, same shared auth + httpx client). Both are constructed in `create_runtime`, held on `RuntimeServices`, and exposed through tools in `server.py`:

- `sharepoint_files.py` (`SharePointFilesClient`) — **read-only** browsing of SharePoint sites, document libraries (drives), and folders; `sharepoint_*` tools.
- `excel_workbook.py` (`ExcelWorkbookClient`) — **in-place** `.xlsx`/`.xlsm` editing via the Graph Workbook API; `workbook_*` tools. The client methods take a `WorkbookItemRef`; the tools take flat `driveId`/`itemId` args and build the ref internally (the `workbook_resolve` tool returns those IDs). Beyond single-range read/write, append-table-row, and sessions, it has: batch read/write (`get_ranges`/`update_ranges`, both via Graph `POST /$batch`, auto-chunked to ≤20 sub-requests, per-request errors aggregated in input order without failing the whole batch — and `update_ranges`/`update_range` write `formulas`, including cross-sheet strings like `='Unit Mix'!H11`); `calculate` (force recalc); defined names (`list_names`/`get_name_range`); and `clear_range`/`copy_range`/`insert_range`/`delete_range` (the last shifts cells `Up`/`Left` to close the gap — pass a full-row address for a delete-row — versus `clear_range`, which only blanks cells in place). Note `delete_range` deletes *cells/rows inside a worksheet* via the Workbook API; this is in-place editing, **not** the forbidden drive-item file deletion (see the no-deletion policy below). Worksheets are read-only here — `list_worksheets` enumerates tabs, but there are no add/copy/rename/delete-worksheet tools (deliberately removed). All accept an optional `sessionId` so a caller can run read→write→calculate→read inside one `persistChanges:true` session.

These need the `Sites.Read.All` and `Files.ReadWrite.All` scopes (already in the hardcoded list in `config.py` + `tests/conftest.py`). The browse client is read-only, so only `Sites.Read.All` is requested — there is intentionally **no** `Sites.ReadWrite.All`; in-place workbook edits are authorized by `Files.ReadWrite.All` because they go through `/drives/{id}/items/{id}/workbook`, not `/sites`. Unlike the mail/contacts/calendar tools, the file/workbook tools take no `mailbox` arg — they operate by `driveId`/`itemId`, so the audit logger records no mailbox for them.

**No SharePoint/OneDrive deletion — intentional and must stay that way.** The `Files.ReadWrite.All` scope *permits* `DELETE` at the Graph level, but no tool exposes it for *drive items*: `sharepoint_*` is read-only and `workbook_*` only edits the contents of a workbook in place (cell values, table rows, and cell/row deletion via the Workbook API). The distinction is the unit being deleted: deleting cells/rows **inside** a workbook (`.../workbook/.../range(...)/delete`) is allowed in-place editing; deleting the **file or folder itself** is not. There must be **no** tool that deletes a SharePoint/OneDrive file or folder (`DELETE /drives/{id}/items/{id}`), removes a drive item, or recursively clears a folder. Deleting files is a manual, human action: if a user asks for one, do not attempt it — explain that deletion requires explicit confirmation and is performed manually outside this server. Do not add a delete tool or a `DELETE` request path *against drive items* here without an explicit, deliberate decision to change this policy.

## Tests

`tests/` uses pytest + anyio (asyncio backend). `conftest.py` provides a `config_factory` fixture and a `make_jwt` helper. Graph/auth tests stub HTTP by injecting an `httpx.AsyncClient` (or transport) into the services rather than hitting the network — follow that pattern for new tests.
