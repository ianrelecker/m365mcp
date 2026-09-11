# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A **local** MCP server (stdio) that gives an MCP client (Claude Desktop / Codex) delegated Microsoft Graph access to a single user's M365 mailbox, calendar, contacts, folders, rules, categories, and small attachments. It is not a remote connector — the client launches it on the user's machine via `uv`. End-user setup lives in `README.md`; this file is for working on the code.

## Commands

```bash
uv sync --extra dev                      # install deps incl. dev extras (plain `uv sync` skips pytest)
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
- **`microsoft_graph.py`** (`MicrosoftGraphClient`, ~2500 lines) — all Microsoft Graph REST calls. Central helpers: `_request` (JSON, always hits `https://graph.microsoft.com/v1.0`, sends `Prefer: IdType="ImmutableId"`), `_request_bytes` (attachment downloads), `_attachment_bytes` (prefers the inline `contentBytes` Graph returns, falling back to `/$value`), and `_base_path(mailbox)` which returns `/me` or `/users/{mailbox}`. **Mailbox routing is the core pattern**: nearly every tool takes an optional `mailbox` arg; absent → own mailbox (`/me`), present → that shared/delegated mailbox.
- **`microsoft_auth.py`** (`MicrosoftAuthService`) — OAuth authorization-code + PKCE flow, token exchange/refresh, scope/status reporting. `get_access_token()` auto-refreshes when within 60s of expiry. Requires a confidential-client secret (`Web` app), and `offline_access` for the refresh token.
- **`token_store.py` + `crypto.py`** — `EncryptedFileStore` persists tokens to `.tokens/microsoft-graph-token.json`, encrypted at rest with AES-256-GCM using `TOKEN_ENCRYPTION_KEY` (base64 32-byte key).
- **`config.py`** — `load_config()` reads env (via `python-dotenv`). Graph scopes come from `graph_scopes()`; `Mail.Send` / `Mail.Send.Shared` are included only when `M365_MAIL_SEND_ENABLED` is true. The test fixture in `tests/conftest.py` must stay in sync.
- **`pid_policy.py`** — optional fail-closed allowlist/blocklist/redaction layer injected into the Graph, SharePoint, and Excel clients. Blocked calls raise `BlockedError` with a reason code and must never put PID in the exception message.
- **`models.py`** — pydantic models for all tool results and stored token/payload shapes.

### Auditing

`_AuditedFastMCP` (in `server.py`) overrides `call_tool` to record every tool invocation (success or error) via `LocalAuditLogger` (`audit.py`) to `.audit/m365-mcp-audit.jsonl`. `audit.py` classifies tools by category, captures key IDs (`ID_FIELDS`), and **redacts** `SENSITIVE_FIELDS` (bodies, subjects, addresses, etc.) from logged error messages. Never log token/secret/body content here.

## Conventions & gotchas

- **Tool/param names use camelCase** (e.g. `messageId`, `folderPath`, `inferenceClassification`) to mirror Graph, even though this is Python. The `from` mail field is aliased to `from_` via a pydantic `Field` because `from` is a Python keyword.
- **Prefer drafts over sending.** Send tools are unregistered unless `M365_MAIL_SEND_ENABLED` is true. When they exist, `mail_send`/`mail_send_reply` send immediately; tool descriptions instruct preferring `mail_create_draft`/`mail_create_reply_draft` unless the user explicitly approved sending. Preserve that guidance.
- **Folders** can be addressed three ways: well-known name (`Inbox`), `folderPath` (`Inbox/Clients/Acme`), or raw `folderId`. Resolution helpers live in the graph client (`_resolve_mail_folder_*`).
- **Categories** use Outlook *master categories*. Graph cannot rename an existing master category — `mail_update_category` only changes color; renaming means create-new + delete-old.
- **`manifest.json`** is the [MCP Bundle](https://github.com/anthropics/mcpb) manifest used to install this server into Claude Desktop as a `.mcpb`. It uses `server.type: "uv"` (bundle ships `src/` + `pyproject.toml` + `uv.lock`; the host installs deps with `uv`) and maps `user_config` entries to the env vars `config.py` reads, so a bundle install needs no `.env`. Its `tools` array lists every tool — after adding/renaming/removing one, run `uv run python scripts/sync_mcpb_tools.py`; `tests/test_manifest.py` fails if it is stale. Adding a new required env var to `config.py` means adding a matching `user_config` entry and `mcp_config.env` mapping. `.mcpbignore` keeps `.env`, `.tokens/`, `.audit/`, and `claude_desktop_config.json` out of the packed bundle. The packed `.mcpb` is **not** committed (it is a zip of this repo); `.github/workflows/release-mcpb.yml` builds it on a `v*` tag and attaches it to the GitHub Release, and it fails the build if the tag and `manifest.json` version disagree — so bump `version` in `manifest.json` and `pyproject.toml` together.
- **Picture ingestion is the one place tools are not plain pydantic returns.** `mail_get_attachment_image`, `mail_get_inline_images`, and `mail_get_attachment_pdf_pages` are registered with `structured_output=False` and return `list[ContentBlock]`: a `TextContent` metadata summary plus one `ImageContent` per picture, built by `_image_content_blocks(result, field)` in `server.py` (`field` names the model attribute holding the payloads — `image`, `images`, or `pages` — and a singular field stays singular in the summary). The base64 payload must ride in the image blocks only — the text summary drops `dataBase64` so the encoded bytes are never duplicated as text. The graph client still returns ordinary models (`MailAttachmentImageResult`, `MailInlineImagesResult`, `MailAttachmentPdfResult`); only the rendering lives in `server.py`. Supported image formats are PNG/JPEG/GIF/WEBP (what MCP clients render); anything else gets an `unsupportedReason`. Every payload model must expose `mimeType` + `dataBase64`, which is what `_image_content_blocks` relies on.
- **PDF attachments have two paths, deliberately.** `mail_get_attachment_content` extracts the text layer with `pypdf` (cheap, exact, useless on a scan); `mail_get_attachment_pdf_pages` rasterizes pages with `pypdfium2` + `pillow` (`_render_pdf_pages`) to JPEG at a `longEdge`-bounded resolution, which is what makes scans, tables, and layout readable. Both renderer imports are optional at module load like `PdfReader`, so a missing dependency degrades to an `unsupportedReason` instead of an import error — keep that pattern, and keep `_PdfRenderError` as the way render failures (password-protected, damaged) become reasons rather than exceptions. Page images are bounded by `maxPages`, `longEdge`, and the shared `maxTotalBytes` budget; the first page is always rendered even if it alone exceeds the budget, so a caller never gets an empty result for a valid page.
- **`M365_MCP_CAPABILITIES.md`** is the model-facing usage guide, served both as the `m365_capabilities` tool and the `m365://capabilities` MCP resource (loaded from disk at call time). Keep it in sync when tool behavior changes.
- `create_contact`/`update_contact` issue a follow-up GET (`_build_contact_read_query` with `$expand`) after the POST/PATCH, because Graph does not return the `personalHomePage` extended property on a write response. Tests that mock these must answer that re-fetch.

### SharePoint files and Excel workbooks

Two extra Graph clients live alongside `MicrosoftGraphClient`, each self-contained (own pydantic models, own `_request`, same shared auth + httpx client). Both are constructed in `create_runtime`, held on `RuntimeServices`, and exposed through tools in `server.py`:

- `sharepoint_files.py` (`SharePointFilesClient`) — browsing of SharePoint sites, document libraries (drives), and folders plus sharing-link/permission management; `sharepoint_*` tools. Sharing mutations require an explicit `confirm=true`; invitation email defaults off, inherited permissions are always retained, and sharing URLs, recipients, and invitation messages are sensitive audit fields.
- `excel_workbook.py` (`ExcelWorkbookClient`) — **in-place** `.xlsx`/`.xlsm` editing via the Graph Workbook API; `workbook_*` tools. The client methods take a `WorkbookItemRef`; the tools take flat `driveId`/`itemId` args and build the ref internally (the `workbook_resolve` tool returns those IDs). Beyond single-range read/write, append-table-row, and sessions, it has: batch read/write (`get_ranges`/`update_ranges`, both via Graph `POST /$batch`, auto-chunked to ≤20 sub-requests, per-request errors aggregated in input order without failing the whole batch — and `update_ranges`/`update_range` write `formulas`, including cross-sheet strings like `='Unit Mix'!H11`); `calculate` (force recalc); defined names (`list_names`/`get_name_range`); `clear_range`/`copy_range`/`insert_range`/`delete_range` (the last shifts cells `Up`/`Left` to close the gap — pass a full-row address for a delete-row — versus `clear_range`, which only blanks cells in place); table structure (`create_table` = convert a range to a table, `sort_table`, `filter_table`, `clear_table_filters`); and formatting (`format_range` — fill/font/alignment/borders, issuing one PATCH per sub-property; `set_column_width`/`set_row_height`, each also doing autofit via `autofit=True`). Number formats stay on `update_range`/`update_ranges`, not `format_range`. Note `delete_range` deletes *cells/rows inside a worksheet* via the Workbook API; this is in-place editing, **not** the forbidden drive-item file deletion (see the no-deletion policy below). Worksheets are read-only here — `list_worksheets` enumerates tabs, but there are no add/copy/rename/delete-worksheet tools (deliberately removed). All accept an optional `sessionId` so a caller can run read→write→calculate→read inside one `persistChanges:true` session.

These need the `Sites.Read.All` and `Files.ReadWrite.All` scopes (already in the hardcoded list in `config.py` + `tests/conftest.py`). There is intentionally **no** `Sites.ReadWrite.All`; sharing and in-place workbook edits are authorized by `Files.ReadWrite.All` because they go through `/drives/{id}/items/{id}`, while site browsing uses `Sites.Read.All`. Unlike the mail/contacts/calendar tools, the file/workbook tools take no `mailbox` arg — they operate by `driveId`/`itemId`, so the audit logger records no mailbox for them.

**No SharePoint/OneDrive file deletion — intentional and must stay that way.** The `Files.ReadWrite.All` scope *permits* `DELETE` at the Graph level, but no tool deletes a drive item. `sharepoint_revoke_permission` deletes only `/permissions/{permissionId}`, and `workbook_*` only edits workbook contents. There must be **no** tool that deletes a file or folder (`DELETE /drives/{id}/items/{id}`), removes a drive item, or recursively clears a folder. Deleting files remains a manual human action outside this server.

## Tests

`tests/` uses pytest + anyio (asyncio backend). `conftest.py` provides a `config_factory` fixture and a `make_jwt` helper. Graph/auth tests stub HTTP by injecting an `httpx.AsyncClient` (or transport) into the services rather than hitting the network — follow that pattern for new tests.
