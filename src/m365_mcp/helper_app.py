from __future__ import annotations

from html import escape

from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import HTMLResponse, JSONResponse, RedirectResponse
from starlette.routing import Route

from .config import AppConfig
from .microsoft_auth import MicrosoftAuthService


def render_home_page(config: AppConfig, status_text: dict[str, object]) -> str:
    connected = (
        "<p><strong>Microsoft status:</strong> Connected"
        + (
            f" as <code>{escape(str(status_text['preferredUsername']))}</code>"
            if status_text.get("preferredUsername")
            else ""
        )
        + ".</p>"
        if status_text["connected"]
        else "<p><strong>Microsoft status:</strong> Not connected yet.</p>"
    )

    mailboxes = (
        "<ul>"
        + "".join(
            f"<li><code>{escape(str(mailbox))}</code></li>"
            for mailbox in status_text["knownMailboxes"]
        )
        + "</ul>"
        if status_text["knownMailboxes"]
        else "<p>No mailbox hints configured.</p>"
    )

    return f"""<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Claude M365 MCP</title>
    <style>
      body {{ font-family: ui-sans-serif, system-ui, sans-serif; margin: 2rem; color: #111827; }}
      main {{ max-width: 56rem; margin: 0 auto; display: grid; gap: 1rem; }}
      section {{ border: 1px solid #d1d5db; border-radius: 0.85rem; padding: 1rem 1.2rem; }}
      code {{ background: #f3f4f6; padding: 0.15rem 0.35rem; border-radius: 0.3rem; }}
      a.button {{ display: inline-block; padding: 0.75rem 0.95rem; background: #111827; color: #fff; border-radius: 0.6rem; text-decoration: none; margin-right: 0.5rem; }}
      p, li {{ line-height: 1.5; }}
    </style>
  </head>
  <body>
    <main>
      <section>
        <h1>Claude M365 MCP</h1>
        <p>This server exposes Microsoft 365 mail and calendar tools to Claude Desktop over a local MCP stdio connection.</p>
        <p><strong>Local helper URL:</strong> <code>{escape(config.localBaseUrl)}</code></p>
        <p><strong>Microsoft callback URI:</strong> <code>{escape(config.microsoft.redirectUri)}</code></p>
        <p>No public HTTPS endpoint is required for Claude Desktop local MCP.</p>
      </section>

      <section>
        <h2>Microsoft Delegated Auth</h2>
        {connected}
        <p>
          <a class="button" href="/auth/microsoft/start">Connect Microsoft 365</a>
          <a class="button" href="/auth/microsoft/disconnect">Disconnect Microsoft 365</a>
        </p>
      </section>

      <section>
        <h2>Known Mailboxes</h2>
        {mailboxes}
        <p>The MCP tools also accept an explicit <code>mailbox</code> argument for any shared or delegated mailbox/calendar the connected Microsoft user can access.</p>
      </section>

      <section>
        <h2>Desktop Setup</h2>
        <p>Point Claude Desktop at this project as a local MCP stdio server, then open the Microsoft connect URL once in your browser.</p>
      </section>
    </main>
  </body>
</html>"""


def create_helper_app(config: AppConfig, microsoft_auth: MicrosoftAuthService) -> Starlette:
    async def home(_request: Request) -> HTMLResponse:
        status = await microsoft_auth.get_status()
        return HTMLResponse(
            render_home_page(
                config,
                {
                    "connected": status.connected,
                    "preferredUsername": status.account.preferredUsername if status.account else None,
                    "knownMailboxes": status.knownMailboxes,
                },
            )
        )

    async def health(_request: Request) -> JSONResponse:
        status = await microsoft_auth.get_status()
        return JSONResponse(
            {
                "ok": True,
                "microsoftConnected": status.connected,
                "localBaseUrl": config.localBaseUrl,
                "microsoftRedirectUri": config.microsoft.redirectUri,
            }
        )

    async def auth_start(_request: Request) -> RedirectResponse:
        return RedirectResponse(microsoft_auth.build_authorization_url(), status_code=302)

    async def auth_callback(request: Request) -> HTMLResponse:
        try:
            await microsoft_auth.handle_authorization_code_callback(
                code=request.query_params.get("code"),
                state=request.query_params.get("state"),
                error=request.query_params.get("error"),
                errorDescription=request.query_params.get("error_description"),
            )
            return HTMLResponse(
                """<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Microsoft Connected</title>
  </head>
  <body style="font-family: ui-sans-serif, system-ui, sans-serif; margin: 2rem;">
    <h1>Microsoft 365 connected</h1>
    <p>You can close this tab and go back to Claude.</p>
    <p><a href="/">Return to status page</a></p>
  </body>
</html>"""
            )
        except Exception as error:
            return HTMLResponse(
                f"""<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Microsoft Auth Error</title>
  </head>
  <body style="font-family: ui-sans-serif, system-ui, sans-serif; margin: 2rem;">
    <h1>Microsoft 365 connection failed</h1>
    <pre style="white-space: pre-wrap;">{escape(str(error))}</pre>
    <p><a href="/">Return to status page</a></p>
  </body>
</html>""",
                status_code=400,
            )

    async def disconnect(_request: Request) -> RedirectResponse:
        await microsoft_auth.disconnect()
        return RedirectResponse("/", status_code=302)

    return Starlette(
        routes=[
            Route("/", home),
            Route("/health", health),
            Route("/auth/microsoft/start", auth_start),
            Route("/auth/microsoft/callback", auth_callback),
            Route("/auth/microsoft/disconnect", disconnect),
        ]
    )
