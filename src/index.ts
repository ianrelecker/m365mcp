import express from "express";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

import { config } from "./config.js";
import { createMcpServer } from "./mcp-server.js";
import { MicrosoftAuthService } from "./microsoft-auth.js";
import { MicrosoftGraphClient } from "./microsoft-graph.js";

function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function renderHomePage(status: Awaited<ReturnType<MicrosoftAuthService["getStatus"]>>): string {
  const connected = status.connected
    ? `<p><strong>Microsoft status:</strong> Connected${
        status.account?.preferredUsername
          ? ` as <code>${escapeHtml(status.account.preferredUsername)}</code>`
          : ""
      }.</p>`
    : `<p><strong>Microsoft status:</strong> Not connected yet.</p>`;

  const mailboxes =
    status.knownMailboxes.length > 0
      ? `<ul>${status.knownMailboxes
          .map((mailbox) => `<li><code>${escapeHtml(mailbox)}</code></li>`)
          .join("")}</ul>`
      : "<p>No mailbox hints configured.</p>";

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Claude M365 MCP</title>
    <style>
      body { font-family: ui-sans-serif, system-ui, sans-serif; margin: 2rem; color: #111827; }
      main { max-width: 56rem; margin: 0 auto; display: grid; gap: 1rem; }
      section { border: 1px solid #d1d5db; border-radius: 0.85rem; padding: 1rem 1.2rem; }
      code { background: #f3f4f6; padding: 0.15rem 0.35rem; border-radius: 0.3rem; }
      a.button { display: inline-block; padding: 0.75rem 0.95rem; background: #111827; color: #fff; border-radius: 0.6rem; text-decoration: none; margin-right: 0.5rem; }
      p, li { line-height: 1.5; }
    </style>
  </head>
  <body>
    <main>
      <section>
        <h1>Claude M365 MCP</h1>
        <p>This server exposes Microsoft 365 mail and calendar tools to Claude Desktop over a local MCP stdio connection.</p>
        <p><strong>Local helper URL:</strong> <code>${escapeHtml(config.localBaseUrl.toString())}</code></p>
        <p><strong>Microsoft callback URI:</strong> <code>${escapeHtml(config.microsoft.redirectUri)}</code></p>
        <p>No public HTTPS endpoint is required for Claude Desktop local MCP.</p>
      </section>

      <section>
        <h2>Microsoft Delegated Auth</h2>
        ${connected}
        <p>
          <a class="button" href="/auth/microsoft/start">Connect Microsoft 365</a>
          <a class="button" href="/auth/microsoft/disconnect">Disconnect Microsoft 365</a>
        </p>
      </section>

      <section>
        <h2>Known Mailboxes</h2>
        ${mailboxes}
        <p>The MCP tools also accept an explicit <code>mailbox</code> argument for any shared or delegated mailbox/calendar the connected Microsoft user can access.</p>
      </section>

      <section>
        <h2>Desktop Setup</h2>
        <p>Point Claude Desktop at this project as a local MCP stdio server, then open the Microsoft connect URL once in your browser.</p>
      </section>
    </main>
  </body>
</html>`;
}

const app = express();

app.use(express.urlencoded({ extended: false }));

const microsoftAuth = new MicrosoftAuthService(config);
const graph = new MicrosoftGraphClient(microsoftAuth);

app.get("/", async (_req, res) => {
  const status = await microsoftAuth.getStatus();
  res.type("html").send(renderHomePage(status));
});

app.get("/health", async (_req, res) => {
  const status = await microsoftAuth.getStatus();
  res.json({
    ok: true,
    microsoftConnected: status.connected,
    localBaseUrl: config.localBaseUrl.toString(),
    microsoftRedirectUri: config.microsoft.redirectUri,
  });
});

app.get("/auth/microsoft/start", (_req, res) => {
  res.redirect(microsoftAuth.buildAuthorizationUrl());
});

app.get("/auth/microsoft/callback", async (req, res) => {
  try {
    await microsoftAuth.handleAuthorizationCodeCallback({
      code: typeof req.query.code === "string" ? req.query.code : undefined,
      state: typeof req.query.state === "string" ? req.query.state : undefined,
      error: typeof req.query.error === "string" ? req.query.error : undefined,
      errorDescription:
        typeof req.query.error_description === "string"
          ? req.query.error_description
          : undefined,
    });

    res
      .status(200)
      .type("html")
      .send(`<!doctype html>
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
</html>`);
  } catch (error) {
    res
      .status(400)
      .type("html")
      .send(`<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Microsoft Auth Error</title>
  </head>
  <body style="font-family: ui-sans-serif, system-ui, sans-serif; margin: 2rem;">
    <h1>Microsoft 365 connection failed</h1>
    <pre style="white-space: pre-wrap;">${escapeHtml(String(error))}</pre>
    <p><a href="/">Return to status page</a></p>
  </body>
</html>`);
  }
});

app.get("/auth/microsoft/disconnect", async (_req, res) => {
  await microsoftAuth.disconnect();
  res.redirect("/");
});

async function main(): Promise<void> {
  app.listen(config.port, "localhost", () => {
    console.error(`Claude M365 MCP local helper listening on port ${config.port}`);
    console.error(`Local helper URL: ${config.localBaseUrl.toString()}`);
    console.error(`Microsoft callback URI: ${config.microsoft.redirectUri}`);
  });

  const server = createMcpServer({
    config,
    microsoftAuth,
    graph,
  });
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Claude M365 MCP stdio server ready");
}

main().catch((error) => {
  console.error("Fatal startup error:", error);
  process.exit(1);
});
