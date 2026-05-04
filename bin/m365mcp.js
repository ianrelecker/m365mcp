#!/usr/bin/env node
import { spawn } from "node:child_process";
import { randomBytes } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const packageRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const stateDir = resolve(expandHome(process.env.M365_MCP_HOME || join(homedir(), ".m365mcp")));
const parsedArgs = parseLauncherArgs(process.argv.slice(2));
const envFile = resolve(expandHome(parsedArgs.envFile || join(stateDir, ".env")));
const uvCommand = findUvCommand();
const userArgs = parsedArgs.serverArgs;
const requiredEnv = [
  "MICROSOFT_TENANT_ID",
  "MICROSOFT_CLIENT_ID",
  "MICROSOFT_CLIENT_SECRET",
  "TOKEN_ENCRYPTION_KEY",
];

function expandHome(value) {
  if (value === "~") {
    return homedir();
  }
  if (value.startsWith("~/") || value.startsWith("~\\")) {
    return join(homedir(), value.slice(2));
  }
  return value;
}

function parseLauncherArgs(args) {
  const serverArgs = [];
  let envFilePath = process.env.M365_MCP_ENV_FILE || null;
  let force = false;

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (arg === "--env-file") {
      envFilePath = args[index + 1];
      index += 1;
      continue;
    }
    if (arg.startsWith("--env-file=")) {
      envFilePath = arg.slice("--env-file=".length);
      continue;
    }
    if (arg === "--force") {
      force = true;
      continue;
    }
    serverArgs.push(arg);
  }

  return { envFile: envFilePath, force, serverArgs };
}

function findUvCommand() {
  if (process.env.UV_PATH || process.env.UV) {
    return process.env.UV_PATH || process.env.UV;
  }

  const home = homedir();
  const candidates = process.platform === "win32"
    ? [
        join(home, ".local", "bin", "uv.exe"),
        join(process.env.LOCALAPPDATA || "", "uv", "uv.exe"),
        join(process.env.APPDATA || "", "uv", "uv.exe"),
      ]
    : [
        join(home, ".local", "bin", "uv"),
        join(home, ".local", "uv-bin", "uv"),
        join(home, ".cargo", "bin", "uv"),
        "/opt/homebrew/bin/uv",
        "/usr/local/bin/uv",
      ];

  return candidates.find((candidate) => candidate && existsSync(candidate)) || "uv";
}

function printHelp() {
  console.log(`M365 MCP

Usage:
  m365mcp              Run the MCP stdio server
  m365mcp init         Create ~/.m365mcp/.env and print a Claude config
  m365mcp doctor       Check local runner configuration
  m365mcp auth         Start the local auth helper and open Microsoft sign-in
  m365mcp status       Print Microsoft connection status as JSON
  m365mcp logout       Clear the local Microsoft token cache

Environment:
  M365_MCP_HOME        Working directory for token/cache files (default: ~/.m365mcp)
  M365_MCP_ENV_FILE    Env file to load before starting the server (default: ~/.m365mcp/.env)
  UV_PATH or UV        Path to uv if it is not on PATH

Examples:
  npx -y @ianrelecker/m365mcp init
  npx -y @ianrelecker/m365mcp auth
`);
}

function ensureStateDir() {
  if (!existsSync(stateDir)) {
    mkdirSync(stateDir, { recursive: true });
  }
}

function parseEnvFile(content) {
  const values = {};
  for (const rawLine of content.replace(/^\uFEFF/, "").split(/\r?\n/)) {
    let line = rawLine.trim();
    if (!line || line.startsWith("#")) {
      continue;
    }
    if (line.startsWith("export ")) {
      line = line.slice("export ".length).trim();
    }

    const equalsIndex = line.indexOf("=");
    if (equalsIndex === -1) {
      continue;
    }

    const key = line.slice(0, equalsIndex).trim();
    let value = line.slice(equalsIndex + 1).trim();
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) {
      continue;
    }

    if (
      (value.startsWith("\"") && value.endsWith("\"")) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    values[key] = value;
  }
  return values;
}

function loadEnvFileIfPresent() {
  if (!existsSync(envFile)) {
    return false;
  }

  const values = parseEnvFile(readFileSync(envFile, "utf-8"));
  for (const [key, value] of Object.entries(values)) {
    if (process.env[key] === undefined) {
      process.env[key] = value;
    }
  }
  return true;
}

function isValidEncryptionKey(value) {
  try {
    const decoded = Buffer.from(value, "base64");
    return decoded.length === 32 && decoded.toString("base64") === value;
  } catch {
    return false;
  }
}

function configProblems() {
  const problems = requiredEnv.filter((key) => {
    const value = process.env[key];
    return !value || value === "replace-me" || value.startsWith("your-");
  });

  const tokenKey = process.env.TOKEN_ENCRYPTION_KEY;
  if (
    tokenKey &&
    tokenKey !== "replace-me" &&
    !tokenKey.startsWith("your-") &&
    !isValidEncryptionKey(tokenKey)
  ) {
    problems.push("TOKEN_ENCRYPTION_KEY must be a base64-encoded 32-byte key");
  }

  return problems;
}

function createEnvTemplate() {
  return `PORT=8787
LOCAL_BASE_URL=http://localhost:8787

# Microsoft Entra app registration
MICROSOFT_TENANT_ID=your-tenant-id
MICROSOFT_CLIENT_ID=your-client-id
MICROSOFT_CLIENT_SECRET=your-client-secret-value

# Generated by m365mcp init. Rotate this by deleting the token cache and reconnecting.
TOKEN_ENCRYPTION_KEY=${randomBytes(32).toString("base64")}

# Optional comma-separated shared/delegated mailboxes, such as shared@company.com
KNOWN_MAILBOXES=

# Local JSONL audit log. Records metadata only, never message bodies or tokens.
M365_AUDIT_LOG_ENABLED=true
M365_AUDIT_LOG_FILE=.audit/m365-mcp-audit.jsonl
`;
}

function printClaudeConfig() {
  const envBlock = {};
  if (process.env.M365_MCP_HOME) {
    envBlock.M365_MCP_HOME = stateDir;
  }
  if (parsedArgs.envFile) {
    envBlock.M365_MCP_ENV_FILE = envFile;
  }

  const server = {
    command: "npx",
    args: ["-y", "@ianrelecker/m365mcp"],
  };
  if (Object.keys(envBlock).length > 0) {
    server.env = envBlock;
  }

  console.log(JSON.stringify({ mcpServers: { m365: server } }, null, 2));
}

function initConfig() {
  ensureStateDir();
  const exists = existsSync(envFile);
  if (exists && !parsedArgs.force) {
    console.log(`M365 MCP env file already exists: ${envFile}`);
    console.log("Run with --force to replace it.");
  } else {
    mkdirSync(dirname(envFile), { recursive: true });
    writeFileSync(envFile, createEnvTemplate(), { encoding: "utf-8", flag: "w" });
    console.log(`${exists ? "Replaced" : "Created"} M365 MCP env file: ${envFile}`);
  }

  console.log(`
Next steps:
1. Edit ${envFile} and fill in the Microsoft tenant, client ID, and client secret.
2. Add this MCP server entry to Claude Desktop:
`);
  printClaudeConfig();
  console.log(`
3. Restart Claude Desktop, then ask Claude to check Microsoft auth status.
4. If needed, run: npx -y @ianrelecker/m365mcp auth
`);
}

function doctor() {
  ensureStateDir();
  const loaded = loadEnvFileIfPresent();
  const problems = configProblems();

  console.log(`M365 MCP doctor

State directory: ${stateDir}
Env file: ${envFile} (${loaded ? "found" : "not found"})
uv command: ${uvCommand}
Config: ${problems.length === 0 ? "ready" : `needs attention: ${problems.join(", ")}`}
`);

  if (problems.length > 0) {
    console.log("Run `npx -y @ianrelecker/m365mcp init`, edit the env file, then restart Claude Desktop.");
    process.exitCode = 1;
  }
}

function preflightConfig() {
  const loaded = loadEnvFileIfPresent();
  const problems = configProblems();
  if (problems.length === 0) {
    return true;
  }

  console.error("M365 MCP is not configured yet.");
  console.error(`Configuration issues: ${problems.join(", ")}`);
  console.error(`Checked env file: ${envFile} (${loaded ? "found" : "not found"})`);
  console.error("Run `npx -y @ianrelecker/m365mcp init`, fill in the env file, then restart Claude Desktop.");
  process.exitCode = 1;
  return false;
}

function run() {
  if (userArgs.includes("--help") || userArgs.includes("-h")) {
    printHelp();
    return;
  }

  const command = userArgs[0];
  if (command === "init" || command === "setup" || command === "configure") {
    initConfig();
    return;
  }
  if (command === "doctor") {
    doctor();
    return;
  }

  ensureStateDir();
  if (!preflightConfig()) {
    return;
  }

  const args = [
    "run",
    "--project",
    packageRoot,
    "--no-dev",
    "python",
    "-m",
    "m365_mcp.server",
    ...userArgs,
  ];

  const child = spawn(uvCommand, args, {
    cwd: stateDir,
    env: {
      ...process.env,
      M365_MCP_ENV_FILE: envFile,
      M365_MCP_HOME: stateDir,
      UV_PROJECT_ENVIRONMENT: process.env.UV_PROJECT_ENVIRONMENT || join(stateDir, ".venv"),
    },
    stdio: "inherit",
  });

  child.on("error", (error) => {
    console.error(`Failed to start uv (${uvCommand}): ${error.message}`);
    console.error("Install uv from https://docs.astral.sh/uv/ or set UV_PATH.");
    process.exitCode = 1;
  });

  child.on("exit", (code, signal) => {
    if (signal) {
      process.kill(process.pid, signal);
      return;
    }
    process.exitCode = code ?? 0;
  });
}

run();
