#!/usr/bin/env node
import { spawn } from "node:child_process";
import { existsSync, mkdirSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const packageRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const stateDir = resolve(process.env.M365_MCP_HOME || join(homedir(), ".m365mcp"));
const uvCommand = findUvCommand();
const userArgs = process.argv.slice(2);

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
  m365mcp auth         Start the local auth helper and open Microsoft sign-in
  m365mcp status       Print Microsoft connection status as JSON
  m365mcp logout       Clear the local Microsoft token cache

Environment:
  M365_MCP_HOME        Working directory for token/cache files (default: ~/.m365mcp)
  UV_PATH or UV        Path to uv if it is not on PATH
`);
}

function run() {
  if (userArgs.includes("--help") || userArgs.includes("-h")) {
    printHelp();
    return;
  }

  if (!existsSync(stateDir)) {
    mkdirSync(stateDir, { recursive: true });
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
    env: process.env,
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
