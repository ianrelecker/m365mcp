from __future__ import annotations

import json
import os
import stat
import subprocess
from base64 import b64encode
from pathlib import Path


ROOT = Path(__file__).parents[1]


def test_npm_package_exposes_executable_launcher() -> None:
    package = json.loads((ROOT / "package.json").read_text("utf-8"))

    assert package["name"] == "@ianrelecker/m365mcp"
    assert package["private"] is False
    assert package["bin"] == {"m365mcp": "./bin/m365mcp.js"}
    assert "src/m365_mcp/*.py" in package["files"]
    assert "uv.lock" in package["files"]
    assert "CODEX_MCP_SETUP.md" in package["files"]


def test_npm_launcher_is_executable_and_uses_uv_project() -> None:
    launcher = ROOT / "bin" / "m365mcp.js"
    content = launcher.read_text("utf-8")
    mode = launcher.stat().st_mode

    assert content.startswith("#!/usr/bin/env node")
    assert mode & stat.S_IXUSR
    assert '"--project"' in content
    assert '"python"' in content
    assert '"-m"' in content
    assert '"m365_mcp.server"' in content
    assert "M365_MCP_HOME" in content
    assert "M365_MCP_ENV_FILE" in content
    assert "UV_PROJECT_ENVIRONMENT" in content


def test_npm_launcher_init_creates_default_env_file(tmp_path: Path) -> None:
    launcher = ROOT / "bin" / "m365mcp.js"
    env = os.environ.copy()
    env["M365_MCP_HOME"] = str(tmp_path)
    for key in (
        "MICROSOFT_TENANT_ID",
        "MICROSOFT_CLIENT_ID",
        "MICROSOFT_CLIENT_SECRET",
        "TOKEN_ENCRYPTION_KEY",
    ):
        env.pop(key, None)

    result = subprocess.run(
        ["node", str(launcher), "init"],
        check=False,
        capture_output=True,
        env=env,
        text=True,
    )

    assert result.returncode == 0
    env_file = tmp_path / ".env"
    content = env_file.read_text("utf-8")

    assert "MICROSOFT_TENANT_ID=your-tenant-id" in content
    assert "MICROSOFT_CLIENT_ID=your-client-id" in content
    assert "MICROSOFT_CLIENT_SECRET=your-client-secret-value" in content
    assert "TOKEN_ENCRYPTION_KEY=replace-me" not in content
    assert '"command": "npx"' in result.stdout
    assert '"@ianrelecker/m365mcp"' in result.stdout


def test_npm_launcher_doctor_checks_env_file(tmp_path: Path) -> None:
    launcher = ROOT / "bin" / "m365mcp.js"
    env_file = tmp_path / "custom.env"
    token_key = b64encode(b"0" * 32).decode()
    env_file.write_text(
        "\n".join(
            [
                "MICROSOFT_TENANT_ID=tenant",
                "MICROSOFT_CLIENT_ID=client",
                "MICROSOFT_CLIENT_SECRET=secret",
                f"TOKEN_ENCRYPTION_KEY={token_key}",
            ]
        ),
        "utf-8",
    )
    env = os.environ.copy()
    env["M365_MCP_HOME"] = str(tmp_path)
    for key in (
        "MICROSOFT_TENANT_ID",
        "MICROSOFT_CLIENT_ID",
        "MICROSOFT_CLIENT_SECRET",
        "TOKEN_ENCRYPTION_KEY",
    ):
        env.pop(key, None)

    result = subprocess.run(
        ["node", str(launcher), "--env-file", str(env_file), "doctor"],
        check=False,
        capture_output=True,
        env=env,
        text=True,
    )

    assert result.returncode == 0
    assert f"Env file: {env_file} (found)" in result.stdout
    assert "Config: ready" in result.stdout
