from __future__ import annotations

import json
import stat
from pathlib import Path


ROOT = Path(__file__).parents[1]


def test_npm_package_exposes_executable_launcher() -> None:
    package = json.loads((ROOT / "package.json").read_text("utf-8"))

    assert package["name"] == "@ianrelecker/m365mcp"
    assert package["private"] is False
    assert package["bin"] == {"m365mcp": "./bin/m365mcp.js"}
    assert "src/m365_mcp/*.py" in package["files"]
    assert "uv.lock" in package["files"]


def test_npm_launcher_is_executable_and_uses_uv_project() -> None:
    launcher = ROOT / "bin" / "m365mcp.js"
    content = launcher.read_text("utf-8")
    mode = launcher.stat().st_mode

    assert content.startswith("#!/usr/bin/env node")
    assert mode & stat.S_IXUSR
    assert '"--project"' in content
    assert '"m365-mcp"' in content
    assert "M365_MCP_HOME" in content
