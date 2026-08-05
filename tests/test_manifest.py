from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "scripts"))

from sync_mcpb_tools import build_tools, load_manifest  # noqa: E402


@pytest.fixture(scope="module")
def manifest() -> dict:
    return load_manifest()


def test_manifest_declares_every_tool(manifest: dict) -> None:
    assert manifest["tools"] == build_tools(), (
        "manifest.json tools are stale. "
        "Run: uv run python scripts/sync_mcpb_tools.py"
    )


def test_manifest_version_matches_pyproject(manifest: dict) -> None:
    pyproject = (REPO_ROOT / "pyproject.toml").read_text("utf-8")
    version_line = next(
        line for line in pyproject.splitlines() if line.startswith("version = ")
    )
    assert manifest["version"] == version_line.split('"')[1]


def test_manifest_runs_the_packaged_entry_point(manifest: dict) -> None:
    server = manifest["server"]
    assert server["type"] == "uv"
    assert (REPO_ROOT / server["entry_point"]).is_file()

    config = server["mcp_config"]
    assert config["command"] == "uv"
    assert config["args"] == [
        "run",
        "--directory",
        "${__dirname}",
        server["entry_point"],
    ]


def test_manifest_env_covers_required_config(manifest: dict) -> None:
    env = manifest["server"]["mcp_config"]["env"]
    user_config = manifest["user_config"]

    required_env = {
        "MICROSOFT_TENANT_ID",
        "MICROSOFT_CLIENT_ID",
        "MICROSOFT_CLIENT_SECRET",
        "TOKEN_ENCRYPTION_KEY",
    }
    assert required_env <= env.keys()

    for value in env.values():
        assert value.startswith("${user_config.") and value.endswith("}")
        key = value[len("${user_config.") : -1]
        assert key in user_config, f"env references unknown user_config key: {key}"

    for name in required_env:
        key = env[name][len("${user_config.") : -1]
        assert user_config[key]["required"] is True


def test_manifest_marks_secrets_sensitive(manifest: dict) -> None:
    user_config = manifest["user_config"]
    for key in ("microsoft_client_secret", "token_encryption_key"):
        assert user_config[key]["sensitive"] is True


def test_mcpbignore_excludes_local_secrets() -> None:
    ignored = {
        line.strip()
        for line in (REPO_ROOT / ".mcpbignore").read_text("utf-8").splitlines()
        if line.strip() and not line.startswith("#")
    }
    assert {".env", ".tokens/", ".audit/"} <= ignored


def test_manifest_is_formatted_as_written_by_the_sync_script(manifest: dict) -> None:
    raw = (REPO_ROOT / "manifest.json").read_text("utf-8")
    assert raw == json.dumps(manifest, indent=2) + "\n"
