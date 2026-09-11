"""Regenerate the ``tools`` array in ``manifest.json`` from the live MCP registry.

The MCPB manifest advertises every tool the bundle provides so the installing
client can show them before the server ever runs. Rather than hand-maintaining
that list next to ~80 tool definitions in ``server.py``, generate it:

    uv run python scripts/sync_mcpb_tools.py          # rewrite manifest.json
    uv run python scripts/sync_mcpb_tools.py --check  # fail if out of date

``tests/test_manifest.py`` runs the same comparison, so a new tool that is not
synced here fails the test suite.
"""

from __future__ import annotations

import argparse
import asyncio
import json
import os
import sys
from pathlib import Path

os.environ["M365_MAIL_SEND_ENABLED"] = "false"

REPO_ROOT = Path(__file__).resolve().parents[1]
MANIFEST_PATH = REPO_ROOT / "manifest.json"

# Manifest descriptions are one-line summaries for an install dialog, not the
# full model-facing tool descriptions.
MAX_DESCRIPTION_LENGTH = 200


def _summarize(description: str | None) -> str:
    text = " ".join((description or "").split())
    if len(text) <= MAX_DESCRIPTION_LENGTH:
        return text

    truncated = text[: MAX_DESCRIPTION_LENGTH - 1]
    cutoff = truncated.rfind(" ")
    if cutoff > 0:
        truncated = truncated[:cutoff]
    return truncated.rstrip(" ,;:") + "…"


def build_tools() -> list[dict[str, str]]:
    from m365_mcp.server import mcp

    tools = asyncio.run(mcp.list_tools())
    return [
        {"name": tool.name, "description": _summarize(tool.description)}
        for tool in tools
    ]


def load_manifest() -> dict:
    return json.loads(MANIFEST_PATH.read_text("utf-8"))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--check",
        action="store_true",
        help="exit non-zero if manifest.json is out of date instead of rewriting it",
    )
    args = parser.parse_args(argv)

    manifest = load_manifest()
    tools = build_tools()

    if manifest.get("tools") == tools:
        print(f"manifest.json is up to date ({len(tools)} tools)")
        return 0

    if args.check:
        print(
            "manifest.json tools are out of date. "
            "Run: uv run python scripts/sync_mcpb_tools.py",
            file=sys.stderr,
        )
        return 1

    manifest["tools"] = tools
    MANIFEST_PATH.write_text(json.dumps(manifest, indent=2) + "\n", "utf-8")
    print(f"Wrote {len(tools)} tools to manifest.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
