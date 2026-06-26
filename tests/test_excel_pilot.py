"""Pilot / acceptance integration test against a *real* fragile workbook.

This is the gating test for the schema-driven match/fill skill. The whole point
of editing through the Graph Workbook API (instead of openpyxl / XML surgery) is
that Excel's own engine preserves fragile workbook parts on write: external
links, drawings, VBA (.xlsm), defined names, and dynamic-array spill formulas.

It is **skipped unless** you point it at a live workbook you own, because it
mutates the file and needs real OAuth tokens. To run it:

    export M365_PILOT_DRIVE_ID=...        # from workbook_resolve
    export M365_PILOT_ITEM_ID=...
    # ...or instead:  export M365_PILOT_SHARE_URL='https://...'
    export M365_PILOT_PLAN=/path/to/plan.json
    uv run pytest tests/test_excel_pilot.py -q -s

`plan.json` describes the scattered writes and the cells to read back, e.g.:

    {
      "writes": [
        {"worksheet": "Model", "address": "Z1", "formulas": [["='Unit Mix'!H11"]]},
        {"worksheet": "Model", "address": "Z2", "formulas": [["=SUM(Z1:Z1)"]]}
        // ... ~50 scattered across multiple sheets
      ],
      "reads":  [{"worksheet": "Model", "address": "Z1:Z2"}],
      "spill":  {"worksheet": "Model", "address": "AB1"}   // optional dyn-array cell
    }

Use a throwaway/scratch region of the workbook for `writes` so the test does not
clobber real model cells. The token store / app config is loaded exactly as the
server loads it (`create_runtime()` -> `load_config()`).
"""

from __future__ import annotations

import io
import json
import os
import zipfile
from pathlib import Path

import httpx
import pytest

pytestmark = pytest.mark.anyio

_DRIVE_ID = os.environ.get("M365_PILOT_DRIVE_ID")
_ITEM_ID = os.environ.get("M365_PILOT_ITEM_ID")
_SHARE_URL = os.environ.get("M365_PILOT_SHARE_URL")
_PLAN = os.environ.get("M365_PILOT_PLAN")

_HAVE_TARGET = bool((_DRIVE_ID and _ITEM_ID) or _SHARE_URL) and bool(_PLAN)

skip_reason = (
    "set M365_PILOT_DRIVE_ID+M365_PILOT_ITEM_ID (or M365_PILOT_SHARE_URL) and "
    "M365_PILOT_PLAN to run the live fragile-workbook pilot"
)


def _fragile_parts(blob: bytes) -> dict[str, object]:
    """Inspect an .xlsx/.xlsm OOXML package and report the fragile parts that
    are most likely to be dropped by a naive (non-Excel) rewrite."""
    with zipfile.ZipFile(io.BytesIO(blob)) as zf:
        names = zf.namelist()
        workbook_xml = b""
        if "xl/workbook.xml" in names:
            workbook_xml = zf.read("xl/workbook.xml")
    return {
        "external_links": sum(
            1 for n in names if n.startswith("xl/externalLinks/") and n.endswith(".xml")
        ),
        "drawings": sum(
            1 for n in names if n.startswith("xl/drawings/") and n.endswith(".xml")
        ),
        "has_vba": "xl/vbaProject.bin" in names,
        "defined_names": workbook_xml.count(b"<definedName "),
        "part_count": len(names),
    }


async def _download(http: httpx.AsyncClient, token: str, drive_id: str, item_id: str) -> bytes:
    resp = await http.get(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content",
        headers={"Authorization": f"Bearer {token}"},
        follow_redirects=True,
    )
    resp.raise_for_status()
    return resp.content


@pytest.mark.skipif(not _HAVE_TARGET, reason=skip_reason)
async def test_pilot_fragile_workbook_round_trip() -> None:
    # Import here so the module still imports cleanly when the server's runtime
    # deps / config are not available in a plain unit-test environment.
    from m365_mcp.server import create_runtime

    plan = json.loads(Path(_PLAN).read_text("utf-8"))
    writes = plan["writes"]
    reads = plan.get("reads", [])
    spill = plan.get("spill")
    assert writes, "plan.json must contain at least one write"

    runtime = create_runtime(start_helper_server=False)
    excel = runtime.excel
    http = runtime.http_client
    try:
        token = await runtime.microsoft_auth.get_access_token()

        if _SHARE_URL:
            item = await excel.resolve_workbook(shareUrl=_SHARE_URL)
        else:
            item = await excel.resolve_workbook(driveId=_DRIVE_ID, itemId=_ITEM_ID)
        assert item.driveId and item.itemId

        # (d-baseline) snapshot the fragile parts before we touch anything.
        before = _fragile_parts(await _download(http, token, item.driveId, item.itemId))
        print("PILOT baseline fragile parts:", before)

        # (a) batch-write scattered formulas across sheets in ONE session.
        session = await excel.create_session(item, persistChanges=True)
        try:
            write_result = await excel.update_ranges(
                item, updates=writes, sessionId=session.sessionId
            )
            failed = [r for r in write_result.ranges if r.error]
            assert not failed, f"batch write reported errors: {failed}"

            # (b) force a full recalc inside the same session.
            await excel.calculate(item, calculationType="Full", sessionId=session.sessionId)

            # (c) batch-read back formulas AND computed values.
            read_specs = list(reads)
            if spill:
                read_specs.append(spill)
            read_result = await excel.get_ranges(
                item, ranges=read_specs, sessionId=session.sessionId
            )
            read_errors = [r for r in read_result.ranges if r.error]
            assert not read_errors, f"batch read reported errors: {read_errors}"
            for r in read_result.ranges:
                # formulas round-tripped and a computed value is present
                assert r.formulas is not None, f"no formulas for {r.address}"
                assert r.values is not None, f"no computed values for {r.address}"
                print(f"PILOT read {r.address}: formulas={r.formulas} values={r.values}")

            if spill:
                spilled = read_result.ranges[-1]
                # a dynamic-array spill should yield more than a single cell
                cell_count = sum(len(row) for row in (spilled.values or []))
                print(f"PILOT spill {spilled.address} cell_count={cell_count}")
                assert cell_count >= 1
        finally:
            await excel.close_session(item, sessionId=session.sessionId)

        # (d) re-download and confirm the fragile parts survived intact.
        blob = await _download(http, token, item.driveId, item.itemId)
        # zipfile.testzip() returning None means no CRC errors == not corrupt.
        assert zipfile.ZipFile(io.BytesIO(blob)).testzip() is None, "package corrupt"
        after = _fragile_parts(blob)
        print("PILOT post-write fragile parts:", after)

        problems = []
        if after["external_links"] < before["external_links"]:
            problems.append("external links dropped")
        if after["drawings"] < before["drawings"]:
            problems.append("drawings dropped")
        if before["has_vba"] and not after["has_vba"]:
            problems.append("VBA project dropped")
        if after["defined_names"] < before["defined_names"]:
            problems.append("defined names dropped")
        assert not problems, (
            f"Graph round-trip corrupted fragile parts: {problems}. "
            f"before={before} after={after}"
        )
    finally:
        if runtime.owns_http_client:
            await http.aclose()
