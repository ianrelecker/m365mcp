from __future__ import annotations

import pytest

from m365_mcp.pid_extract import LocalPidExtractor
from m365_mcp.pid_policy import (
    REDACTED,
    BlockedError,
    Location,
    PidPolicy,
    labels_from_graph,
)


def test_disabled_policy_allows_everything() -> None:
    policy = PidPolicy.disabled()
    policy.require_mailbox("secret@example.com")
    policy.require_location(Location(drive_id="investor-drive"))
    assert policy.redact_text("SSN 123-45-6789") == "SSN 123-45-6789"


def test_mailbox_allowlist_is_fail_closed() -> None:
    policy = PidPolicy(
        enabled=True,
        mailbox_allowlist=["partners@example.com", "ebremner@example.com"],
        mailbox_blocklist=["blocked@example.com"],
    )
    with pytest.raises(BlockedError) as blocked:
        policy.require_mailbox(None)
    assert blocked.value.reason == "mailbox_not_allowlisted"
    policy.require_mailbox("ebremner@example.com")
    policy.require_mailbox("partners@example.com")
    with pytest.raises(BlockedError) as blocked:
        policy.require_mailbox("other@example.com")
    assert blocked.value.reason == "mailbox_not_allowlisted"
    with pytest.raises(BlockedError) as blocked:
        policy.require_mailbox("blocked@example.com")
    assert blocked.value.reason == "mailbox_blocklisted"
    assert "partners@" not in str(blocked.value)


def test_empty_mailbox_allowlist_blocks_all_mail() -> None:
    policy = PidPolicy(enabled=True)
    with pytest.raises(BlockedError) as blocked:
        policy.require_mailbox("anyone@example.com")
    assert blocked.value.reason == "mailbox_not_allowlisted"


def test_sharepoint_fail_closed_without_allowlist() -> None:
    policy = PidPolicy(enabled=True)
    with pytest.raises(BlockedError) as blocked:
        policy.require_location(Location(drive_id="drive-1", path="/docs"))
    assert blocked.value.reason == "location_not_allowlisted"


def test_drive_allowlist_and_folder_blocklist() -> None:
    policy = PidPolicy(
        enabled=True,
        drive_allowlist=["drive-ok"],
        location_blocklist=["investor-pid", "subscription-agreements"],
    )
    policy.require_location(Location(drive_id="drive-ok", path="/Public"))
    with pytest.raises(BlockedError) as blocked:
        policy.require_location(
            Location(drive_id="drive-ok", path="/Funds/investor-pid/Jane")
        )
    assert blocked.value.reason == "location_blocklisted"
    assert "Jane" not in str(blocked.value)


def test_id_only_location_is_blocked_when_folder_rules_exist() -> None:
    policy = PidPolicy(
        enabled=True,
        drive_allowlist=["drive-ok"],
        folder_allowlist=["approved"],
        location_blocklist=["investors"],
    )
    with pytest.raises(BlockedError):
        policy.require_location(
            Location(drive_id="drive-ok", item_id="01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K")
        )


def test_folder_allowlist_matches_path_not_filename() -> None:
    policy = PidPolicy(
        enabled=True,
        folder_allowlist=["shared documents/approved"],
    )
    policy.require_location(
        Location(path="/drive/root:/Shared Documents/Approved/Q1")
    )
    with pytest.raises(BlockedError):
        policy.require_location(Location(path="/drive/root:/Investors/PID"))
    with pytest.raises(BlockedError):
        policy.require_location(
            Location(
                path="/drive/root:/Investors/approved.xlsx",
                web_url="https://contoso.sharepoint.com/Investors/approved.xlsx",
            )
        )


def test_sensitivity_label_block() -> None:
    policy = PidPolicy(
        enabled=True,
        drive_allowlist=["drive-ok"],
        blocked_sensitivity_labels=["highly confidential", "label-id-1"],
    )
    policy.require_location(Location(drive_id="drive-ok", path="/Public"))
    with pytest.raises(BlockedError) as blocked:
        policy.require_location(
            Location(
                drive_id="drive-ok",
                path="/Public",
                labels=("Highly Confidential",),
            )
        )
    assert blocked.value.reason == "sensitivity_label"


def test_unredactable_content_is_blocked() -> None:
    policy = PidPolicy(enabled=True, mailbox_allowlist=["user@example.com"])
    with pytest.raises(BlockedError) as blocked:
        policy.require_unredactable_content()
    assert blocked.value.reason == "unredactable_content"


def test_redacts_ssn_and_ein_only() -> None:
    policy = PidPolicy(enabled=True)
    text = "SSN 123-45-6789 EIN 12-3456789 invoice 123456789"
    redacted = policy.redact_text(text)
    assert "123-45-6789" not in redacted
    assert "12-3456789" not in redacted
    assert "123456789" in redacted
    assert REDACTED in redacted


def test_labels_from_graph_payload() -> None:
    labels = labels_from_graph(
        {
            "sensitivityLabel": {
                "id": "label-id-1",
                "displayName": "Confidential",
            }
        }
    )
    assert "label-id-1" in labels
    assert "Confidential" in labels
    extracted = labels_from_graph(
        {"labels": [{"sensitivityLabelId": "mip-9", "displayName": "Secret"}]}
    )
    assert "mip-9" in extracted
    assert "Secret" in extracted


def test_local_extractor_is_unavailable() -> None:
    extractor = LocalPidExtractor()
    assert extractor.is_available() is False
    with pytest.raises(BlockedError) as blocked:
        extractor.extract()
    assert blocked.value.reason == "pid_extractor_unavailable"
