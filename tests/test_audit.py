from __future__ import annotations

from m365_mcp.audit import audit_metadata, classify_tool, redact_error_message


def test_sharing_mutations_are_audited_as_writes() -> None:
    assert classify_tool("sharepoint_grant_access") == "write"
    assert classify_tool("sharepoint_revoke_permission") == "write"


def test_sharing_audit_keeps_ids_and_redacts_sensitive_values() -> None:
    arguments = {
        "driveId": "drive-1",
        "itemId": "item-1",
        "permissionId": "permission-1",
        "shareUrl": "https://example.sharepoint.com/private-link",
        "recipients": ["ada@example.com"],
        "password": "secret-value",
        "message": "Private invitation text",
    }
    metadata = audit_metadata("sharepoint_revoke_permission", arguments)
    assert metadata["ids"] == {
        "driveId": "drive-1",
        "itemId": "item-1",
        "permissionId": "permission-1",
    }
    redacted = redact_error_message(
        "https://example.sharepoint.com/private-link ada@example.com "
        "secret-value Private invitation text",
        arguments,
    )
    assert "example.sharepoint.com" not in redacted
    assert "ada@example.com" not in redacted
    assert "secret-value" not in redacted
    assert "Private invitation text" not in redacted


def test_audit_redacts_normalized_recipient_values() -> None:
    redacted = redact_error_message(
        "Graph rejected ada@example.com",
        {"recipients": ["  ada@example.com  "]},
    )
    assert "ada@example.com" not in redacted
