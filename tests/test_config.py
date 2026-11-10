from __future__ import annotations

from m365_mcp.config import build_config_from_env

from .conftest import TEST_KEY_B64


def test_build_config_parses_defaults_and_known_mailboxes() -> None:
    config = build_config_from_env(
        {
            "MICROSOFT_TENANT_ID": "tenant",
            "MICROSOFT_CLIENT_ID": "client",
            "MICROSOFT_CLIENT_SECRET": "secret",
            "TOKEN_ENCRYPTION_KEY": TEST_KEY_B64,
            "KNOWN_MAILBOXES": " shared@company.com , second@company.com ",
            "M365_AUDIT_LOG_ENABLED": "false",
            "M365_AUDIT_LOG_FILE": "custom-audit.jsonl",
        }
    )

    assert config.port == 8787
    assert config.localBaseUrl == "http://localhost:8787"
    assert config.microsoft.redirectUri == "http://localhost:8787/auth/microsoft/callback"
    assert config.knownMailboxes == ["shared@company.com", "second@company.com"]
    assert config.encryptionKey == bytes(range(32))
    assert "Contacts.ReadWrite.Shared" in config.microsoft.scopes
    assert "MailboxSettings.ReadWrite" in config.microsoft.scopes
    assert "Mail.ReadWrite" in config.microsoft.scopes
    assert "Mail.Send" not in config.microsoft.scopes
    assert "Mail.Send.Shared" not in config.microsoft.scopes
    assert config.mailSendEnabled is False
    assert config.pidSafeMode is False
    assert config.auditLogEnabled is False
    assert str(config.auditLogFile) == "custom-audit.jsonl"


def test_build_config_enables_mail_send_and_pid_safe_mode() -> None:
    config = build_config_from_env(
        {
            "MICROSOFT_TENANT_ID": "tenant",
            "MICROSOFT_CLIENT_ID": "client",
            "MICROSOFT_CLIENT_SECRET": "secret",
            "TOKEN_ENCRYPTION_KEY": TEST_KEY_B64,
            "M365_MAIL_SEND_ENABLED": "true",
            "M365_PID_SAFE_MODE": "true",
            "M365_PID_MAILBOX_ALLOWLIST": " partners@example.com ",
            "M365_PID_LOCATION_BLOCKLIST": "investor-pid,subscription-agreements",
            "M365_PID_BLOCKED_SENSITIVITY_LABELS": "Highly Confidential",
        }
    )

    assert config.mailSendEnabled is True
    assert "Mail.Send" in config.microsoft.scopes
    assert "Mail.Send.Shared" in config.microsoft.scopes
    assert config.pidSafeMode is True
    assert config.pidMailboxAllowlist == ["partners@example.com"]
    assert config.pidLocationBlocklist == [
        "investor-pid",
        "subscription-agreements",
    ]
    assert config.pidBlockedSensitivityLabels == ["Highly Confidential"]


def test_build_config_rejects_invalid_encryption_key() -> None:
    try:
        build_config_from_env(
            {
                "MICROSOFT_TENANT_ID": "tenant",
                "MICROSOFT_CLIENT_ID": "client",
                "MICROSOFT_CLIENT_SECRET": "secret",
                "TOKEN_ENCRYPTION_KEY": "not-base64",
            }
        )
    except ValueError as error:
        assert "TOKEN_ENCRYPTION_KEY" in str(error)
    else:  # pragma: no cover - defensive
        raise AssertionError("Expected invalid encryption key to fail")


def test_build_config_rejects_invalid_audit_bool() -> None:
    try:
        build_config_from_env(
            {
                "MICROSOFT_TENANT_ID": "tenant",
                "MICROSOFT_CLIENT_ID": "client",
                "MICROSOFT_CLIENT_SECRET": "secret",
                "TOKEN_ENCRYPTION_KEY": TEST_KEY_B64,
                "M365_AUDIT_LOG_ENABLED": "maybe",
            }
        )
    except ValueError as error:
        assert "M365_AUDIT_LOG_ENABLED" in str(error)
    else:  # pragma: no cover - defensive
        raise AssertionError("Expected invalid audit bool to fail")
