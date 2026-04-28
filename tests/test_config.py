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
