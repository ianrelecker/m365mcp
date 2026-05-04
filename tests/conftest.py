from __future__ import annotations

import base64
import json
from collections.abc import Callable
from pathlib import Path

import pytest

from m365_mcp.config import AppConfig, MicrosoftConfig

TEST_KEY = bytes(range(32))
TEST_KEY_B64 = base64.b64encode(TEST_KEY).decode("ascii")


def make_jwt(payload: dict[str, object]) -> str:
    header = {"alg": "none", "typ": "JWT"}
    encode = lambda value: base64.urlsafe_b64encode(  # noqa: E731
        json.dumps(value, separators=(",", ":")).encode("utf-8")
    ).decode("ascii").rstrip("=")
    return f"{encode(header)}.{encode(payload)}."


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture
def config_factory(
    tmp_path: Path,
) -> Callable[..., AppConfig]:
    def factory(**overrides: object) -> AppConfig:
        port = int(overrides.pop("port", 8787))
        local_base_url = str(
            overrides.pop("localBaseUrl", f"http://localhost:{port}")
        )
        token_file = Path(
            overrides.pop(
                "tokenFile",
                tmp_path / ".tokens" / "microsoft-graph-token.json",
            )
        )
        audit_log_file = Path(
            overrides.pop(
                "auditLogFile",
                tmp_path / ".audit" / "m365-mcp-audit.jsonl",
            )
        )
        audit_log_enabled = bool(overrides.pop("auditLogEnabled", True))
        microsoft = MicrosoftConfig(
            tenantId=str(overrides.pop("tenantId", "tenant-id")),
            clientId=str(overrides.pop("clientId", "client-id")),
            clientSecret=str(overrides.pop("clientSecret", "client-secret")),
            redirectUri=str(
                overrides.pop(
                    "redirectUri",
                    f"{local_base_url}/auth/microsoft/callback",
                )
            ),
            scopes=list(
                overrides.pop(
                    "scopes",
                    [
                        "openid",
                        "profile",
                        "email",
                        "offline_access",
                        "Mail.ReadWrite",
                        "Mail.ReadWrite.Shared",
                        "Mail.Send",
                        "Mail.Send.Shared",
                        "Calendars.ReadWrite.Shared",
                        "Contacts.ReadWrite.Shared",
                        "MailboxSettings.ReadWrite",
                    ],
                )
            ),
        )

        if overrides:
            raise AssertionError(f"Unexpected config overrides: {sorted(overrides)}")

        return AppConfig(
            port=port,
            localBaseUrl=local_base_url,
            microsoft=microsoft,
            encryptionKey=TEST_KEY,
            knownMailboxes=["shared@example.com"],
            tokenFile=token_file,
            auditLogEnabled=audit_log_enabled,
            auditLogFile=audit_log_file,
        )

    return factory
