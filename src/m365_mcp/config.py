from __future__ import annotations

import base64
import os
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urljoin, urlparse

from dotenv import load_dotenv

load_dotenv()


def _require_env(env: Mapping[str, str], name: str) -> str:
    value = env.get(name)
    if not value:
        raise ValueError(f"Missing required environment variable: {name}")
    return value


def _optional_comma_list(value: str | None) -> list[str]:
    if not value:
        return []

    return [item.strip() for item in value.split(",") if item.strip()]


def _parse_encryption_key(env: Mapping[str, str], name: str) -> bytes:
    value = _require_env(env, name)
    try:
        key = base64.b64decode(value, validate=True)
    except Exception as exc:  # pragma: no cover - exact decoder errors vary
        raise ValueError(f"{name} must be a base64-encoded 32-byte key") from exc

    if len(key) != 32:
        raise ValueError(f"{name} must be a base64-encoded 32-byte key")

    return key


def _parse_url(value: str, *, name: str) -> str:
    parsed = urlparse(value)
    if not parsed.scheme or not parsed.netloc:
        raise ValueError(f"{name} must be a valid absolute URL")
    return value


@dataclass(frozen=True)
class MicrosoftConfig:
    tenantId: str
    clientId: str
    clientSecret: str
    redirectUri: str
    scopes: list[str]


@dataclass(frozen=True)
class AppConfig:
    port: int
    localBaseUrl: str
    microsoft: MicrosoftConfig
    encryptionKey: bytes
    knownMailboxes: list[str]
    tokenFile: Path


def build_config_from_env(env: Mapping[str, str] | None = None) -> AppConfig:
    source = dict(os.environ if env is None else env)
    port = int(source.get("PORT", "8787"))
    local_base_url = _parse_url(
        source.get("LOCAL_BASE_URL", f"http://localhost:{port}"),
        name="LOCAL_BASE_URL",
    )

    return AppConfig(
        port=port,
        localBaseUrl=local_base_url,
        microsoft=MicrosoftConfig(
            tenantId=_require_env(source, "MICROSOFT_TENANT_ID"),
            clientId=_require_env(source, "MICROSOFT_CLIENT_ID"),
            clientSecret=_require_env(source, "MICROSOFT_CLIENT_SECRET"),
            redirectUri=urljoin(local_base_url, "/auth/microsoft/callback"),
            scopes=[
                "openid",
                "profile",
                "email",
                "offline_access",
                "Mail.ReadWrite.Shared",
                "Mail.Send.Shared",
                "Calendars.ReadWrite.Shared",
            ],
        ),
        encryptionKey=_parse_encryption_key(source, "TOKEN_ENCRYPTION_KEY"),
        knownMailboxes=_optional_comma_list(source.get("KNOWN_MAILBOXES")),
        tokenFile=Path(".tokens/microsoft-graph-token.json"),
    )


def load_config() -> AppConfig:
    return build_config_from_env()

