from __future__ import annotations

import base64
import os
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import urljoin, urlparse

from dotenv import load_dotenv

load_dotenv()


GRAPH_SCOPES_WITHOUT_SEND = [
    "openid",
    "profile",
    "email",
    "offline_access",
    "Mail.ReadWrite",
    "Mail.ReadWrite.Shared",
    "Calendars.ReadWrite.Shared",
    "Contacts.ReadWrite.Shared",
    "MailboxSettings.ReadWrite",
    # Sites.Read.All (not ReadWrite): the SharePoint tools only
    # browse sites/drives read-only. Workbook edits go through
    # /drives/{id}/items/{id}/workbook, which is governed by
    # Files.ReadWrite.All, so no SharePoint *write* scope is needed.
    "Sites.Read.All",
    "Files.ReadWrite.All",
]

GRAPH_SEND_SCOPES = [
    "Mail.Send",
    "Mail.Send.Shared",
]


def graph_scopes(*, mail_send_enabled: bool) -> list[str]:
    scopes = [
        "openid",
        "profile",
        "email",
        "offline_access",
        "Mail.ReadWrite",
        "Mail.ReadWrite.Shared",
    ]
    if mail_send_enabled:
        scopes.extend(GRAPH_SEND_SCOPES)
    scopes.extend(
        [
            "Calendars.ReadWrite.Shared",
            "Contacts.ReadWrite.Shared",
            "MailboxSettings.ReadWrite",
            "Sites.Read.All",
            "Files.ReadWrite.All",
        ]
    )
    return scopes


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


def _parse_bool(value: str | None, *, default: bool, name: str) -> bool:
    if value is None or value == "":
        return default

    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "on"}:
        return True
    if normalized in {"0", "false", "no", "off"}:
        return False
    raise ValueError(f"{name} must be true or false")


def mail_send_enabled_from_env(env: Mapping[str, str] | None = None) -> bool:
    source = os.environ if env is None else env
    return _parse_bool(
        source.get("M365_MAIL_SEND_ENABLED"),
        default=False,
        name="M365_MAIL_SEND_ENABLED",
    )


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
    auditLogEnabled: bool = True
    auditLogFile: Path = Path(".audit/m365-mcp-audit.jsonl")
    mailSendEnabled: bool = False
    pidSafeMode: bool = False
    pidMailboxAllowlist: list[str] = field(default_factory=list)
    pidMailboxBlocklist: list[str] = field(default_factory=list)
    pidSiteAllowlist: list[str] = field(default_factory=list)
    pidDriveAllowlist: list[str] = field(default_factory=list)
    pidFolderAllowlist: list[str] = field(default_factory=list)
    pidLocationBlocklist: list[str] = field(default_factory=list)
    pidBlockedSensitivityLabels: list[str] = field(default_factory=list)
    pidRedactIdentifiers: bool = True
    pidLocalExtractorEnabled: bool = False


def build_config_from_env(env: Mapping[str, str] | None = None) -> AppConfig:
    source = dict(os.environ if env is None else env)
    port = int(source.get("PORT", "8787"))
    local_base_url = _parse_url(
        source.get("LOCAL_BASE_URL", f"http://localhost:{port}"),
        name="LOCAL_BASE_URL",
    )
    mail_send_enabled = mail_send_enabled_from_env(source)

    return AppConfig(
        port=port,
        localBaseUrl=local_base_url,
        microsoft=MicrosoftConfig(
            tenantId=_require_env(source, "MICROSOFT_TENANT_ID"),
            clientId=_require_env(source, "MICROSOFT_CLIENT_ID"),
            clientSecret=_require_env(source, "MICROSOFT_CLIENT_SECRET"),
            redirectUri=urljoin(local_base_url, "/auth/microsoft/callback"),
            scopes=graph_scopes(mail_send_enabled=mail_send_enabled),
        ),
        encryptionKey=_parse_encryption_key(source, "TOKEN_ENCRYPTION_KEY"),
        knownMailboxes=_optional_comma_list(source.get("KNOWN_MAILBOXES")),
        tokenFile=Path(".tokens/microsoft-graph-token.json"),
        auditLogEnabled=_parse_bool(
            source.get("M365_AUDIT_LOG_ENABLED"),
            default=True,
            name="M365_AUDIT_LOG_ENABLED",
        ),
        auditLogFile=Path(
            source.get("M365_AUDIT_LOG_FILE", ".audit/m365-mcp-audit.jsonl")
        ),
        mailSendEnabled=mail_send_enabled,
        pidSafeMode=_parse_bool(
            source.get("M365_PID_SAFE_MODE"),
            default=False,
            name="M365_PID_SAFE_MODE",
        ),
        pidMailboxAllowlist=_optional_comma_list(
            source.get("M365_PID_MAILBOX_ALLOWLIST")
        ),
        pidMailboxBlocklist=_optional_comma_list(
            source.get("M365_PID_MAILBOX_BLOCKLIST")
        ),
        pidSiteAllowlist=_optional_comma_list(source.get("M365_PID_SITE_ALLOWLIST")),
        pidDriveAllowlist=_optional_comma_list(source.get("M365_PID_DRIVE_ALLOWLIST")),
        pidFolderAllowlist=_optional_comma_list(
            source.get("M365_PID_FOLDER_ALLOWLIST")
        ),
        pidLocationBlocklist=_optional_comma_list(
            source.get("M365_PID_LOCATION_BLOCKLIST")
        ),
        pidBlockedSensitivityLabels=_optional_comma_list(
            source.get("M365_PID_BLOCKED_SENSITIVITY_LABELS")
        ),
        pidRedactIdentifiers=_parse_bool(
            source.get("M365_PID_REDACT_IDENTIFIERS"),
            default=True,
            name="M365_PID_REDACT_IDENTIFIERS",
        ),
        pidLocalExtractorEnabled=_parse_bool(
            source.get("M365_PID_LOCAL_EXTRACTOR_ENABLED"),
            default=False,
            name="M365_PID_LOCAL_EXTRACTOR_ENABLED",
        ),
    )


def load_config() -> AppConfig:
    return build_config_from_env()
