from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any
from urllib.parse import urlparse

SSN_RE = re.compile(r"\b\d{3}[-\s]\d{2}[-\s]\d{4}\b")
EIN_RE = re.compile(r"\b\d{2}-\d{7}\b")
REDACTED = "[redacted-id]"


class BlockedError(ValueError):
    def __init__(self, reason: str) -> None:
        self.reason = reason
        super().__init__(
            f"Request blocked ({reason}). Protected content was not returned."
        )


@dataclass(frozen=True)
class Location:
    site_id: str | None = None
    drive_id: str | None = None
    item_id: str | None = None
    path: str | None = None
    web_url: str | None = None
    labels: tuple[str, ...] = ()


def _normalize(value: str | None) -> str:
    return (value or "").strip().lower()


def _dir_path(value: str | None) -> str:
    raw = _normalize(value).replace("\\", "/")
    if not raw:
        return ""
    raw = raw.split("?", 1)[0].split("#", 1)[0]
    if "/root:" in raw:
        raw = raw.split("/root:", 1)[-1]
    raw = raw.strip("/")
    if not raw:
        return ""
    last = raw.rsplit("/", 1)[-1]
    if "." in last and not last.startswith("."):
        parent = raw.rsplit("/", 1)[0] if "/" in raw else ""
        return parent
    return raw


def _url_dir(value: str | None) -> str:
    raw = _normalize(value)
    if not raw:
        return ""
    if "://" in raw:
        parsed = urlparse(raw)
        return _dir_path(parsed.path)
    return _dir_path(raw)


def _folder_prefix_match(pathish: str, allowlist: list[str]) -> bool:
    if not pathish:
        return False
    candidate = pathish.strip("/")
    for entry in allowlist:
        needle = _dir_path(entry) or _normalize(entry).strip("/")
        if not needle:
            continue
        if candidate == needle or candidate.startswith(needle + "/"):
            return True
        if f"/{needle}/" in f"/{candidate}/":
            return True
    return False


def _exact_or_id_match(values: list[str], allowlist: list[str]) -> bool:
    for value in values:
        if not value:
            continue
        if value in allowlist:
            return True
    return False


@dataclass(frozen=True)
class PidPolicy:
    enabled: bool = False
    mailbox_allowlist: list[str] = field(default_factory=list)
    mailbox_blocklist: list[str] = field(default_factory=list)
    site_allowlist: list[str] = field(default_factory=list)
    drive_allowlist: list[str] = field(default_factory=list)
    folder_allowlist: list[str] = field(default_factory=list)
    location_blocklist: list[str] = field(default_factory=list)
    blocked_sensitivity_labels: list[str] = field(default_factory=list)
    redact_identifiers: bool = True

    @classmethod
    def disabled(cls) -> PidPolicy:
        return cls(enabled=False)

    @classmethod
    def from_config(cls, config: Any) -> PidPolicy:
        return cls(
            enabled=bool(getattr(config, "pidSafeMode", False)),
            mailbox_allowlist=_normalized_list(
                getattr(config, "pidMailboxAllowlist", [])
            ),
            mailbox_blocklist=_normalized_list(
                getattr(config, "pidMailboxBlocklist", [])
            ),
            site_allowlist=_normalized_list(getattr(config, "pidSiteAllowlist", [])),
            drive_allowlist=_normalized_list(getattr(config, "pidDriveAllowlist", [])),
            folder_allowlist=_normalized_list(
                getattr(config, "pidFolderAllowlist", [])
            ),
            location_blocklist=_normalized_list(
                getattr(config, "pidLocationBlocklist", [])
            ),
            blocked_sensitivity_labels=_normalized_list(
                getattr(config, "pidBlockedSensitivityLabels", [])
            ),
            redact_identifiers=bool(getattr(config, "pidRedactIdentifiers", True)),
        )

    def needs_item_metadata(self) -> bool:
        return bool(
            self.folder_allowlist
            or self.location_blocklist
            or self.blocked_sensitivity_labels
        )

    def require_mailbox(self, mailbox: str | None) -> None:
        if not self.enabled:
            return
        normalized = _normalize(mailbox)
        if normalized and normalized in self.mailbox_blocklist:
            raise BlockedError("mailbox_blocklisted")
        if not self.mailbox_allowlist:
            raise BlockedError("mailbox_not_allowlisted")
        if not normalized or normalized not in self.mailbox_allowlist:
            raise BlockedError("mailbox_not_allowlisted")

    def require_unredactable_content(self) -> None:
        if self.enabled:
            raise BlockedError("unredactable_content")

    def require_location(self, location: Location) -> None:
        if not self.enabled:
            return
        if self._is_blocklisted(location):
            raise BlockedError("location_blocklisted")
        if self._has_blocked_label(location.labels):
            raise BlockedError("sensitivity_label")
        if not self._is_allowlisted(location):
            raise BlockedError("location_not_allowlisted")

    def allows_location(self, location: Location) -> bool:
        try:
            self.require_location(location)
        except BlockedError:
            return False
        return True

    def redact_text(self, value: str) -> str:
        if not self.enabled or not self.redact_identifiers or not value:
            return value
        redacted = SSN_RE.sub(REDACTED, value)
        return EIN_RE.sub(REDACTED, redacted)

    def redact_value(self, value: Any) -> Any:
        if isinstance(value, str):
            return self.redact_text(value)
        return value

    def redact_grid(self, grid: list[list[Any]] | None) -> list[list[Any]] | None:
        if grid is None:
            return None
        return [[self.redact_value(cell) for cell in row] for row in grid]

    def _is_blocklisted(self, location: Location) -> bool:
        if not self.location_blocklist:
            return False
        identities = [
            _normalize(location.site_id),
            _normalize(location.drive_id),
        ]
        pathish = " ".join(
            part
            for part in (_dir_path(location.path), _url_dir(location.web_url))
            if part
        )
        for entry in self.location_blocklist:
            if any(entry == ident or entry in ident for ident in identities if ident):
                return True
            if pathish and entry in pathish:
                return True
        if not pathish:
            return True
        return False

    def _is_allowlisted(self, location: Location) -> bool:
        has_any_allowlist = bool(
            self.site_allowlist or self.drive_allowlist or self.folder_allowlist
        )
        if not has_any_allowlist:
            return False
        matched = False
        if self.site_allowlist:
            site_values = [_normalize(location.site_id)]
            if _exact_or_id_match(site_values, self.site_allowlist):
                matched = True
            elif _folder_prefix_match(_url_dir(location.web_url), self.site_allowlist):
                matched = True
            elif _normalize(location.site_id) or _url_dir(location.web_url):
                return False
        if self.drive_allowlist:
            drive_values = [_normalize(location.drive_id)]
            if _exact_or_id_match(drive_values, self.drive_allowlist):
                matched = True
            elif _folder_prefix_match(_url_dir(location.web_url), self.drive_allowlist):
                matched = True
            elif _normalize(location.drive_id) or _url_dir(location.web_url):
                return False
        if self.folder_allowlist:
            pathish = _dir_path(location.path) or _url_dir(location.web_url)
            if not pathish:
                return False
            if not _folder_prefix_match(pathish, self.folder_allowlist):
                return False
            matched = True
        return matched

    def _has_blocked_label(self, labels: tuple[str, ...]) -> bool:
        if not self.blocked_sensitivity_labels or not labels:
            return False
        blocked = set(self.blocked_sensitivity_labels)
        return any(_normalize(label) in blocked for label in labels)


def _normalized_list(values: list[str] | tuple[str, ...] | None) -> list[str]:
    return [_normalize(value) for value in (values or []) if _normalize(value)]


def labels_from_graph(payload: dict[str, Any] | None) -> tuple[str, ...]:
    if not isinstance(payload, dict):
        return ()
    labels: list[str] = []

    def collect(raw: Any) -> None:
        if isinstance(raw, dict):
            for field_name in (
                "id",
                "displayName",
                "name",
                "sensitivityLabelId",
            ):
                value = raw.get(field_name)
                if isinstance(value, str) and value.strip():
                    labels.append(value)
            nested = raw.get("sensitivityLabel")
            if nested is not None and nested is not raw:
                collect(nested)
        elif isinstance(raw, list):
            for item in raw:
                collect(item)

    collect(payload.get("sensitivityLabel"))
    collect(payload.get("sensitivityLabelAssignment"))
    collect(payload.get("labels"))
    return tuple(labels)
