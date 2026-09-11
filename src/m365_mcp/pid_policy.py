from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

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


def _contains_any(haystack: str, needles: list[str]) -> bool:
    if not haystack or not needles:
        return False
    return any(needle in haystack for needle in needles)


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

    def require_mailbox(self, mailbox: str | None) -> None:
        if not self.enabled:
            return
        normalized = _normalize(mailbox)
        if not normalized:
            return
        if normalized in self.mailbox_blocklist:
            raise BlockedError("mailbox_blocklisted")
        if self.mailbox_allowlist and normalized not in self.mailbox_allowlist:
            raise BlockedError("mailbox_not_allowlisted")

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
        haystack = " ".join(
            part
            for part in (
                _normalize(location.site_id),
                _normalize(location.drive_id),
                _normalize(location.item_id),
                _normalize(location.path),
                _normalize(location.web_url),
            )
            if part
        )
        return _contains_any(haystack, self.location_blocklist)

    def _is_allowlisted(self, location: Location) -> bool:
        has_any_allowlist = bool(
            self.site_allowlist or self.drive_allowlist or self.folder_allowlist
        )
        if not has_any_allowlist:
            return False
        matched = False
        site_ok, site_matched = _allowlist_category(
            [_normalize(location.site_id), _normalize(location.web_url)],
            self.site_allowlist,
        )
        if not site_ok:
            return False
        matched = matched or site_matched
        drive_ok, drive_matched = _allowlist_category(
            [_normalize(location.drive_id), _normalize(location.web_url)],
            self.drive_allowlist,
        )
        if not drive_ok:
            return False
        matched = matched or drive_matched
        folder_ok, folder_matched = _allowlist_category(
            [_normalize(location.path), _normalize(location.web_url)],
            self.folder_allowlist,
        )
        if not folder_ok:
            return False
        matched = matched or folder_matched
        return matched

    def _has_blocked_label(self, labels: tuple[str, ...]) -> bool:
        if not self.blocked_sensitivity_labels or not labels:
            return False
        blocked = set(self.blocked_sensitivity_labels)
        return any(_normalize(label) in blocked for label in labels)


def _normalized_list(values: list[str] | tuple[str, ...] | None) -> list[str]:
    return [_normalize(value) for value in (values or []) if _normalize(value)]


def _allowlist_category(
    candidates: list[str], allowlist: list[str]
) -> tuple[bool, bool]:
    if not allowlist:
        return True, False
    haystack = " ".join(part for part in candidates if part)
    if not haystack:
        return True, False
    if _contains_any(haystack, allowlist):
        return True, True
    return False, False


def labels_from_graph(payload: dict[str, Any] | None) -> tuple[str, ...]:
    if not isinstance(payload, dict):
        return ()
    labels: list[str] = []
    for key in ("sensitivityLabel", "sensitivityLabelAssignment"):
        raw = payload.get(key)
        if isinstance(raw, dict):
            for field_name in ("id", "displayName", "name"):
                value = raw.get(field_name)
                if isinstance(value, str) and value.strip():
                    labels.append(value)
            nested = raw.get("sensitivityLabel")
            if isinstance(nested, dict):
                for field_name in ("id", "displayName", "name"):
                    value = nested.get(field_name)
                    if isinstance(value, str) and value.strip():
                        labels.append(value)
    return tuple(labels)
