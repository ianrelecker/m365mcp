from __future__ import annotations

import json
from datetime import UTC, datetime
from pathlib import Path
from typing import Any


ID_FIELDS = {
    "attachmentId",
    "categoryId",
    "contactId",
    "conversationId",
    "destinationFolderId",
    "destinationFolderPath",
    "eventId",
    "folderId",
    "folderPath",
    "messageId",
    "parentFolderId",
    "parentFolderPath",
    "ruleId",
}
SENSITIVE_FIELDS = {
    "bcc",
    "body",
    "cc",
    "comment",
    "emailAddresses",
    "from",
    "from_",
    "mobilePhone",
    "query",
    "subject",
    "to",
}


def _utc_timestamp() -> str:
    return datetime.now(UTC).isoformat().replace("+00:00", "Z")


def classify_tool(tool_name: str) -> str:
    if tool_name == "auth_status":
        return "auth"
    if tool_name in {"mail_send", "mail_send_draft", "mail_send_reply"}:
        return "send"
    if "delete" in tool_name or tool_name in {"mail_clear_categories"}:
        return "delete"
    if "rule" in tool_name:
        return "rule"
    if any(
        marker in tool_name
        for marker in (
            "create",
            "update",
            "rename",
            "move",
            "set",
            "add",
            "remove",
            "mark",
        )
    ):
        return "write"
    return "read"


def audit_metadata(tool_name: str, arguments: dict[str, Any]) -> dict[str, Any]:
    metadata: dict[str, Any] = {
        "tool": tool_name,
        "category": classify_tool(tool_name),
    }

    if tool_name.startswith(("mail_", "contacts_", "calendar_")):
        metadata["mailbox"] = arguments.get("mailbox") or "me"
    elif "mailbox" in arguments:
        metadata["mailbox"] = arguments.get("mailbox") or "me"

    ids = {
        field: value
        for field, value in arguments.items()
        if field in ID_FIELDS and value is not None
    }
    if ids:
        metadata["ids"] = ids

    return metadata


def _sensitive_values(arguments: dict[str, Any]) -> list[str]:
    values: list[str] = []

    def collect(value: Any) -> None:
        if isinstance(value, str) and value:
            values.append(value)
        elif isinstance(value, list):
            for item in value:
                collect(item)
        elif isinstance(value, dict):
            for item in value.values():
                collect(item)

    for field, value in arguments.items():
        if field in SENSITIVE_FIELDS:
            collect(value)
    return values


def redact_error_message(message: str, arguments: dict[str, Any]) -> str:
    redacted = message
    for value in _sensitive_values(arguments):
        redacted = redacted.replace(value, "[redacted]")
    return redacted[:500]


class LocalAuditLogger:
    def __init__(self, *, enabled: bool, file_path: Path) -> None:
        self._enabled = enabled
        self._file_path = file_path

    async def record_tool_call(
        self,
        *,
        tool_name: str,
        arguments: dict[str, Any],
        outcome: str,
        error: BaseException | None = None,
    ) -> None:
        if not self._enabled:
            return

        record: dict[str, Any] = {
            "timestamp": _utc_timestamp(),
            "outcome": outcome,
            **audit_metadata(tool_name, arguments),
        }
        if error is not None:
            record["error"] = {
                "type": error.__class__.__name__,
                "message": redact_error_message(str(error), arguments),
            }

        self._file_path.parent.mkdir(parents=True, exist_ok=True)
        with self._file_path.open("a", encoding="utf-8") as audit_file:
            audit_file.write(json.dumps(record, sort_keys=True) + "\n")
