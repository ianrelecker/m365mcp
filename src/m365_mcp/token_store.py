from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .crypto import decrypt_json, encrypt_json
from .models import EncryptedPayload


class EncryptedFileStore:
    def __init__(self, file_path: Path, encryption_key: bytes) -> None:
        self._file_path = file_path
        self._encryption_key = encryption_key

    async def load(self) -> dict[str, Any] | None:
        try:
            raw = self._file_path.read_text("utf-8")
        except FileNotFoundError:
            return None

        payload = EncryptedPayload.model_validate_json(raw)
        decrypted = decrypt_json(payload, self._encryption_key)
        if not isinstance(decrypted, dict):
            raise ValueError("Encrypted token payload must decode to an object")
        return decrypted

    async def save(self, value: Any) -> None:
        self._file_path.parent.mkdir(parents=True, exist_ok=True)
        encrypted = encrypt_json(value, self._encryption_key)
        self._file_path.write_text(
            json.dumps(encrypted.model_dump(mode="json"), indent=2),
            encoding="utf-8",
        )

    async def clear(self) -> None:
        try:
            self._file_path.unlink()
        except FileNotFoundError:
            return
