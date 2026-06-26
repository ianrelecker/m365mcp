from __future__ import annotations

import json
import os
import sys
import tempfile
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
        parent = self._file_path.parent
        parent.mkdir(parents=True, exist_ok=True)
        if sys.platform != "win32":
            # Restrict the directory so only the owner can list/enter it.
            # Windows relies on user-profile ACLs instead.
            try:
                os.chmod(parent, 0o700)
            except OSError:
                pass

        encrypted = encrypt_json(value, self._encryption_key)
        content = json.dumps(encrypted.model_dump(mode="json"), indent=2)

        # Atomic write: write to a sibling temp file then rename so the token
        # file is never left in a truncated state on a crash or interrupt.
        fd, tmp_path = tempfile.mkstemp(dir=str(parent), text=True)
        try:
            if sys.platform != "win32":
                os.chmod(tmp_path, 0o600)
            with os.fdopen(fd, "w", encoding="utf-8") as fh:
                fh.write(content)
            fd = -1  # fdopen took ownership; don't double-close
            os.replace(tmp_path, str(self._file_path))
        except Exception:
            if fd != -1:
                os.close(fd)
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
            raise

    async def clear(self) -> None:
        try:
            self._file_path.unlink()
        except FileNotFoundError:
            return
