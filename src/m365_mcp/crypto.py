from __future__ import annotations

import base64
import json
import os
from typing import Any

from cryptography.hazmat.primitives.ciphers.aead import AESGCM

from .models import EncryptedPayload

IV_LENGTH = 12


def encrypt_json(value: Any, encryption_key: bytes) -> EncryptedPayload:
    iv = os.urandom(IV_LENGTH)
    plaintext = json.dumps(value).encode("utf-8")
    encrypted = AESGCM(encryption_key).encrypt(iv, plaintext, associated_data=None)

    return EncryptedPayload(
        iv=base64.b64encode(iv).decode("ascii"),
        tag=base64.b64encode(encrypted[-16:]).decode("ascii"),
        ciphertext=base64.b64encode(encrypted[:-16]).decode("ascii"),
    )


def decrypt_json(payload: EncryptedPayload, encryption_key: bytes) -> Any:
    iv = base64.b64decode(payload.iv)
    tag = base64.b64decode(payload.tag)
    ciphertext = base64.b64decode(payload.ciphertext)
    plaintext = AESGCM(encryption_key).decrypt(iv, ciphertext + tag, associated_data=None)
    return json.loads(plaintext.decode("utf-8"))

