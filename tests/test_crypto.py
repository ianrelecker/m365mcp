from __future__ import annotations

from m365_mcp.crypto import decrypt_json, encrypt_json
from m365_mcp.models import EncryptedPayload

from .conftest import TEST_KEY


def test_encrypt_json_round_trip() -> None:
    value = {"hello": "world", "count": 3}
    payload = encrypt_json(value, TEST_KEY)
    assert decrypt_json(payload, TEST_KEY) == value


def test_decrypts_existing_aes_gcm_payload() -> None:
    payload = EncryptedPayload(
        iv="AAECAwQFBgcICQoL",
        tag="QZG7Dacsjb2NoqsIN4Pj6Q==",
        ciphertext=(
            "PCC3eKaAsWjZLvzu38tCT+K15FGDCHIIVwyA6z9FIsBkdtyZ3KlG9x/BEc+ypVpdiCsF/jL717VU8"
            "kQ7NMGQloBVtB+gkFJDJmWdX928Ot0OvfFbRlJ/So6e8bwd0oOvpIg8iM0l6Grsc1jbJBABGtcGC8f"
            "B6QfvPZs2ueeCL3qZaDHxheX4Zz3XeJEN6ny/"
        ),
    )

    assert decrypt_json(payload, TEST_KEY) == {
        "accessToken": "access-token",
        "refreshToken": "refresh-token",
        "expiresAt": 1712345678901,
        "scope": "openid profile email",
        "updatedAt": 1712345600000,
    }
