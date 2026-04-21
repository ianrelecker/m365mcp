from __future__ import annotations

import httpx
import pytest

from m365_mcp.microsoft_auth import MicrosoftAuthService, decode_id_claims

from .conftest import make_jwt


def test_decode_id_claims_extracts_payload() -> None:
    token = make_jwt(
        {
            "name": "A User",
            "preferred_username": "user@example.com",
            "oid": "oid-123",
            "tid": "tid-456",
        }
    )

    claims = decode_id_claims(token)
    assert claims == {
        "name": "A User",
        "preferred_username": "user@example.com",
        "oid": "oid-123",
        "tid": "tid-456",
    }


@pytest.mark.anyio
async def test_callback_rejects_invalid_state(config_factory) -> None:
    client = httpx.AsyncClient(transport=httpx.MockTransport(lambda request: httpx.Response(500)))
    auth = MicrosoftAuthService(config_factory(), client)
    auth.build_authorization_url()

    with pytest.raises(RuntimeError, match="state was invalid"):
        await auth.handle_authorization_code_callback(
            code="auth-code",
            state="wrong-state",
            error=None,
            errorDescription=None,
        )

    await client.aclose()
