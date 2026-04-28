from __future__ import annotations

import httpx
import pytest
from urllib.parse import parse_qs, urlparse

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


@pytest.mark.anyio
async def test_status_reports_missing_scopes(config_factory) -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(
            200,
            json={
                "access_token": "access-token",
                "refresh_token": "refresh-token",
                "expires_in": 3600,
                "scope": "openid profile email offline_access Mail.ReadWrite.Shared",
                "id_token": make_jwt({"preferred_username": "user@example.com"}),
            },
        )

    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    auth = MicrosoftAuthService(config_factory(), client)
    state = parse_qs(urlparse(auth.build_authorization_url()).query)["state"][0]

    await auth.handle_authorization_code_callback(
        code="auth-code",
        state=state,
        error=None,
        errorDescription=None,
    )

    status = await auth.get_status()
    assert status.connected is True
    assert "Mail.ReadWrite.Shared" in status.grantedScopes
    assert "Contacts.ReadWrite.Shared" in status.missingScopes
    assert "MailboxSettings.ReadWrite" in status.missingScopes

    await client.aclose()
