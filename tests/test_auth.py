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
    form_bodies: list[dict[str, list[str]]] = []

    def handler(request: httpx.Request) -> httpx.Response:
        form_bodies.append(parse_qs(request.content.decode("utf-8")))
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
    auth_url = auth.build_authorization_url()
    auth_params = parse_qs(urlparse(auth_url).query)
    state = auth_params["state"][0]

    assert auth_params["code_challenge_method"] == ["S256"]
    assert len(auth_params["code_challenge"][0]) >= 43

    await auth.handle_authorization_code_callback(
        code="auth-code",
        state=state,
        error=None,
        errorDescription=None,
    )

    assert "code_verifier" in form_bodies[0]
    assert len(form_bodies[0]["code_verifier"][0]) >= 43

    status = await auth.get_status()
    assert status.connected is True
    assert "Mail.ReadWrite.Shared" in status.grantedScopes
    assert "Contacts.ReadWrite.Shared" in status.missingScopes
    assert "MailboxSettings.ReadWrite" in status.missingScopes

    await client.aclose()


@pytest.mark.anyio
async def test_refresh_token_flow_does_not_send_code_verifier(config_factory) -> None:
    form_bodies: list[dict[str, list[str]]] = []

    def handler(request: httpx.Request) -> httpx.Response:
        form = parse_qs(request.content.decode("utf-8"))
        form_bodies.append(form)
        if form["grant_type"] == ["authorization_code"]:
            return httpx.Response(
                200,
                json={
                    "access_token": "initial-token",
                    "refresh_token": "refresh-token",
                    "expires_in": 0,
                    "scope": " ".join(config_factory().microsoft.scopes),
                    "id_token": make_jwt({"preferred_username": "user@example.com"}),
                },
            )

        return httpx.Response(
            200,
            json={
                "access_token": "refreshed-token",
                "expires_in": 3600,
                "scope": " ".join(config_factory().microsoft.scopes),
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

    token = await auth.get_access_token()

    assert token == "refreshed-token"
    assert form_bodies[0]["grant_type"] == ["authorization_code"]
    assert "code_verifier" in form_bodies[0]
    assert form_bodies[1]["grant_type"] == ["refresh_token"]
    assert "code_verifier" not in form_bodies[1]

    await client.aclose()
