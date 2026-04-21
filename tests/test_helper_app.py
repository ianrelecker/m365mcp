from __future__ import annotations

from urllib.parse import parse_qs, urlparse

import httpx
import pytest

from m365_mcp.helper_app import create_helper_app
from m365_mcp.microsoft_auth import MicrosoftAuthService

from .conftest import make_jwt


@pytest.mark.anyio
async def test_helper_health_route_reports_connection_status(config_factory) -> None:
    auth_client = httpx.AsyncClient(transport=httpx.MockTransport(lambda request: httpx.Response(500)))
    auth = MicrosoftAuthService(config_factory(), auth_client)
    app = create_helper_app(config_factory(), auth)

    async with httpx.AsyncClient(
        transport=httpx.ASGITransport(app=app),
        base_url="http://testserver",
    ) as client:
        response = await client.get("/health")

    assert response.status_code == 200
    assert response.json()["microsoftConnected"] is False
    await auth_client.aclose()


@pytest.mark.anyio
async def test_helper_auth_start_callback_and_disconnect_flow(config_factory) -> None:
    config = config_factory()

    def handler(request: httpx.Request) -> httpx.Response:
        assert request.url.path.endswith("/oauth2/v2.0/token")
        return httpx.Response(
            200,
            json={
                "access_token": "access-token",
                "refresh_token": "refresh-token",
                "expires_in": 3600,
                "scope": "openid profile email offline_access",
                "id_token": make_jwt(
                    {
                        "name": "A User",
                        "preferred_username": "user@example.com",
                        "oid": "oid-123",
                        "tid": "tid-456",
                    }
                ),
            },
        )

    auth_client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    auth = MicrosoftAuthService(config, auth_client)
    app = create_helper_app(config, auth)

    async with httpx.AsyncClient(
        transport=httpx.ASGITransport(app=app),
        base_url="http://testserver",
        follow_redirects=False,
    ) as client:
        start = await client.get("/auth/microsoft/start")
        assert start.status_code == 302

        parsed = urlparse(start.headers["location"])
        state = parse_qs(parsed.query)["state"][0]

        callback = await client.get(
            "/auth/microsoft/callback",
            params={"code": "auth-code", "state": state},
        )
        assert callback.status_code == 200
        assert "Microsoft 365 connected" in callback.text

        health = await client.get("/health")
        assert health.json()["microsoftConnected"] is True

        disconnect = await client.get("/auth/microsoft/disconnect")
        assert disconnect.status_code == 302

        final_health = await client.get("/health")
        assert final_health.json()["microsoftConnected"] is False

    await auth_client.aclose()
