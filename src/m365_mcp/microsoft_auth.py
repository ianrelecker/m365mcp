from __future__ import annotations

import base64
import json
import time
from contextlib import asynccontextmanager
from typing import Any
from uuid import uuid4

import httpx

from .config import AppConfig
from .models import AccountInfo, MicrosoftConnectionStatus, StoredMicrosoftTokens
from .token_store import EncryptedFileStore


def decode_id_claims(id_token: str | None) -> dict[str, Any] | None:
    if not id_token:
        return None

    parts = id_token.split(".")
    if len(parts) < 2:
        return None

    payload = parts[1]
    padding = "=" * (-len(payload) % 4)
    try:
        decoded = base64.urlsafe_b64decode(payload + padding)
        claims = json.loads(decoded.decode("utf-8"))
    except Exception:
        return None

    if not isinstance(claims, dict):
        return None
    return claims


class MicrosoftAuthService:
    def __init__(
        self,
        config: AppConfig,
        http_client: httpx.AsyncClient | None = None,
    ) -> None:
        self._config = config
        self._token_store = EncryptedFileStore(config.tokenFile, config.encryptionKey)
        self._http_client = http_client
        self._pending_state: str | None = None

    @asynccontextmanager
    async def _client(self) -> Any:
        if self._http_client is not None:
            yield self._http_client
            return

        async with httpx.AsyncClient(follow_redirects=False, timeout=30.0) as client:
            yield client

    def build_authorization_url(self) -> str:
        state = str(uuid4())
        self._pending_state = state

        request = httpx.URL(
            f"https://login.microsoftonline.com/{self._config.microsoft.tenantId}/oauth2/v2.0/authorize"
        )
        return str(
            request.copy_add_param("client_id", self._config.microsoft.clientId)
            .copy_add_param("response_type", "code")
            .copy_add_param("redirect_uri", self._config.microsoft.redirectUri)
            .copy_add_param("response_mode", "query")
            .copy_add_param("scope", " ".join(self._config.microsoft.scopes))
            .copy_add_param("state", state)
        )

    async def handle_authorization_code_callback(
        self,
        *,
        code: str | None,
        state: str | None,
        error: str | None,
        errorDescription: str | None,
    ) -> None:
        if error:
            raise RuntimeError(
                f"Microsoft sign-in failed: {errorDescription or error}"
            )

        if not code or not state:
            raise RuntimeError("Microsoft callback is missing the code or state parameter")

        if not self._pending_state or state != self._pending_state:
            raise RuntimeError("Microsoft callback state was invalid")

        self._pending_state = None

        token_response = await self._fetch_token(
            {
                "grant_type": "authorization_code",
                "code": code,
                "redirect_uri": self._config.microsoft.redirectUri,
                "scope": " ".join(self._config.microsoft.scopes),
            }
        )

        refresh_token = token_response.get("refresh_token")
        if not refresh_token:
            raise RuntimeError(
                "Microsoft did not return a refresh token. Make sure offline_access is granted."
            )

        await self._save_token_response(token_response, refresh_token)

    async def get_access_token(self) -> str:
        tokens = await self._load_tokens()
        if tokens is None:
            raise RuntimeError(
                "Microsoft Graph is not connected yet. Visit /auth/microsoft/start first."
            )

        if tokens.expiresAt > int(time.time() * 1000) + 60_000:
            return tokens.accessToken

        refreshed = await self._fetch_token(
            {
                "grant_type": "refresh_token",
                "refresh_token": tokens.refreshToken,
                "scope": " ".join(self._config.microsoft.scopes),
            }
        )

        await self._save_token_response(
            refreshed,
            refreshed.get("refresh_token") or tokens.refreshToken,
        )

        return str(refreshed["access_token"])

    async def disconnect(self) -> None:
        await self._token_store.clear()

    async def get_status(self) -> MicrosoftConnectionStatus:
        required_scopes = self._config.microsoft.scopes
        tokens = await self._load_tokens()
        if tokens is None:
            return MicrosoftConnectionStatus(
                connected=False,
                knownMailboxes=self._config.knownMailboxes,
                requiredScopes=required_scopes,
                grantedScopes=[],
                missingScopes=required_scopes,
            )

        granted_scopes = self._parse_scope_string(tokens.scope)
        return MicrosoftConnectionStatus(
            connected=True,
            account=tokens.account,
            expiresAt=tokens.expiresAt,
            knownMailboxes=self._config.knownMailboxes,
            requiredScopes=required_scopes,
            grantedScopes=granted_scopes,
            missingScopes=self._missing_scopes(required_scopes, granted_scopes),
        )

    async def _load_tokens(self) -> StoredMicrosoftTokens | None:
        raw = await self._token_store.load()
        return StoredMicrosoftTokens.model_validate(raw) if raw else None

    async def _fetch_token(self, payload: dict[str, str]) -> dict[str, Any]:
        token_url = (
            f"https://login.microsoftonline.com/"
            f"{self._config.microsoft.tenantId}/oauth2/v2.0/token"
        )
        body = {
            "client_id": self._config.microsoft.clientId,
            "client_secret": self._config.microsoft.clientSecret,
            **payload,
        }

        async with self._client() as client:
            response = await client.post(
                token_url,
                data=body,
                headers={"Content-Type": "application/x-www-form-urlencoded"},
            )

        data = response.json()
        if not isinstance(data, dict):
            raise RuntimeError(
                f"Microsoft token exchange failed: {response.reason_phrase}"
            )

        if not response.is_success or data.get("error"):
            raise RuntimeError(
                "Microsoft token exchange failed: "
                f"{data.get('error_description') or data.get('error') or response.reason_phrase}"
            )

        return data

    def _parse_scope_string(self, value: str) -> list[str]:
        return [scope for scope in value.split() if scope]

    def _missing_scopes(
        self,
        required_scopes: list[str],
        granted_scopes: list[str],
    ) -> list[str]:
        granted = {scope.lower() for scope in granted_scopes}
        return [scope for scope in required_scopes if scope.lower() not in granted]

    async def _save_token_response(
        self,
        token_response: dict[str, Any],
        refresh_token: str,
    ) -> None:
        claims = decode_id_claims(
            token_response.get("id_token")
            if isinstance(token_response.get("id_token"), str)
            else None
        )
        expires_in = int(token_response.get("expires_in", 3600))
        expires_at = int(time.time() * 1000) + expires_in * 1000

        account = AccountInfo(
            name=claims.get("name") if claims else None,
            preferredUsername=(
                claims.get("preferred_username")
                or claims.get("upn")
                or claims.get("email")
            )
            if claims
            else None,
            oid=claims.get("oid") if claims else None,
            tid=claims.get("tid") if claims else None,
        )

        tokens = StoredMicrosoftTokens(
            accessToken=str(token_response["access_token"]),
            refreshToken=refresh_token,
            expiresAt=expires_at,
            scope=str(token_response.get("scope") or " ".join(self._config.microsoft.scopes)),
            idToken=(
                str(token_response["id_token"])
                if token_response.get("id_token") is not None
                else None
            ),
            account=account,
            updatedAt=int(time.time() * 1000),
        )

        await self._token_store.save(tokens.model_dump(mode="json", by_alias=True))
