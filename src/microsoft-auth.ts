import { Buffer } from "node:buffer";
import { randomUUID } from "node:crypto";

import type { AppConfig } from "./config.js";
import { EncryptedFileStore } from "./token-store.js";

type MicrosoftIdClaims = {
  name?: string;
  preferred_username?: string;
  upn?: string;
  email?: string;
  oid?: string;
  tid?: string;
};

type StoredMicrosoftTokens = {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
  scope: string;
  idToken?: string;
  account?: {
    name?: string;
    preferredUsername?: string;
    oid?: string;
    tid?: string;
  };
  updatedAt: number;
};

type TokenResponse = {
  access_token: string;
  refresh_token?: string;
  expires_in?: number;
  scope?: string;
  id_token?: string;
  error?: string;
  error_description?: string;
};

export class MicrosoftAuthService {
  private readonly tokenStore: EncryptedFileStore<StoredMicrosoftTokens>;
  private pendingState: string | null = null;

  constructor(private readonly config: AppConfig) {
    this.tokenStore = new EncryptedFileStore<StoredMicrosoftTokens>(
      ".tokens/microsoft-graph-token.json",
      config.encryptionKey,
    );
  }

  buildAuthorizationUrl(): string {
    const state = randomUUID();
    this.pendingState = state;
    const authorizeUrl = new URL(
      `https://login.microsoftonline.com/${this.config.microsoft.tenantId}/oauth2/v2.0/authorize`,
    );

    authorizeUrl.searchParams.set("client_id", this.config.microsoft.clientId);
    authorizeUrl.searchParams.set("response_type", "code");
    authorizeUrl.searchParams.set("redirect_uri", this.config.microsoft.redirectUri);
    authorizeUrl.searchParams.set("response_mode", "query");
    authorizeUrl.searchParams.set("scope", this.config.microsoft.scopes.join(" "));
    authorizeUrl.searchParams.set("state", state);

    return authorizeUrl.toString();
  }

  async handleAuthorizationCodeCallback(params: {
    code?: string;
    state?: string;
    error?: string;
    errorDescription?: string;
  }): Promise<void> {
    if (params.error) {
      throw new Error(
        `Microsoft sign-in failed: ${params.errorDescription ?? params.error}`,
      );
    }

    if (!params.code || !params.state) {
      throw new Error("Microsoft callback is missing the code or state parameter");
    }

    if (!this.pendingState || params.state !== this.pendingState) {
      throw new Error("Microsoft callback state was invalid");
    }

    this.pendingState = null;

    const tokenResponse = await this.fetchToken({
      grant_type: "authorization_code",
      code: params.code,
      redirect_uri: this.config.microsoft.redirectUri,
      scope: this.config.microsoft.scopes.join(" "),
    });

    if (!tokenResponse.refresh_token) {
      throw new Error(
        "Microsoft did not return a refresh token. Make sure offline_access is granted.",
      );
    }

    await this.saveTokenResponse(tokenResponse, tokenResponse.refresh_token);
  }

  async getAccessToken(): Promise<string> {
    const tokens = await this.tokenStore.load();

    if (!tokens) {
      throw new Error(
        "Microsoft Graph is not connected yet. Visit /auth/microsoft/start first.",
      );
    }

    if (tokens.expiresAt > Date.now() + 60_000) {
      return tokens.accessToken;
    }

    const refreshed = await this.fetchToken({
      grant_type: "refresh_token",
      refresh_token: tokens.refreshToken,
      scope: this.config.microsoft.scopes.join(" "),
    });

    await this.saveTokenResponse(refreshed, refreshed.refresh_token ?? tokens.refreshToken);

    return refreshed.access_token;
  }

  async disconnect(): Promise<void> {
    await this.tokenStore.clear();
  }

  async getStatus(): Promise<{
    connected: boolean;
    account?: StoredMicrosoftTokens["account"];
    expiresAt?: number;
    knownMailboxes: string[];
  }> {
    const tokens = await this.tokenStore.load();

    if (!tokens) {
      return {
        connected: false,
        knownMailboxes: this.config.knownMailboxes,
      };
    }

    return {
      connected: true,
      account: tokens.account,
      expiresAt: tokens.expiresAt,
      knownMailboxes: this.config.knownMailboxes,
    };
  }

  private async fetchToken(
    payload: Record<string, string>,
  ): Promise<TokenResponse> {
    const tokenUrl = `https://login.microsoftonline.com/${this.config.microsoft.tenantId}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
      client_id: this.config.microsoft.clientId,
      client_secret: this.config.microsoft.clientSecret,
      ...payload,
    });

    const response = await fetch(tokenUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: body.toString(),
    });

    const json = (await response.json()) as TokenResponse;

    if (!response.ok || json.error) {
      throw new Error(
        `Microsoft token exchange failed: ${json.error_description ?? json.error ?? response.statusText}`,
      );
    }

    return json;
  }

  private async saveTokenResponse(
    tokenResponse: TokenResponse,
    refreshToken: string,
  ): Promise<void> {
    const claims = this.decodeIdClaims(tokenResponse.id_token);
    const expiresAt =
      Date.now() + (tokenResponse.expires_in ?? 3600) * 1000;

    await this.tokenStore.save({
      accessToken: tokenResponse.access_token,
      refreshToken,
      expiresAt,
      scope: tokenResponse.scope ?? this.config.microsoft.scopes.join(" "),
      idToken: tokenResponse.id_token,
      account: {
        name: claims?.name,
        preferredUsername:
          claims?.preferred_username ?? claims?.upn ?? claims?.email,
        oid: claims?.oid,
        tid: claims?.tid,
      },
      updatedAt: Date.now(),
    });
  }

  private decodeIdClaims(idToken: string | undefined): MicrosoftIdClaims | null {
    if (!idToken) {
      return null;
    }

    const [, payload] = idToken.split(".");
    if (!payload) {
      return null;
    }

    try {
      return JSON.parse(Buffer.from(payload, "base64url").toString("utf8")) as MicrosoftIdClaims;
    } catch {
      return null;
    }
  }
}
