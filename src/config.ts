import dotenv from "dotenv";

dotenv.config();

function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }

  return value;
}

function optionalCommaList(value: string | undefined): string[] {
  if (!value) {
    return [];
  }

  return value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function parseEncryptionKey(name: string): Buffer {
  const value = requireEnv(name);
  const key = Buffer.from(value, "base64");

  if (key.length !== 32) {
    throw new Error(`${name} must be a base64-encoded 32-byte key`);
  }

  return key;
}

function parseUrl(name: string): URL {
  return new URL(requireEnv(name));
}

const port = Number.parseInt(process.env.PORT ?? "8787", 10);
const localBaseUrl = process.env.LOCAL_BASE_URL
  ? parseUrl("LOCAL_BASE_URL")
  : new URL(`http://127.0.0.1:${port}`);

export const config = {
  port,
  localBaseUrl,
  microsoft: {
    tenantId: requireEnv("MICROSOFT_TENANT_ID"),
    clientId: requireEnv("MICROSOFT_CLIENT_ID"),
    clientSecret: requireEnv("MICROSOFT_CLIENT_SECRET"),
    redirectUri: new URL("/auth/microsoft/callback", localBaseUrl).toString(),
    scopes: [
      "openid",
      "profile",
      "email",
      "offline_access",
      "Mail.ReadWrite.Shared",
      "Mail.Send.Shared",
      "Calendars.ReadWrite.Shared",
    ],
  },
  encryptionKey: parseEncryptionKey("TOKEN_ENCRYPTION_KEY"),
  knownMailboxes: optionalCommaList(process.env.KNOWN_MAILBOXES),
};

export type AppConfig = typeof config;
