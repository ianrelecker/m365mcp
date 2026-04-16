import {
  createCipheriv,
  createDecipheriv,
  randomBytes,
} from "node:crypto";

const IV_LENGTH = 12;

export type EncryptedPayload = {
  iv: string;
  tag: string;
  ciphertext: string;
};

export function encryptJson(value: unknown, encryptionKey: Buffer): EncryptedPayload {
  const iv = randomBytes(IV_LENGTH);
  const cipher = createCipheriv("aes-256-gcm", encryptionKey, iv);
  const plaintext = Buffer.from(JSON.stringify(value), "utf8");
  const ciphertext = Buffer.concat([cipher.update(plaintext), cipher.final()]);
  const tag = cipher.getAuthTag();

  return {
    iv: iv.toString("base64"),
    tag: tag.toString("base64"),
    ciphertext: ciphertext.toString("base64"),
  };
}

export function decryptJson<T>(payload: EncryptedPayload, encryptionKey: Buffer): T {
  const decipher = createDecipheriv(
    "aes-256-gcm",
    encryptionKey,
    Buffer.from(payload.iv, "base64"),
  );

  decipher.setAuthTag(Buffer.from(payload.tag, "base64"));

  const plaintext = Buffer.concat([
    decipher.update(Buffer.from(payload.ciphertext, "base64")),
    decipher.final(),
  ]);

  return JSON.parse(plaintext.toString("utf8")) as T;
}
