import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";

import { decryptJson, encryptJson, type EncryptedPayload } from "./crypto.js";

export class EncryptedFileStore<T> {
  constructor(
    private readonly filePath: string,
    private readonly encryptionKey: Buffer,
  ) {}

  async load(): Promise<T | null> {
    try {
      const raw = await readFile(this.filePath, "utf8");
      const payload = JSON.parse(raw) as EncryptedPayload;

      return decryptJson<T>(payload, this.encryptionKey);
    } catch (error) {
      const nodeError = error as NodeJS.ErrnoException;
      if (nodeError.code === "ENOENT") {
        return null;
      }

      throw error;
    }
  }

  async save(value: T): Promise<void> {
    await mkdir(path.dirname(this.filePath), { recursive: true });
    const encrypted = encryptJson(value, this.encryptionKey);
    await writeFile(this.filePath, JSON.stringify(encrypted, null, 2), "utf8");
  }

  async clear(): Promise<void> {
    try {
      await rm(this.filePath);
    } catch (error) {
      const nodeError = error as NodeJS.ErrnoException;
      if (nodeError.code !== "ENOENT") {
        throw error;
      }
    }
  }
}
