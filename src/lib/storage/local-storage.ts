import { promises as fs } from "node:fs";
import path from "node:path";

import type { ObjectStorage, StoredObject } from "@/lib/storage/types";

export class LocalObjectStorage implements ObjectStorage {
  private readonly rootDir: string;

  constructor(rootDir = path.join(process.cwd(), ".storage")) {
    this.rootDir = rootDir;
  }

  async putBuffer(key: string, buffer: Buffer, contentType: string): Promise<StoredObject> {
    const filePath = this.toFilePath(key);

    await fs.mkdir(path.dirname(filePath), { recursive: true });
    await fs.writeFile(filePath, buffer);

    return {
      key,
      contentType,
      size: buffer.byteLength,
    };
  }

  async getBuffer(key: string): Promise<Buffer> {
    return fs.readFile(this.toFilePath(key));
  }

  async exists(key: string): Promise<boolean> {
    try {
      await fs.access(this.toFilePath(key));
      return true;
    } catch {
      return false;
    }
  }

  private toFilePath(key: string): string {
    return path.join(this.rootDir, key);
  }
}
