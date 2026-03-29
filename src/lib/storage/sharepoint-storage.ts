import { SharePointGraphClient } from "@/lib/sharepoint/graph-client";
import type { ObjectStorage, StoredObject } from "@/lib/storage/types";

export class SharePointObjectStorage implements ObjectStorage {
  private readonly client = new SharePointGraphClient();

  async putBuffer(key: string, buffer: Buffer, contentType: string): Promise<StoredObject> {
    await this.client.putBuffer(key, buffer, contentType);

    return {
      key,
      contentType,
      size: buffer.byteLength,
    };
  }

  async getBuffer(key: string): Promise<Buffer> {
    return this.client.getBuffer(key);
  }

  async exists(key: string): Promise<boolean> {
    return this.client.exists(key);
  }
}
