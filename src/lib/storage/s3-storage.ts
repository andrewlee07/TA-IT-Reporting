import { GetObjectCommand, HeadObjectCommand, PutObjectCommand, S3Client } from "@aws-sdk/client-s3";

import { getEnv } from "@/lib/env";
import type { ObjectStorage, StoredObject } from "@/lib/storage/types";

async function streamToBuffer(stream: unknown): Promise<Buffer> {
  if (!(stream instanceof ReadableStream)) {
    const asyncIterable = stream as AsyncIterable<Uint8Array>;
    const chunks: Uint8Array[] = [];

    for await (const chunk of asyncIterable) {
      chunks.push(chunk);
    }

    return Buffer.concat(chunks);
  }

  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];

  while (true) {
    const { done, value } = await reader.read();

    if (done) {
      break;
    }

    if (value) {
      chunks.push(value);
    }
  }

  return Buffer.concat(chunks);
}

export class S3ObjectStorage implements ObjectStorage {
  private readonly client: S3Client;
  private readonly bucket: string;

  constructor() {
    const env = getEnv();

    if (!env.S3_BUCKET || !env.S3_ACCESS_KEY_ID || !env.S3_SECRET_ACCESS_KEY) {
      throw new Error("S3 storage is configured but bucket or credentials are missing.");
    }

    this.bucket = env.S3_BUCKET;
    this.client = new S3Client({
      region: env.S3_REGION,
      endpoint: env.S3_ENDPOINT,
      forcePathStyle: Boolean(env.S3_ENDPOINT),
      credentials: {
        accessKeyId: env.S3_ACCESS_KEY_ID,
        secretAccessKey: env.S3_SECRET_ACCESS_KEY,
      },
    });
  }

  async putBuffer(key: string, buffer: Buffer, contentType: string): Promise<StoredObject> {
    await this.client.send(
      new PutObjectCommand({
        Bucket: this.bucket,
        Key: key,
        Body: buffer,
        ContentType: contentType,
      }),
    );

    return {
      key,
      contentType,
      size: buffer.byteLength,
    };
  }

  async getBuffer(key: string): Promise<Buffer> {
    const response = await this.client.send(
      new GetObjectCommand({
        Bucket: this.bucket,
        Key: key,
      }),
    );

    return streamToBuffer(response.Body);
  }

  async exists(key: string): Promise<boolean> {
    try {
      await this.client.send(
        new HeadObjectCommand({
          Bucket: this.bucket,
          Key: key,
        }),
      );

      return true;
    } catch {
      return false;
    }
  }
}
