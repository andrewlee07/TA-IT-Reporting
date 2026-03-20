import { getEnv } from "@/lib/env";
import { LocalObjectStorage } from "@/lib/storage/local-storage";
import { S3ObjectStorage } from "@/lib/storage/s3-storage";
import type { ObjectStorage } from "@/lib/storage/types";

let cachedStorage: ObjectStorage | null = null;

export function getObjectStorage(): ObjectStorage {
  if (cachedStorage) {
    return cachedStorage;
  }

  cachedStorage = getEnv().STORAGE_MODE === "s3" ? new S3ObjectStorage() : new LocalObjectStorage();
  return cachedStorage;
}
