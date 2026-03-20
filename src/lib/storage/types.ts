export interface StoredObject {
  key: string;
  contentType: string;
  size: number;
}

export interface ObjectStorage {
  putBuffer(key: string, buffer: Buffer, contentType: string): Promise<StoredObject>;
  getBuffer(key: string): Promise<Buffer>;
  exists(key: string): Promise<boolean>;
}
