import { z } from "zod";

const envSchema = z.object({
  DATABASE_URL: z.string().optional(),
  STORAGE_MODE: z.enum(["local", "s3"]).default("local"),
  LOCAL_STORAGE_DIR: z.string().default(".storage"),
  S3_BUCKET: z.string().optional(),
  S3_REGION: z.string().default("eu-west-1"),
  S3_ENDPOINT: z.string().url().optional(),
  S3_ACCESS_KEY_ID: z.string().optional(),
  S3_SECRET_ACCESS_KEY: z.string().optional(),
  APP_BASE_URL: z.string().url().default("http://localhost:3000"),
  PLAYWRIGHT_BROWSER_PATH: z.string().optional(),
});

export type AppEnv = z.infer<typeof envSchema>;

let cachedEnv: AppEnv | null = null;

export function getEnv(): AppEnv {
  if (cachedEnv) {
    return cachedEnv;
  }

  cachedEnv = envSchema.parse({
    DATABASE_URL: process.env.DATABASE_URL,
    STORAGE_MODE: process.env.STORAGE_MODE,
    LOCAL_STORAGE_DIR: process.env.LOCAL_STORAGE_DIR,
    S3_BUCKET: process.env.S3_BUCKET,
    S3_REGION: process.env.S3_REGION,
    S3_ENDPOINT: process.env.S3_ENDPOINT,
    S3_ACCESS_KEY_ID: process.env.S3_ACCESS_KEY_ID,
    S3_SECRET_ACCESS_KEY: process.env.S3_SECRET_ACCESS_KEY,
    APP_BASE_URL: process.env.APP_BASE_URL,
    PLAYWRIGHT_BROWSER_PATH: process.env.PLAYWRIGHT_BROWSER_PATH,
  });

  return cachedEnv;
}

export function requireDatabaseUrl(): string {
  const env = getEnv();

  if (!env.DATABASE_URL) {
    throw new Error("DATABASE_URL is required for database-backed operations.");
  }

  return env.DATABASE_URL;
}
