import { z } from "zod";

function parseBoolean(value: string | undefined, fallback: boolean): boolean {
  if (value === undefined) {
    return fallback;
  }

  return value === "true" || value === "1";
}

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
  PLATFORM_WORKSPACE_ENABLED: z.boolean().default(true),
  PLATFORM_HOME_REDIRECT: z.boolean().default(false),
  PLATFORM_DEFAULT_TENANT_SLUG: z.string().default("teacheractive"),
  PLATFORM_DEFAULT_ENVIRONMENT_SLUG: z.string().default("development"),
  PLATFORM_GIT_OUTPUT_DIR: z.string().default("platform-manifests"),
  PLATFORM_GIT_AUTO_COMMIT: z.boolean().default(false),
  PLATFORM_GIT_AUTHOR_NAME: z.string().default("TeacherActive Platform"),
  PLATFORM_GIT_AUTHOR_EMAIL: z.string().email().default("platform@local.test"),
  PLATFORM_LOCAL_DEV_MODE: z.boolean().default(process.env.NODE_ENV !== "production"),
  PLATFORM_SESSION_SECRET: z.string().optional(),
  PLATFORM_DEV_ACTOR_EMAIL: z.string().email().default("builder@teacheractive.local"),
  PLATFORM_DEV_ACTOR_NAME: z.string().default("Local Builder"),
  PLATFORM_DEV_ACTOR_ROLE: z.enum(["SUPER_ADMIN", "BUILDER_ADMIN", "USER"]).default("SUPER_ADMIN"),
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
    PLATFORM_WORKSPACE_ENABLED: parseBoolean(process.env.PLATFORM_WORKSPACE_ENABLED, true),
    PLATFORM_HOME_REDIRECT: parseBoolean(process.env.PLATFORM_HOME_REDIRECT, false),
    PLATFORM_DEFAULT_TENANT_SLUG: process.env.PLATFORM_DEFAULT_TENANT_SLUG,
    PLATFORM_DEFAULT_ENVIRONMENT_SLUG: process.env.PLATFORM_DEFAULT_ENVIRONMENT_SLUG,
    PLATFORM_GIT_OUTPUT_DIR: process.env.PLATFORM_GIT_OUTPUT_DIR,
    PLATFORM_GIT_AUTO_COMMIT: parseBoolean(process.env.PLATFORM_GIT_AUTO_COMMIT, false),
    PLATFORM_GIT_AUTHOR_NAME: process.env.PLATFORM_GIT_AUTHOR_NAME,
    PLATFORM_GIT_AUTHOR_EMAIL: process.env.PLATFORM_GIT_AUTHOR_EMAIL,
    PLATFORM_LOCAL_DEV_MODE: parseBoolean(process.env.PLATFORM_LOCAL_DEV_MODE, process.env.NODE_ENV !== "production"),
    PLATFORM_SESSION_SECRET: process.env.PLATFORM_SESSION_SECRET,
    PLATFORM_DEV_ACTOR_EMAIL: process.env.PLATFORM_DEV_ACTOR_EMAIL,
    PLATFORM_DEV_ACTOR_NAME: process.env.PLATFORM_DEV_ACTOR_NAME,
    PLATFORM_DEV_ACTOR_ROLE: process.env.PLATFORM_DEV_ACTOR_ROLE,
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

export function resetEnvCache(): void {
  cachedEnv = null;
}
