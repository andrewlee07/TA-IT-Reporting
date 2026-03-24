import { createHmac, timingSafeEqual } from "node:crypto";

import { getEnv } from "@/lib/env";
import { PlatformUnauthorizedError } from "@/lib/platform/errors";
import type { PlatformRole, PlatformViewAsState } from "@/lib/platform/types";

export const PLATFORM_SESSION_COOKIE = "ta_platform_session";
export const PLATFORM_VIEW_AS_COOKIE = "ta_platform_view_as";
export const PLATFORM_SESSION_MAX_AGE_SECONDS = 60 * 60 * 24 * 14;

interface PlatformSessionPayload {
  email: string;
  name: string;
  issuedAt: string;
}

interface PlatformViewAsPayload {
  role: PlatformRole;
  personaLabel: string;
  actorEmail?: string;
  issuedAt: string;
}

export interface PlatformActorIdentity {
  email: string;
  name: string;
  source: "session" | "local_dev";
}

function encodeBase64Url(value: string): string {
  return Buffer.from(value, "utf8").toString("base64url");
}

function decodeBase64Url(value: string): string {
  return Buffer.from(value, "base64url").toString("utf8");
}

function getSessionSecret(): string {
  const env = getEnv();
  if (env.PLATFORM_SESSION_SECRET) {
    return env.PLATFORM_SESSION_SECRET;
  }

  if (env.PLATFORM_LOCAL_DEV_MODE) {
    return "teacheractive-platform-local-dev-secret";
  }

  throw new PlatformUnauthorizedError("PLATFORM_SESSION_SECRET is required to validate platform sessions.");
}

function signPayload(encodedPayload: string): string {
  return createHmac("sha256", getSessionSecret()).update(encodedPayload).digest("base64url");
}

function parseCookies(cookieHeader: string | null | undefined): Record<string, string> {
  if (!cookieHeader) {
    return {};
  }

  return cookieHeader
    .split(";")
    .map((cookie) => cookie.trim())
    .filter(Boolean)
    .reduce<Record<string, string>>((accumulator, cookie) => {
      const separatorIndex = cookie.indexOf("=");
      if (separatorIndex === -1) {
        return accumulator;
      }

      const key = cookie.slice(0, separatorIndex).trim();
      const value = cookie.slice(separatorIndex + 1).trim();
      accumulator[key] = decodeURIComponent(value);
      return accumulator;
    }, {});
}

function parseSignedSession(rawValue: string): PlatformSessionPayload {
  const separatorIndex = rawValue.lastIndexOf(".");
  if (separatorIndex <= 0) {
    throw new PlatformUnauthorizedError("Platform session cookie is invalid.");
  }

  const encodedPayload = rawValue.slice(0, separatorIndex);
  const signature = rawValue.slice(separatorIndex + 1);
  const expectedSignature = signPayload(encodedPayload);

  const providedBuffer = Buffer.from(signature);
  const expectedBuffer = Buffer.from(expectedSignature);
  if (providedBuffer.length !== expectedBuffer.length || !timingSafeEqual(providedBuffer, expectedBuffer)) {
    throw new PlatformUnauthorizedError("Platform session signature is invalid.");
  }

  const payload = JSON.parse(decodeBase64Url(encodedPayload)) as Partial<PlatformSessionPayload>;
  if (!payload.email || !payload.name) {
    throw new PlatformUnauthorizedError("Platform session payload is incomplete.");
  }

  return {
    email: payload.email,
    name: payload.name,
    issuedAt: payload.issuedAt ?? new Date().toISOString(),
  };
}

function getCookieHeader(request?: Request | Headers): string | null | undefined {
  if (!request) {
    return undefined;
  }

  return request instanceof Headers ? request.get("cookie") : request.headers.get("cookie");
}

export function resolvePlatformActorIdentity(request?: Request | Headers): PlatformActorIdentity {
  const env = getEnv();
  const cookies = parseCookies(getCookieHeader(request));
  const sessionCookie = cookies[PLATFORM_SESSION_COOKIE];

  if (sessionCookie) {
    const payload = parseSignedSession(sessionCookie);
    return {
      email: payload.email,
      name: payload.name,
      source: "session",
    };
  }

  if (env.PLATFORM_LOCAL_DEV_MODE) {
    return {
      email: env.PLATFORM_DEV_ACTOR_EMAIL,
      name: env.PLATFORM_DEV_ACTOR_NAME,
      source: "local_dev",
    };
  }

  throw new PlatformUnauthorizedError();
}

export function tryResolvePlatformActorIdentity(request?: Request | Headers): PlatformActorIdentity | null {
  try {
    return resolvePlatformActorIdentity(request);
  } catch {
    return null;
  }
}

function parseSignedViewAs(rawValue: string): PlatformViewAsPayload {
  const separatorIndex = rawValue.lastIndexOf(".");
  if (separatorIndex <= 0) {
    throw new PlatformUnauthorizedError("Platform view-as cookie is invalid.");
  }

  const encodedPayload = rawValue.slice(0, separatorIndex);
  const signature = rawValue.slice(separatorIndex + 1);
  const expectedSignature = signPayload(encodedPayload);
  const providedBuffer = Buffer.from(signature);
  const expectedBuffer = Buffer.from(expectedSignature);

  if (providedBuffer.length !== expectedBuffer.length || !timingSafeEqual(providedBuffer, expectedBuffer)) {
    throw new PlatformUnauthorizedError("Platform view-as signature is invalid.");
  }

  const payload = JSON.parse(decodeBase64Url(encodedPayload)) as Partial<PlatformViewAsPayload>;
  if (!payload.role || !payload.personaLabel) {
    throw new PlatformUnauthorizedError("Platform view-as payload is incomplete.");
  }

  return {
    role: payload.role,
    personaLabel: payload.personaLabel,
    actorEmail: payload.actorEmail,
    issuedAt: payload.issuedAt ?? new Date().toISOString(),
  };
}

export function createSignedPlatformSessionValue(input: { email: string; name: string }): string {
  const payload: PlatformSessionPayload = {
    email: input.email,
    name: input.name,
    issuedAt: new Date().toISOString(),
  };
  const encodedPayload = encodeBase64Url(JSON.stringify(payload));
  return `${encodedPayload}.${signPayload(encodedPayload)}`;
}

export function createSignedPlatformViewAsValue(input: {
  role: PlatformRole;
  personaLabel: string;
  actorEmail?: string;
}): string {
  const payload: PlatformViewAsPayload = {
    role: input.role,
    personaLabel: input.personaLabel,
    actorEmail: input.actorEmail,
    issuedAt: new Date().toISOString(),
  };
  const encodedPayload = encodeBase64Url(JSON.stringify(payload));
  return `${encodedPayload}.${signPayload(encodedPayload)}`;
}

export function tryResolvePlatformViewAsState(request?: Request | Headers): PlatformViewAsState | null {
  try {
    const cookies = parseCookies(getCookieHeader(request));
    const viewAsCookie = cookies[PLATFORM_VIEW_AS_COOKIE];
    if (!viewAsCookie) {
      return null;
    }

    const payload = parseSignedViewAs(viewAsCookie);
    return {
      active: true,
      role: payload.role,
      personaLabel: payload.personaLabel,
      actorEmail: payload.actorEmail,
    };
  } catch {
    return null;
  }
}

export function getPlatformSessionCookieSettings(): {
  httpOnly: boolean;
  sameSite: "lax";
  secure: boolean;
  path: string;
  maxAge: number;
} {
  return {
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
    path: "/",
    maxAge: PLATFORM_SESSION_MAX_AGE_SECONDS,
  };
}
