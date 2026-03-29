import { headers } from "next/headers";

import { getEnv } from "@/lib/env";

export interface AppUser {
  id: string;
  name: string;
  email: string | null;
}

function decodeClientPrincipal(rawValue: string): Record<string, unknown> | null {
  try {
    const json = Buffer.from(rawValue, "base64").toString("utf8");
    return JSON.parse(json) as Record<string, unknown>;
  } catch {
    return null;
  }
}

function getStringClaim(record: Record<string, unknown> | null, key: string): string | null {
  const value = record?.[key];
  return typeof value === "string" && value.length > 0 ? value : null;
}

export async function getAppUser(): Promise<AppUser | null> {
  const env = getEnv();

  if (env.AUTH_MODE === "development") {
    return {
      id: env.DEV_USER_ID,
      name: env.DEV_USER_NAME,
      email: null,
    };
  }

  const requestHeaders = await headers();
  const principalPayload = requestHeaders.get("x-ms-client-principal");
  const principal = principalPayload ? decodeClientPrincipal(principalPayload) : null;
  const identityProvider = requestHeaders.get("x-ms-client-principal-idp");

  if (!identityProvider && !principalPayload) {
    return null;
  }

  const id =
    requestHeaders.get("x-ms-client-principal-id") ??
    getStringClaim(principal, "userId") ??
    getStringClaim(principal, "sub");

  const email =
    requestHeaders.get("x-ms-client-principal-name") ??
    getStringClaim(principal, "userDetails") ??
    getStringClaim(principal, "preferred_username");

  const name =
    getStringClaim(principal, "name") ??
    email ??
    requestHeaders.get("x-ms-client-principal-name") ??
    "Authenticated User";

  if (!id) {
    return null;
  }

  return {
    id,
    name,
    email,
  };
}

export async function requireAppUser(): Promise<AppUser> {
  const user = await getAppUser();
  if (!user) {
    throw new Error("Authentication is required.");
  }

  return user;
}
