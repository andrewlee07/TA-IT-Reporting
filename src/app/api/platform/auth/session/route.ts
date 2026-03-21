import { NextResponse } from "next/server";

import { acceptInviteSessionSchema } from "@/lib/platform/api-schemas";
import { getPlatformSessionCookieSettings, PLATFORM_SESSION_COOKIE } from "@/lib/platform/auth";
import { jsonError } from "@/lib/platform/http";
import { acceptPlatformInvite, getCurrentPlatformSession } from "@/lib/platform/service";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    const session = await getCurrentPlatformSession({ request });
    return NextResponse.json({ session });
  } catch (error) {
    return jsonError(error, "Failed to load platform session.", 400);
  }
}

export async function POST(request: Request) {
  try {
    const payload = acceptInviteSessionSchema.parse(await request.json());
    const accepted = await acceptPlatformInvite(payload);
    const response = NextResponse.json({
      session: accepted.session,
      tenantSlug: accepted.tenantSlug,
    });
    response.cookies.set(PLATFORM_SESSION_COOKIE, accepted.sessionValue, getPlatformSessionCookieSettings());
    return response;
  } catch (error) {
    return jsonError(error, "Failed to create a platform session.", 400);
  }
}

export async function DELETE() {
  const response = NextResponse.json({ ok: true });
  response.cookies.set(PLATFORM_SESSION_COOKIE, "", {
    ...getPlatformSessionCookieSettings(),
    maxAge: 0,
  });
  return response;
}
