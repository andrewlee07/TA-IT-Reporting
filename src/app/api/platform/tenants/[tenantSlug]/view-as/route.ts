import { NextResponse } from "next/server";

import {
  createSignedPlatformViewAsValue,
  getPlatformSessionCookieSettings,
  PLATFORM_VIEW_AS_COOKIE,
} from "@/lib/platform/auth";
import { viewAsSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    await params;
    const payload = viewAsSchema.parse(await request.json());
    const response = NextResponse.json({ ok: true });

    if (!payload.role || !payload.personaLabel || payload.active === false) {
      response.cookies.set({
        name: PLATFORM_VIEW_AS_COOKIE,
        value: "",
        ...getPlatformSessionCookieSettings(),
        maxAge: 0,
      });
      return response;
    }

    response.cookies.set({
      name: PLATFORM_VIEW_AS_COOKIE,
      value: createSignedPlatformViewAsValue({
        role: payload.role,
        personaLabel: payload.personaLabel,
        actorEmail: payload.actorEmail,
      }),
      ...getPlatformSessionCookieSettings(),
    });

    return response;
  } catch (error) {
    return jsonError(error, "Failed to update view-as state.", 400);
  }
}
