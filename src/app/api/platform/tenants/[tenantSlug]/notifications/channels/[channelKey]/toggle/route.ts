import { NextResponse } from "next/server";

import { channelToggleSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { toggleNotificationChannel } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; channelKey: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, channelKey } = await params;
    const payload = channelToggleSchema.parse(await request.json().catch(() => ({})));
    const health = await toggleNotificationChannel({
      tenantSlug,
      channelKey,
      enabled: payload.enabled,
      request,
    });
    return NextResponse.json({ health });
  } catch (error) {
    return jsonError(error, "Failed to toggle notification channel.", 400);
  }
}
