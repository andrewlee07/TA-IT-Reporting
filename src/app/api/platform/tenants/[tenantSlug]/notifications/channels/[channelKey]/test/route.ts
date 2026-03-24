import { NextResponse } from "next/server";

import { channelTestSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { testNotificationChannel } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; channelKey: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, channelKey } = await params;
    const payload = channelTestSchema.parse(await request.json().catch(() => ({})));
    const delivery = await testNotificationChannel({
      tenantSlug,
      channelKey,
      subject: payload.subject,
      body: payload.body,
      request,
    });
    return NextResponse.json({ delivery });
  } catch (error) {
    return jsonError(error, "Failed to test notification channel.", 400);
  }
}
