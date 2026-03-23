import { NextResponse } from "next/server";

import { notificationCenterSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { getNotificationCenterDefinition, saveNotificationCenterDefinition } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const notifications = await getNotificationCenterDefinition({ tenantSlug, request });
    return NextResponse.json({ notifications });
  } catch (error) {
    return jsonError(error, "Failed to load notification center.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const payload = notificationCenterSchema.parse(await request.json());
    const notifications = {
      channels: payload.channels.map((channel) => ({
        ...channel,
        id: channel.id ?? crypto.randomUUID(),
      })),
      templates: payload.templates.map((template) => ({
        ...template,
        id: template.id ?? crypto.randomUUID(),
      })),
      rules: payload.rules.map((rule) => ({
        ...rule,
        id: rule.id ?? crypto.randomUUID(),
      })),
    };
    const saved = await saveNotificationCenterDefinition({
      tenantSlug,
      notifications,
      request,
    });
    return NextResponse.json({ notifications: saved });
  } catch (error) {
    return jsonError(error, "Failed to save notification center.", 400);
  }
}
