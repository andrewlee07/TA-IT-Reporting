import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { getNotificationDeliveryDetail } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; deliveryId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, deliveryId } = await params;
    const detail = await getNotificationDeliveryDetail({ tenantSlug, deliveryId, request });
    return NextResponse.json({ detail });
  } catch (error) {
    return jsonError(error, "Failed to load delivery detail.");
  }
}
