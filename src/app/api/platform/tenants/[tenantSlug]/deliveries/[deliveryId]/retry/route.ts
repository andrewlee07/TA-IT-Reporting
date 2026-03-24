import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { retryNotificationDelivery } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; deliveryId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, deliveryId } = await params;
    const delivery = await retryNotificationDelivery({ tenantSlug, deliveryId, request });
    return NextResponse.json({ delivery });
  } catch (error) {
    return jsonError(error, "Failed to retry delivery.", 400);
  }
}
