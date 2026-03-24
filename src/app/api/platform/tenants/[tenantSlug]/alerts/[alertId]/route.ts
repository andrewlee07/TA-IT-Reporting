import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { getPlatformAlertDetail } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; alertId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, alertId } = await params;
    const detail = await getPlatformAlertDetail({ tenantSlug, alertId, request });
    return NextResponse.json({ detail });
  } catch (error) {
    return jsonError(error, "Failed to load alert detail.");
  }
}
