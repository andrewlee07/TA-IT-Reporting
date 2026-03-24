import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { acknowledgePlatformAlert } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; alertId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, alertId } = await params;
    const alert = await acknowledgePlatformAlert({ tenantSlug, alertId, request });
    return NextResponse.json({ alert });
  } catch (error) {
    return jsonError(error, "Failed to acknowledge alert.", 400);
  }
}
