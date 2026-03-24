import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { listPlatformAlerts } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const alerts = await listPlatformAlerts({ tenantSlug, request });
    return NextResponse.json({ alerts });
  } catch (error) {
    return jsonError(error, "Failed to load alerts.");
  }
}
