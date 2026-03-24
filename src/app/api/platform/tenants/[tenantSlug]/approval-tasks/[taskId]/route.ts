import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { getApprovalTaskDetail } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; taskId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, taskId } = await params;
    const detail = await getApprovalTaskDetail({ tenantSlug, taskId, request });
    return NextResponse.json({ detail });
  } catch (error) {
    return jsonError(error, "Failed to load approval task detail.");
  }
}
