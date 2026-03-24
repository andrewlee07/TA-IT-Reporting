import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { getWorkflowRunDetail } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; runId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, runId } = await params;
    const detail = await getWorkflowRunDetail({ tenantSlug, runId, request });
    return NextResponse.json({ detail });
  } catch (error) {
    return jsonError(error, "Failed to load workflow run detail.");
  }
}
