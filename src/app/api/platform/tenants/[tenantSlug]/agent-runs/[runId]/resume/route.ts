import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { resumeBlockedAgentRun } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; runId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, runId } = await params;
    const run = await resumeBlockedAgentRun({
      tenantSlug,
      runId,
      request,
    });
    return NextResponse.json({ run });
  } catch (error) {
    return jsonError(error, "Failed to resume agent run.", 400);
  }
}
