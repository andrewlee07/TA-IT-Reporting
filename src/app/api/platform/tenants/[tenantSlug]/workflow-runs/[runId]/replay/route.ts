import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { replayWorkflowRun } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; runId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, runId } = await params;
    const run = await replayWorkflowRun({
      tenantSlug,
      runId,
      request,
    });
    return NextResponse.json({ run });
  } catch (error) {
    return jsonError(error, "Failed to replay workflow run.", 400);
  }
}
