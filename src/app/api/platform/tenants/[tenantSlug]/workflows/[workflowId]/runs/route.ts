import { NextResponse } from "next/server";
import { z } from "zod";

import { jsonError } from "@/lib/platform/http";
import { listWorkflowRuns, queueWorkflowRun } from "@/lib/platform/service";

const runSchema = z.object({
  payload: z.record(z.string(), z.unknown()).optional(),
});

interface RouteProps {
  params: Promise<{ tenantSlug: string; workflowId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, workflowId } = await params;
    const runs = await listWorkflowRuns({
      tenantSlug,
      workflowId,
      request,
    });

    return NextResponse.json({ runs });
  } catch (error) {
    return jsonError(error, "Failed to load workflow runs.", 400);
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, workflowId } = await params;
    const payload = runSchema.parse(await request.json().catch(() => ({})));
    const run = await queueWorkflowRun({
      tenantSlug,
      workflowId,
      payload: payload.payload,
      request,
    });

    return NextResponse.json({ run });
  } catch (error) {
    return jsonError(error, "Failed to queue workflow run.", 400);
  }
}
