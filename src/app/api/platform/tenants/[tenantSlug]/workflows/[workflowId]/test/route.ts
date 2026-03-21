import { NextResponse } from "next/server";
import { z } from "zod";

import { jsonError } from "@/lib/platform/http";
import { runWorkflowTest } from "@/lib/platform/service";

const workflowTestSchema = z.object({
  payload: z.record(z.string(), z.unknown()).optional(),
});

interface RouteProps {
  params: Promise<{ tenantSlug: string; workflowId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, workflowId } = await params;
    const payload = workflowTestSchema.parse(await request.json().catch(() => ({})));
    const run = await runWorkflowTest({
      tenantSlug,
      workflowId,
      payload: payload.payload,
      request,
    });
    return NextResponse.json({ run });
  } catch (error) {
    return jsonError(error, "Failed to run workflow test.", 400);
  }
}
