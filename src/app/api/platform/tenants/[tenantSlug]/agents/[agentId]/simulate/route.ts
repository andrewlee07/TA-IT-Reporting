import { NextResponse } from "next/server";

import { agentSimulationSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { simulateAgentRun } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; agentId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, agentId } = await params;
    const payload = agentSimulationSchema.parse(await request.json().catch(() => ({})));
    const result = await simulateAgentRun({
      tenantSlug,
      agentId,
      prompt: payload.prompt,
      objectKey: payload.objectKey,
      sampleSize: payload.sampleSize,
      request,
    });
    return NextResponse.json(result);
  } catch (error) {
    return jsonError(error, "Failed to simulate agent run.", 400);
  }
}
