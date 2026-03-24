import { NextResponse } from "next/server";
import { z } from "zod";

import { jsonError } from "@/lib/platform/http";
import { previewAgentInvocation } from "@/lib/platform/service";

const previewSchema = z.object({
  objectKey: z.string().trim().min(1).optional(),
  sampleSize: z.number().int().min(1).max(10).optional(),
});

interface RouteProps {
  params: Promise<{ tenantSlug: string; agentId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, agentId } = await params;
    const payload = previewSchema.parse(await request.json().catch(() => ({})));
    const preview = await previewAgentInvocation({
      tenantSlug,
      agentId,
      objectKey: payload.objectKey,
      sampleSize: payload.sampleSize,
      request,
    });

    return NextResponse.json({ preview });
  } catch (error) {
    return jsonError(error, "Failed to preview agent invocation.", 400);
  }
}
