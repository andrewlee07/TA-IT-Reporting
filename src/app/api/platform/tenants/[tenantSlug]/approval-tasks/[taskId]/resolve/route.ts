import { NextResponse } from "next/server";
import { z } from "zod";

import { jsonError } from "@/lib/platform/http";
import { resolveApprovalTask } from "@/lib/platform/service";

const resolutionSchema = z.object({
  resolution: z.enum(["approved", "rejected"]),
});

interface RouteProps {
  params: Promise<{ tenantSlug: string; taskId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, taskId } = await params;
    const payload = resolutionSchema.parse(await request.json().catch(() => ({})));
    const task = await resolveApprovalTask({
      tenantSlug,
      taskId,
      resolution: payload.resolution,
      request,
    });
    return NextResponse.json({ task });
  } catch (error) {
    return jsonError(error, "Failed to resolve approval task.", 400);
  }
}
