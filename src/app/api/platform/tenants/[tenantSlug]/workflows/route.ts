import { NextResponse } from "next/server";

import { workflowSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { listWorkflowDefinitions, saveWorkflowDefinition } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const workflows = await listWorkflowDefinitions({ tenantSlug, request });
    return NextResponse.json({ workflows });
  } catch (error) {
    return jsonError(error, "Failed to load workflows.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const workflow = workflowSchema.parse(await request.json());
    const savedWorkflow = await saveWorkflowDefinition({
      tenantSlug,
      workflow,
      request,
    });
    return NextResponse.json({ workflow: savedWorkflow });
  } catch (error) {
    return jsonError(error, "Failed to save workflow definition.", 400);
  }
}
