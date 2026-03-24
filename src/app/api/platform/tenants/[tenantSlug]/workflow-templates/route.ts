import { NextResponse } from "next/server";
import { z } from "zod";

import { jsonError } from "@/lib/platform/http";
import { listWorkflowTemplates, saveWorkflowTemplateDefinition } from "@/lib/platform/service";

const saveWorkflowTemplateSchema = z.object({
  workflowId: z.string().trim().min(1),
  templateKey: z.string().trim().min(1).optional(),
  name: z.string().trim().min(1).optional(),
  description: z.string().trim().min(1).optional(),
});

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const templates = await listWorkflowTemplates({ tenantSlug, request });
    return NextResponse.json({ templates });
  } catch (error) {
    return jsonError(error, "Failed to load workflow templates.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const payload = saveWorkflowTemplateSchema.parse(await request.json().catch(() => ({})));
    const template = await saveWorkflowTemplateDefinition({
      tenantSlug,
      request,
      ...payload,
    });
    return NextResponse.json({ template });
  } catch (error) {
    return jsonError(error, "Failed to save workflow template.", 400);
  }
}
