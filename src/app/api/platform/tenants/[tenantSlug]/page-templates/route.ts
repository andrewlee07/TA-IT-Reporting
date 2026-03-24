import { NextResponse } from "next/server";
import { z } from "zod";

import { jsonError } from "@/lib/platform/http";
import { listPageTemplates, savePageTemplateDefinition } from "@/lib/platform/service";

const savePageTemplateSchema = z.object({
  pageKey: z.string().trim().min(1),
  templateKey: z.string().trim().min(1).optional(),
  label: z.string().trim().min(1).optional(),
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
    const templates = await listPageTemplates({ tenantSlug, request });
    return NextResponse.json({ templates });
  } catch (error) {
    return jsonError(error, "Failed to load page templates.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const payload = savePageTemplateSchema.parse(await request.json().catch(() => ({})));
    const template = await savePageTemplateDefinition({
      tenantSlug,
      request,
      ...payload,
    });
    return NextResponse.json({ template });
  } catch (error) {
    return jsonError(error, "Failed to save page template.", 400);
  }
}
