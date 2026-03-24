import { NextResponse } from "next/server";
import { z } from "zod";

import { jsonError } from "@/lib/platform/http";
import { listSectionTemplates, saveSectionTemplateDefinition } from "@/lib/platform/service";

const saveSectionTemplateSchema = z.object({
  layoutKey: z.string().trim().min(1),
  sectionId: z.string().trim().min(1),
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
    const templates = await listSectionTemplates({ tenantSlug, request });
    return NextResponse.json({ templates });
  } catch (error) {
    return jsonError(error, "Failed to load section templates.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const payload = saveSectionTemplateSchema.parse(await request.json().catch(() => ({})));
    const template = await saveSectionTemplateDefinition({
      tenantSlug,
      request,
      ...payload,
    });
    return NextResponse.json({ template });
  } catch (error) {
    return jsonError(error, "Failed to save section template.", 400);
  }
}
