import { NextResponse } from "next/server";

import { layoutSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { listLayoutDefinitions, saveLayoutDefinition } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const layouts = await listLayoutDefinitions({ tenantSlug, request });
    return NextResponse.json({ layouts });
  } catch (error) {
    return jsonError(error, "Failed to load layouts.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const layout = layoutSchema.parse(await request.json());
    const savedLayout = await saveLayoutDefinition({
      tenantSlug,
      layout,
      request,
    });

    return NextResponse.json({ layout: savedLayout });
  } catch (error) {
    return jsonError(error, "Failed to save layout definition.", 400);
  }
}
