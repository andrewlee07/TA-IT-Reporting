import { NextResponse } from "next/server";

import { pageSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { listPageDefinitions, savePageDefinition } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const pages = await listPageDefinitions({ tenantSlug, request });
    return NextResponse.json({ pages });
  } catch (error) {
    return jsonError(error, "Failed to load pages.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const page = pageSchema.parse(await request.json());
    const savedPage = await savePageDefinition({
      tenantSlug,
      page,
      request,
    });
    return NextResponse.json({ page: savedPage });
  } catch (error) {
    return jsonError(error, "Failed to save page definition.", 400);
  }
}
