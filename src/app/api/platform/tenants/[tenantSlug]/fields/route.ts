import { NextResponse } from "next/server";

import { fieldSchema } from "@/lib/platform/api-schemas";
import { getQueryParam, jsonError } from "@/lib/platform/http";
import { listFieldDefinitions, saveFieldDefinition } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const objectId = getQueryParam(request, "objectId");
    if (!objectId) {
      return NextResponse.json({ error: "objectId is required." }, { status: 400 });
    }

    const fields = await listFieldDefinitions({
      tenantSlug,
      objectId,
      request,
    });

    return NextResponse.json({ fields });
  } catch (error) {
    return jsonError(error, "Failed to load fields.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const payload = fieldSchema.parse(await request.json());
    const field = await saveFieldDefinition({
      tenantSlug,
      objectId: payload.objectId,
      field: payload.field,
      request,
    });

    return NextResponse.json({ field });
  } catch (error) {
    return jsonError(error, "Failed to save field definition.", 400);
  }
}
