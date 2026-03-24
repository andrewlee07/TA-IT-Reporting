import { NextResponse } from "next/server";

import { deleteFieldDefinition } from "@/lib/platform/service";
import { getQueryParam, jsonError } from "@/lib/platform/http";

interface RouteProps {
  params: Promise<{ tenantSlug: string; fieldId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function DELETE(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, fieldId } = await params;
    const objectId = getQueryParam(request, "objectId");
    if (!objectId) {
      return NextResponse.json({ error: "objectId is required." }, { status: 400 });
    }

    await deleteFieldDefinition({
      tenantSlug,
      objectId,
      fieldId,
      request,
    });

    return NextResponse.json({ ok: true });
  } catch (error) {
    return jsonError(error, "Failed to delete field definition.", 400);
  }
}

