import { NextResponse } from "next/server";

import { deleteObjectDefinition } from "@/lib/platform/service";
import { jsonError } from "@/lib/platform/http";

interface RouteProps {
  params: Promise<{ tenantSlug: string; objectId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function DELETE(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, objectId } = await params;
    await deleteObjectDefinition({
      tenantSlug,
      objectId,
      request,
    });

    return NextResponse.json({ ok: true });
  } catch (error) {
    return jsonError(error, "Failed to delete object definition.", 400);
  }
}

