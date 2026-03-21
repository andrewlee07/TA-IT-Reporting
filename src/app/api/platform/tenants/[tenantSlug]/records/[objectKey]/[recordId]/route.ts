import { NextResponse } from "next/server";

import { recordSchema } from "@/lib/platform/api-schemas";
import { deletePlatformRecord, savePlatformRecord } from "@/lib/platform/service";
import { jsonError } from "@/lib/platform/http";

interface RouteProps {
  params: Promise<{ tenantSlug: string; objectKey: string; recordId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function PUT(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, objectKey, recordId } = await params;
    const payload = recordSchema.parse(await request.json());
    const record = await savePlatformRecord({
      tenantSlug,
      objectKey,
      recordId,
      data: payload.data,
      request,
    });
    return NextResponse.json({ record });
  } catch (error) {
    return jsonError(error, "Failed to update record.", 400);
  }
}

export async function DELETE(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, objectKey, recordId } = await params;
    await deletePlatformRecord({
      tenantSlug,
      objectKey,
      recordId,
      request,
    });
    return NextResponse.json({ ok: true });
  } catch (error) {
    return jsonError(error, "Failed to delete record.", 400);
  }
}
