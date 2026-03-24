import { NextResponse } from "next/server";

import { recordSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { listPlatformRecords, savePlatformRecord } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; objectKey: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, objectKey } = await params;
    const records = await listPlatformRecords({
      tenantSlug,
      objectKey,
      request,
    });
    return NextResponse.json({ records });
  } catch (error) {
    return jsonError(error, "Failed to load records.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, objectKey } = await params;
    const payload = recordSchema.parse(await request.json());
    const record = await savePlatformRecord({
      tenantSlug,
      objectKey,
      data: payload.data,
      request,
    });
    return NextResponse.json({ record });
  } catch (error) {
    return jsonError(error, "Failed to save record.", 400);
  }
}
