import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { getDeadLetterDetail } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; recordId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, recordId } = await params;
    const detail = await getDeadLetterDetail({ tenantSlug, recordId, request });
    return NextResponse.json({ detail });
  } catch (error) {
    return jsonError(error, "Failed to load dead letter detail.");
  }
}
