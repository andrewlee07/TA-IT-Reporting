import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { listAuditEvents } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const auditEvents = await listAuditEvents({ tenantSlug, request });
    return NextResponse.json({ auditEvents });
  } catch (error) {
    return jsonError(error, "Failed to load audit events.");
  }
}

