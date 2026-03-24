import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { listCostLedgerRecords } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const entries = await listCostLedgerRecords({ tenantSlug, request });
    return NextResponse.json({ entries });
  } catch (error) {
    return jsonError(error, "Failed to load cost ledger.");
  }
}
