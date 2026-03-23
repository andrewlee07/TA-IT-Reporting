import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { listAgentRuns } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const runs = await listAgentRuns({ tenantSlug, request });
    return NextResponse.json({ runs });
  } catch (error) {
    return jsonError(error, "Failed to load agent runs.");
  }
}
