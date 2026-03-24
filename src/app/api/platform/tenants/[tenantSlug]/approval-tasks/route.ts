import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { listApprovalTasks } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const tasks = await listApprovalTasks({ tenantSlug, request });
    return NextResponse.json({ tasks });
  } catch (error) {
    return jsonError(error, "Failed to load approval tasks.");
  }
}
