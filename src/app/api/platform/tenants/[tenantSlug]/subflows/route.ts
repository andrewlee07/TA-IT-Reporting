import { NextResponse } from "next/server";

import { subflowSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { listSubflows, saveSubflowDefinition } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const subflows = await listSubflows({ tenantSlug, request });
    return NextResponse.json({ subflows });
  } catch (error) {
    return jsonError(error, "Failed to load subflows.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const payload = subflowSchema.parse(await request.json().catch(() => ({})));
    const subflow = await saveSubflowDefinition({
      tenantSlug,
      request,
      subflow: payload,
    });
    return NextResponse.json({ subflow });
  } catch (error) {
    return jsonError(error, "Failed to save subflow.", 400);
  }
}
