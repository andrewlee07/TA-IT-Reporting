import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { rollbackPublishedVersion } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; versionId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, versionId } = await params;
    const version = await rollbackPublishedVersion({
      tenantSlug,
      versionId,
      request,
    });
    return NextResponse.json({ version });
  } catch (error) {
    return jsonError(error, "Failed to roll back published version.", 400);
  }
}

