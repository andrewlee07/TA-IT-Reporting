import { NextResponse } from "next/server";
import { z } from "zod";

import { publishDraftManifest } from "@/lib/platform/service";
import { jsonError } from "@/lib/platform/http";

const publishSchema = z.object({
  notes: z.string().trim().min(1).optional(),
});

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const payload = publishSchema.parse(await request.json().catch(() => ({})));
    const version = await publishDraftManifest({
      tenantSlug,
      notes: payload.notes,
      request,
    });
    return NextResponse.json({ version });
  } catch (error) {
    return jsonError(error, "Failed to publish draft manifest.", 400);
  }
}
