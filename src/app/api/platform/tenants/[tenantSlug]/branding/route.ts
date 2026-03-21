import { NextResponse } from "next/server";

import { brandingSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { saveBrandingDefinition } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const branding = brandingSchema.parse(await request.json());
    const savedBranding = await saveBrandingDefinition({
      tenantSlug,
      branding,
      request,
    });

    return NextResponse.json({ branding: savedBranding });
  } catch (error) {
    return jsonError(error, "Failed to save branding definition.", 400);
  }
}
