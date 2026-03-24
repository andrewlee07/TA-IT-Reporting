import { NextResponse } from "next/server";

import { securityPolicySchema } from "@/lib/platform/api-schemas";
import { getSecurityPolicy, saveSecurityPolicy } from "@/lib/platform/service";
import { jsonError } from "@/lib/platform/http";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const securityPolicy = await getSecurityPolicy({ tenantSlug, request });
    return NextResponse.json({ securityPolicy });
  } catch (error) {
    return jsonError(error, "Failed to load security policy.");
  }
}

export async function PUT(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const securityPolicy = securityPolicySchema.parse(await request.json());
    const savedPolicy = await saveSecurityPolicy({
      tenantSlug,
      securityPolicy,
      request,
    });
    return NextResponse.json({ securityPolicy: savedPolicy });
  } catch (error) {
    return jsonError(error, "Failed to save security policy.", 400);
  }
}
