import { NextResponse } from "next/server";

import { inviteSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { createTenantInvite, listTenantInvites } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const invites = await listTenantInvites({
      tenantSlug,
      request,
    });
    return NextResponse.json({ invites });
  } catch (error) {
    return jsonError(error, "Failed to load tenant invites.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const payload = inviteSchema.parse(await request.json());
    const invite = await createTenantInvite({
      tenantSlug,
      email: payload.email,
      role: payload.role,
      expiresInDays: payload.expiresInDays,
      request,
    });
    return NextResponse.json({ invite });
  } catch (error) {
    return jsonError(error, "Failed to create tenant invite.", 400);
  }
}
