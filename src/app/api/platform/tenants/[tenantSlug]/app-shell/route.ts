import { NextResponse } from "next/server";

import { appShellSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { getAppShellDefinition, saveAppShellDefinition } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const appShell = await getAppShellDefinition({ tenantSlug, request });
    return NextResponse.json({ appShell });
  } catch (error) {
    return jsonError(error, "Failed to load app shell.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const appShell = appShellSchema.parse(await request.json());
    const saved = await saveAppShellDefinition({
      tenantSlug,
      appShell,
      request,
    });
    return NextResponse.json({ appShell: saved });
  } catch (error) {
    return jsonError(error, "Failed to save app shell.", 400);
  }
}
