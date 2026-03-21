import { NextResponse } from "next/server";

import { menuSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { listMenuDefinitions, saveMenuDefinition } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const menus = await listMenuDefinitions({ tenantSlug, request });
    return NextResponse.json({ menus });
  } catch (error) {
    return jsonError(error, "Failed to load menus.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const menu = menuSchema.parse(await request.json());
    const savedMenu = await saveMenuDefinition({
      tenantSlug,
      menu,
      request,
    });
    return NextResponse.json({ menu: savedMenu });
  } catch (error) {
    return jsonError(error, "Failed to save menu definition.", 400);
  }
}
