import { NextResponse } from "next/server";

import { objectSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { listObjectDefinitions, saveObjectDefinition } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const objects = await listObjectDefinitions({ tenantSlug, request });
    return NextResponse.json({ objects });
  } catch (error) {
    return jsonError(error, "Failed to load objects.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const object = objectSchema.parse(await request.json());
    const savedObject = await saveObjectDefinition({
      tenantSlug,
      object,
      request,
    });

    return NextResponse.json({ object: savedObject });
  } catch (error) {
    return jsonError(error, "Failed to save object definition.", 400);
  }
}
