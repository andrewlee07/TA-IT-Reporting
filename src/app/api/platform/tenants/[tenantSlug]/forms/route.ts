import { NextResponse } from "next/server";

import { formSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { listFormDefinitions, saveFormDefinition } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const forms = await listFormDefinitions({ tenantSlug, request });
    return NextResponse.json({ forms });
  } catch (error) {
    return jsonError(error, "Failed to load forms.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const form = formSchema.parse(await request.json());
    const savedForm = await saveFormDefinition({ tenantSlug, form, request });
    return NextResponse.json({ form: savedForm });
  } catch (error) {
    return jsonError(error, "Failed to save form definition.", 400);
  }
}
