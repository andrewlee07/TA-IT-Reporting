import { NextResponse } from "next/server";

import { providerSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { listModelProviders, saveModelProviderDefinition } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const modelProviders = await listModelProviders({ tenantSlug, request });
    return NextResponse.json({ modelProviders });
  } catch (error) {
    return jsonError(error, "Failed to load model providers.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const provider = providerSchema.parse(await request.json());
    const savedProvider = await saveModelProviderDefinition({
      tenantSlug,
      provider,
      request,
    });
    return NextResponse.json({ modelProvider: savedProvider });
  } catch (error) {
    return jsonError(error, "Failed to save model provider.", 400);
  }
}
