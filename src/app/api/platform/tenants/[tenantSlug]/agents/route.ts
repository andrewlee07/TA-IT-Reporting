import { NextResponse } from "next/server";

import { agentSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { listAgentDefinitions, saveAgentDefinition } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const agents = await listAgentDefinitions({ tenantSlug, request });
    return NextResponse.json({ agents });
  } catch (error) {
    return jsonError(error, "Failed to load agents.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const agent = agentSchema.parse(await request.json());
    const savedAgent = await saveAgentDefinition({
      tenantSlug,
      agent,
      request,
    });

    return NextResponse.json({ agent: savedAgent });
  } catch (error) {
    return jsonError(error, "Failed to save agent definition.", 400);
  }
}
