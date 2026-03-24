import { NextResponse } from "next/server";

import { workflowTestCaseSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { listWorkflowTests, saveWorkflowTestCaseDefinition } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const tests = await listWorkflowTests({ tenantSlug, request });
    return NextResponse.json({ tests });
  } catch (error) {
    return jsonError(error, "Failed to load workflow tests.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const payload = workflowTestCaseSchema.parse(await request.json().catch(() => ({})));
    const testCase = await saveWorkflowTestCaseDefinition({
      tenantSlug,
      request,
      testCase: payload,
    });
    return NextResponse.json({ testCase });
  } catch (error) {
    return jsonError(error, "Failed to save workflow test.", 400);
  }
}
