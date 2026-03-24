import { NextResponse } from "next/server";

import { formSubmissionSchema } from "@/lib/platform/api-schemas";
import { jsonError } from "@/lib/platform/http";
import { listFormSubmissions, submitFormSubmission } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; formKey: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, formKey } = await params;
    const submissions = await listFormSubmissions({ tenantSlug, formKey, request });
    return NextResponse.json({ submissions });
  } catch (error) {
    return jsonError(error, "Failed to load form submissions.");
  }
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug, formKey } = await params;
    const payload = formSubmissionSchema.parse(await request.json());
    const submission = await submitFormSubmission({
      tenantSlug,
      formKey,
      data: payload.data,
      status: payload.status,
      request,
    });
    return NextResponse.json({ submission });
  } catch (error) {
    return jsonError(error, "Failed to submit form.", 400);
  }
}
