import { NextResponse } from "next/server";
import { z } from "zod";

import { getReportPrepView, saveReportPrepAcknowledgements } from "@/lib/reports/service";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const savePrepSchema = z.object({
  acknowledgedCheckIds: z.array(z.string()).default([]),
});

interface RouteProps {
  params: Promise<{ id: string }>;
}

function getMonthFromRequest(request: Request): string | null {
  const { searchParams } = new URL(request.url);
  return searchParams.get("month");
}

export async function GET(request: Request, { params }: RouteProps) {
  try {
    const { id } = await params;
    const month = getMonthFromRequest(request);
    if (!month) {
      return NextResponse.json({ error: "A month query parameter is required." }, { status: 400 });
    }

    const prep = await getReportPrepView(id, month);
    return NextResponse.json({ prep });
  } catch (caughtError) {
    const message = caughtError instanceof Error ? caughtError.message : "Failed to load prep data.";
    const status = message.includes("Invalid month") || message.includes("Report not found") ? 400 : 500;
    return NextResponse.json({ error: message }, { status });
  }
}

export async function PUT(request: Request, { params }: RouteProps) {
  try {
    const { id } = await params;
    const month = getMonthFromRequest(request);
    if (!month) {
      return NextResponse.json({ error: "A month query parameter is required." }, { status: 400 });
    }

    const body = savePrepSchema.parse(await request.json());
    const prep = await saveReportPrepAcknowledgements(id, month, body.acknowledgedCheckIds);
    return NextResponse.json({ prep });
  } catch (caughtError) {
    const message = caughtError instanceof Error ? caughtError.message : "Failed to save prep acknowledgements.";
    const status = message.includes("read-only") || message.includes("Invalid month") || message.includes("Report not found")
      ? 400
      : 500;

    return NextResponse.json({ error: message }, { status });
  }
}
