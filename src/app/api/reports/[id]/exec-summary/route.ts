import { NextResponse } from "next/server";
import { z } from "zod";

import { getBundledDemoSnapshot, getExecSummaryState, getStoredReport, saveExecSummary } from "@/lib/reports/service";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const saveSummarySchema = z.object({
  contentHtml: z.string(),
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

    if (id === "demo") {
      const snapshot = await getBundledDemoSnapshot();
      if (!snapshot.availableMonths.includes(month)) {
        return NextResponse.json({ error: "Invalid month." }, { status: 400 });
      }
    } else {
      const report = await getStoredReport(id);
      if (!report) {
        return NextResponse.json({ error: "Report not found." }, { status: 404 });
      }
      if (!report.availableMonths.includes(month)) {
        return NextResponse.json({ error: "Invalid month." }, { status: 400 });
      }
    }

    const summary = await getExecSummaryState(id, month);
    return NextResponse.json({ summary });
  } catch (caughtError) {
    return NextResponse.json(
      { error: caughtError instanceof Error ? caughtError.message : "Failed to load exec summary." },
      { status: 500 },
    );
  }
}

export async function PUT(request: Request, { params }: RouteProps) {
  try {
    const { id } = await params;
    const month = getMonthFromRequest(request);
    if (!month) {
      return NextResponse.json({ error: "A month query parameter is required." }, { status: 400 });
    }

    const body = saveSummarySchema.parse(await request.json());
    const summary = await saveExecSummary(id, month, body.contentHtml);
    return NextResponse.json({ summary });
  } catch (caughtError) {
    const message = caughtError instanceof Error ? caughtError.message : "Failed to save exec summary.";
    const status = message.includes("read-only") || message.includes("Invalid month") || message.includes("Report not found")
      ? 400
      : 500;

    return NextResponse.json({ error: message }, { status });
  }
}
