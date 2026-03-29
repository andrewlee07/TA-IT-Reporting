import { NextResponse } from "next/server";

import { reportAnnotationsSaveSchema } from "@/lib/annotations/validation";
import { getBundledDemoSnapshot, getReportAnnotationsState, getStoredReport, saveReportAnnotationsState } from "@/lib/reports/service";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

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

    const annotationState = await getReportAnnotationsState(id, month);
    return NextResponse.json({ annotationState });
  } catch (caughtError) {
    return NextResponse.json(
      { error: caughtError instanceof Error ? caughtError.message : "Failed to load annotations." },
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

    const body = reportAnnotationsSaveSchema.parse(await request.json());
    const annotationState = await saveReportAnnotationsState({
      reportId: id,
      reportingMonth: month,
      baseRevisionId: body.baseRevisionId,
      annotations: body.annotations,
    });

    return NextResponse.json({ annotationState });
  } catch (caughtError) {
    const message = caughtError instanceof Error ? caughtError.message : "Failed to save annotations.";

    if (message.startsWith("Conflict:")) {
      return NextResponse.json({ error: "Conflict detected." }, { status: 409 });
    }

    const status = message.includes("Report not found") ? 404 : message.includes("read-only") || message.includes("Invalid month") ? 400 : 500;

    return NextResponse.json({ error: message }, { status });
  }
}
