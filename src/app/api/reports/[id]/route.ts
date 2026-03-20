import { NextResponse } from "next/server";

import { getBundledDemoSnapshot, getStoredReport } from "@/lib/reports/service";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

interface RouteProps {
  params: Promise<{ id: string }>;
}

export async function GET(_request: Request, { params }: RouteProps) {
  const { id } = await params;

  if (id === "demo") {
    const snapshot = await getBundledDemoSnapshot();
    return NextResponse.json({
      report: {
        id: "demo",
        title: "Bundled Demo Report",
        originalFilename: snapshot.metadata.sourceFilename,
        templateKey: snapshot.metadata.templateKey,
        templateVersion: snapshot.metadata.templateVersion,
        currentMonth: snapshot.currentMonth,
        availableMonths: snapshot.availableMonths,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        snapshot,
      },
    });
  }

  const report = await getStoredReport(id);

  if (!report) {
    return NextResponse.json({ error: "Report not found." }, { status: 404 });
  }

  return NextResponse.json({ report });
}
