import { NextResponse } from "next/server";

import { WorkbookValidationError } from "@/lib/workbook/types";
import { createReportFromWorkbookUpload, listReports } from "@/lib/reports/service";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET() {
  try {
    const reports = await listReports();
    return NextResponse.json({ reports });
  } catch (caughtError) {
    return NextResponse.json(
      { error: caughtError instanceof Error ? caughtError.message : "Failed to load reports." },
      { status: 500 },
    );
  }
}

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const workbook = formData.get("workbook");

    if (!(workbook instanceof File)) {
      return NextResponse.json({ error: "A workbook file is required." }, { status: 400 });
    }

    const buffer = Buffer.from(await workbook.arrayBuffer());
    const report = await createReportFromWorkbookUpload(workbook.name, buffer);

    return NextResponse.json({ report });
  } catch (caughtError) {
    if (caughtError instanceof WorkbookValidationError) {
      return NextResponse.json(
        { error: "Workbook validation failed.", issues: caughtError.issues },
        { status: 422 },
      );
    }

    return NextResponse.json(
      { error: caughtError instanceof Error ? caughtError.message : "Failed to create report." },
      { status: 500 },
    );
  }
}
