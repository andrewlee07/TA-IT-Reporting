import { NextResponse } from "next/server";
import { z } from "zod";

import { requireAppUser } from "@/lib/auth/app-user";
import { WorkbookValidationError } from "@/lib/workbook/types";
import { createBlankReportDraft, createReportFromWorkbookUpload, listReports } from "@/lib/reports/service";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const createBlankReportSchema = z.object({
  title: z.string().min(1),
  initialMonth: z.string().regex(/^\d{4}-\d{2}$/),
});

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
    const user = await requireAppUser();
    const contentType = request.headers.get("content-type") ?? "";

    if (contentType.includes("application/json")) {
      const body = createBlankReportSchema.parse(await request.json());
      const report = await createBlankReportDraft({
        title: body.title,
        initialMonth: body.initialMonth,
        actor: user,
      });

      return NextResponse.json({ report });
    }

    const formData = await request.formData();
    const workbook = formData.get("workbook");

    if (!(workbook instanceof File)) {
      return NextResponse.json({ error: "A workbook file is required." }, { status: 400 });
    }

    const buffer = Buffer.from(await workbook.arrayBuffer());
    const report = await createReportFromWorkbookUpload(workbook.name, buffer, user);

    return NextResponse.json({ report });
  } catch (caughtError) {
    if (caughtError instanceof z.ZodError) {
      return NextResponse.json({ error: "Invalid report creation payload." }, { status: 400 });
    }

    if (caughtError instanceof WorkbookValidationError) {
      return NextResponse.json(
        { error: "Workbook validation failed.", issues: caughtError.issues },
        { status: 422 },
      );
    }

    if (caughtError instanceof Error && caughtError.message === "Authentication is required.") {
      return NextResponse.json({ error: caughtError.message }, { status: 401 });
    }

    return NextResponse.json(
      { error: caughtError instanceof Error ? caughtError.message : "Failed to create report." },
      { status: 500 },
    );
  }
}
