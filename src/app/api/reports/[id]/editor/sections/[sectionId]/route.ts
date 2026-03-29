import { NextResponse } from "next/server";
import { z } from "zod";

import { requireAppUser } from "@/lib/auth/app-user";
import { isSectionId } from "@/lib/editor/sections";
import { saveEditorSection } from "@/lib/reports/service";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const saveSectionSchema = z.object({
  baseRevisionId: z.string().nullable(),
  payload: z.unknown(),
});

interface RouteProps {
  params: Promise<{ id: string; sectionId: string }>;
}

function getMonthFromRequest(request: Request): string | null {
  const { searchParams } = new URL(request.url);
  return searchParams.get("month");
}

export async function PUT(request: Request, { params }: RouteProps) {
  try {
    const user = await requireAppUser();
    const { id, sectionId } = await params;
    const month = getMonthFromRequest(request);
    if (!month) {
      return NextResponse.json({ error: "A month query parameter is required." }, { status: 400 });
    }

    if (!isSectionId(sectionId)) {
      return NextResponse.json({ error: "Invalid section." }, { status: 400 });
    }

    const body = saveSectionSchema.parse(await request.json());
    const draft = await saveEditorSection({
      reportId: id,
      reportingMonth: month,
      sectionId,
      payload: body.payload as never,
      baseRevisionId: body.baseRevisionId,
      actor: user,
    });

    return NextResponse.json({ draft });
  } catch (caughtError) {
    const message = caughtError instanceof Error ? caughtError.message : "Failed to save section.";
    if (message.startsWith("Conflict:")) {
      const changedSections = message.replace("Conflict:", "").split(",").filter(Boolean);
      return NextResponse.json({ error: "Conflict detected.", changedSections }, { status: 409 });
    }

    const status =
      message === "Authentication is required."
        ? 401
        : message.includes("Report not found") || message.includes("Invalid month")
          ? 400
          : 500;

    return NextResponse.json({ error: message }, { status });
  }
}
