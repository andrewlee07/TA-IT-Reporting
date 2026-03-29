import { NextResponse } from "next/server";

import { requireAppUser } from "@/lib/auth/app-user";
import { getEditableReportDraft } from "@/lib/reports/service";

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
    await requireAppUser();
    const { id } = await params;
    const month = getMonthFromRequest(request);
    if (!month) {
      return NextResponse.json({ error: "A month query parameter is required." }, { status: 400 });
    }

    const draft = await getEditableReportDraft(id, month);
    return NextResponse.json({ draft });
  } catch (caughtError) {
    const message = caughtError instanceof Error ? caughtError.message : "Failed to load editor draft.";
    const status =
      message === "Authentication is required."
        ? 401
        : message.includes("Report not found") || message.includes("Invalid month")
          ? 400
          : 500;

    return NextResponse.json({ error: message }, { status });
  }
}
