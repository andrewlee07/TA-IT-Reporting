import { NextResponse } from "next/server";

import { requireAppUser } from "@/lib/auth/app-user";
import { updateEditorPresence } from "@/lib/reports/service";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

interface RouteProps {
  params: Promise<{ id: string }>;
}

function getMonthFromRequest(request: Request): string | null {
  const { searchParams } = new URL(request.url);
  return searchParams.get("month");
}

export async function PUT(request: Request, { params }: RouteProps) {
  try {
    const user = await requireAppUser();
    const { id } = await params;
    const month = getMonthFromRequest(request);
    if (!month) {
      return NextResponse.json({ error: "A month query parameter is required." }, { status: 400 });
    }

    const activePresence = await updateEditorPresence(id, month, user);
    return NextResponse.json({ activePresence });
  } catch (caughtError) {
    const message = caughtError instanceof Error ? caughtError.message : "Failed to update presence.";
    const status = message === "Authentication is required." ? 401 : 500;
    return NextResponse.json({ error: message }, { status });
  }
}
