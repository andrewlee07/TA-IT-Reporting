import { NextResponse } from "next/server";

export function getQueryParam(request: Request, key: string): string | undefined {
  const { searchParams } = new URL(request.url);
  return searchParams.get(key) ?? undefined;
}

export function jsonError(error: unknown, fallbackMessage: string, status = 500): NextResponse {
  const resolvedStatus =
    typeof error === "object" && error !== null && "status" in error && typeof error.status === "number" ? error.status : status;

  return NextResponse.json(
    {
      error: error instanceof Error ? error.message : fallbackMessage,
    },
    { status: resolvedStatus },
  );
}
