import { NextResponse } from "next/server";
import { z } from "zod";

import { isValidBlockId, isValidPageId, resolveTabId } from "@/lib/report/blocks";
import { exportReportArtifact } from "@/lib/reports/export-service";
import { getBundledDemoSnapshot, getExecSummaryState, getStoredReport } from "@/lib/reports/service";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const exportSchema = z.object({
  exportType: z.enum(["page-png", "block-png", "full-pdf", "full-pptx", "full-pptx-editable"]),
  month: z.string(),
  pageId: z.string().optional(),
  tabId: z.string().optional(),
  blockId: z.string().optional(),
});

interface RouteProps {
  params: Promise<{ id: string }>;
}

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { id } = await params;
    const payload = exportSchema.parse(await request.json());
    const bundledDemo = id === "demo" ? await getBundledDemoSnapshot() : null;
    const stored = bundledDemo ? null : await getStoredReport(id);

    if (!bundledDemo && !stored) {
      return NextResponse.json({ error: "Report not found." }, { status: 404 });
    }

    const snapshot = bundledDemo ?? stored!.snapshot;

    if (!snapshot.availableMonths.includes(payload.month)) {
      return NextResponse.json({ error: "Invalid month." }, { status: 400 });
    }

    if (payload.exportType !== "full-pdf" && payload.exportType !== "full-pptx" && payload.exportType !== "full-pptx-editable") {
      if (!payload.pageId || !isValidPageId(payload.pageId)) {
        return NextResponse.json({ error: "A valid pageId is required." }, { status: 400 });
      }
    }

    const resolvedTabId = payload.pageId ? resolveTabId(payload.pageId, payload.tabId) : null;

    if (payload.pageId && payload.tabId && resolvedTabId !== payload.tabId) {
      return NextResponse.json({ error: "A valid tabId is required for this page." }, { status: 400 });
    }

    if (payload.exportType === "block-png") {
      if (!payload.pageId || !payload.blockId || !isValidBlockId(payload.pageId, payload.blockId, resolvedTabId)) {
        return NextResponse.json({ error: "A valid blockId is required for block exports." }, { status: 400 });
      }
    }

    const artifact = await exportReportArtifact({
      reportId: id,
      reportTitle: bundledDemo ? "bundled-demo-report" : stored!.title,
      snapshot,
      exportType: payload.exportType,
      month: payload.month,
      pageId: payload.pageId,
      tabId: resolvedTabId,
      blockId: payload.blockId,
      persist: id !== "demo",
      execSummary: await getExecSummaryState(id, payload.month),
    });

    return new NextResponse(new Uint8Array(artifact.buffer), {
      headers: {
        "content-type": artifact.contentType,
        "content-disposition": `attachment; filename="${artifact.filename}"`,
      },
    });
  } catch (caughtError) {
    return NextResponse.json(
      { error: caughtError instanceof Error ? caughtError.message : "Export failed." },
      { status: 500 },
    );
  }
}
