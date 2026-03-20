import { chromium } from "playwright";

import { getEnv } from "@/lib/env";
import { renderReportHtml } from "@/lib/report/render-report-html";
import type { ExecSummaryState } from "@/lib/reports/exec-summary";
import { saveGeneratedExport } from "@/lib/reports/service";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

export type ExportType = "page-png" | "block-png" | "full-pdf";

interface ExportArtifactInput {
  reportId: string;
  reportTitle: string;
  snapshot: NormalizedReportSnapshot;
  exportType: ExportType;
  month: string;
  pageId?: string;
  blockId?: string;
  persist?: boolean;
  execSummary?: ExecSummaryState;
}

export interface ExportArtifactResult {
  buffer: Buffer;
  contentType: string;
  filename: string;
}

function slugify(value: string): string {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .replace(/-{2,}/g, "-");
}

export async function exportReportArtifact(input: ExportArtifactInput): Promise<ExportArtifactResult> {
  const env = getEnv();
  const html = await renderReportHtml(input.snapshot, {
    month: input.month,
    initialPageId: input.pageId ?? "p-summary",
    showAllPages: input.exportType === "full-pdf",
    hideChrome: true,
    execSummary: input.execSummary,
  });

  const browser = await chromium.launch({
    headless: true,
    executablePath: env.PLAYWRIGHT_BROWSER_PATH,
  });

  try {
    const page = await browser.newPage({
      viewport: {
        width: 1440,
        height: 2200,
      },
    });

    await page.setContent(html, { waitUntil: "networkidle" });
    await page.waitForFunction(() => (window as typeof window & { __REPORT_READY?: boolean }).__REPORT_READY === true);

    let buffer: Buffer;
    let contentType: string;
    let filename: string;

    if (input.exportType === "full-pdf") {
      await page.emulateMedia({ media: "screen" });
      const pdf = await page.pdf({
        landscape: true,
        format: "A4",
        margin: {
          top: "0.25in",
          right: "0.25in",
          bottom: "0.25in",
          left: "0.25in",
        },
        printBackground: true,
      });

      buffer = Buffer.from(pdf);
      contentType = "application/pdf";
      filename = `${slugify(input.reportTitle)}-${input.month}-full-report.pdf`;
    } else if (input.exportType === "page-png") {
      if (!input.pageId) {
        throw new Error("pageId is required for page-png exports.");
      }

      const locator = page.locator(`#${input.pageId}`);
      await locator.waitFor();
      buffer = Buffer.from(await locator.screenshot({ type: "png" }));
      contentType = "image/png";
      filename = `${slugify(input.reportTitle)}-${input.month}-${input.pageId}.png`;
    } else {
      if (!input.blockId) {
        throw new Error("blockId is required for block-png exports.");
      }

      const locator = page.locator(`#${input.blockId}`);
      await locator.waitFor();
      buffer = Buffer.from(await locator.screenshot({ type: "png" }));
      contentType = "image/png";
      filename = `${slugify(input.reportTitle)}-${input.month}-${input.blockId}.png`;
    }

    if (input.persist !== false) {
      await saveGeneratedExport({
        reportId: input.reportId,
        exportType: input.exportType,
        month: input.month,
        pageId: input.pageId,
        blockId: input.blockId,
        contentType,
        data: buffer,
      });
    }

    return {
      buffer,
      contentType,
      filename,
    };
  } finally {
    await browser.close();
  }
}
