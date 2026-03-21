import { chromium } from "playwright";
import PptxGenJS from "pptxgenjs";

import { getEnv } from "@/lib/env";
import { getReportSlides, getSlideId } from "@/lib/report/blocks";
import { renderReportHtml } from "@/lib/report/render-report-html";
import type { ExecSummaryState } from "@/lib/reports/exec-summary";
import { saveGeneratedExport } from "@/lib/reports/service";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

export type ExportType = "page-png" | "block-png" | "full-pdf" | "full-pptx";

const PPTX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const PPTX_LAYOUT_WIDTH = 13.333;
const PPTX_LAYOUT_HEIGHT = 7.5;

interface ExportArtifactInput {
  reportId: string;
  reportTitle: string;
  snapshot: NormalizedReportSnapshot;
  exportType: ExportType;
  month: string;
  pageId?: string;
  tabId?: string | null;
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

function toNodeBuffer(payload: string | ArrayBuffer | Blob | Uint8Array): Buffer {
  if (typeof payload === "string") {
    return Buffer.from(payload, "binary");
  }

  if (payload instanceof Uint8Array) {
    return Buffer.from(payload);
  }

  if (payload instanceof ArrayBuffer) {
    return Buffer.from(new Uint8Array(payload));
  }

  if (typeof Blob !== "undefined" && payload instanceof Blob) {
    throw new Error("Unexpected Blob output while generating PPTX in Node.js.");
  }

  throw new Error("Unsupported PPTX output payload.");
}

async function renderSlideDeckPptx(page: import("playwright").Page): Promise<Buffer> {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "OpenAI Codex";
  pptx.company = "TeacherActive";
  pptx.subject = "IT Reporting slide deck";
  pptx.title = "TeacherActive IT Reporting";

  for (const slide of getReportSlides()) {
    const locator = page.locator(`#${slide.id}`);
    await locator.waitFor();
    const pngBuffer = Buffer.from(await locator.screenshot({ type: "png" }));
    const pptSlide = pptx.addSlide();
    pptSlide.addImage({
      data: `data:image/png;base64,${pngBuffer.toString("base64")}`,
      x: 0,
      y: 0,
      w: PPTX_LAYOUT_WIDTH,
      h: PPTX_LAYOUT_HEIGHT,
    });
  }

  const rawBuffer = await pptx.write({ outputType: "nodebuffer", compression: true });
  return toNodeBuffer(rawBuffer);
}

export async function exportReportArtifact(input: ExportArtifactInput): Promise<ExportArtifactResult> {
  const slideId = input.pageId ? getSlideId(input.pageId, input.tabId) : undefined;
  const env = getEnv();
  const html = await renderReportHtml(input.snapshot, {
    month: input.month,
    initialPageId: input.pageId ?? "p-summary",
    initialTabId: input.tabId,
    showAllPages: input.exportType === "full-pdf" || input.exportType === "full-pptx",
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
        width: 1800,
        height: 1100,
      },
    });

    await page.setContent(html, { waitUntil: "networkidle" });
    await page.waitForFunction(() => (window as typeof window & { __REPORT_READY?: boolean }).__REPORT_READY === true);

    let buffer: Buffer;
    let contentType: string;
    let filename: string;

    if (input.exportType === "full-pdf") {
      await page.emulateMedia({ media: "print" });
      const pdf = await page.pdf({
        width: "13.333in",
        height: "7.5in",
        preferCSSPageSize: true,
        margin: {
          top: "0in",
          right: "0in",
          bottom: "0in",
          left: "0in",
        },
        printBackground: true,
      });

      buffer = Buffer.from(pdf);
      contentType = "application/pdf";
      filename = `${slugify(input.reportTitle)}-${input.month}-full-report.pdf`;
    } else if (input.exportType === "full-pptx") {
      buffer = await renderSlideDeckPptx(page);
      contentType = PPTX_CONTENT_TYPE;
      filename = `${slugify(input.reportTitle)}-${input.month}-full-report.pptx`;
    } else if (input.exportType === "page-png") {
      if (!input.pageId) {
        throw new Error("pageId is required for page-png exports.");
      }

      const locator = page.locator(`#${slideId}`);
      await locator.waitFor();
      buffer = Buffer.from(await locator.screenshot({ type: "png" }));
      contentType = "image/png";
      filename = `${slugify(input.reportTitle)}-${input.month}-${slideId}.png`;
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
        pageId: slideId ?? input.pageId,
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
