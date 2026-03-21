import { readFile } from "node:fs/promises";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { renderReportHtml } from "@/lib/report/render-report-html";
import { createDemoExecSummary } from "@/lib/reports/exec-summary";
import { parseWorkbookBuffer } from "@/lib/workbook/parser";

const FIXTURE_PATH = path.resolve(process.cwd(), "fixtures", "IT_Exec_Reporting_Ingestion_Template_v4_dummy_data.xlsx");

describe("renderReportHtml", () => {
  it("renders a fully inlined report document for the selected month and page", async () => {
    const fixtureBuffer = await readFile(FIXTURE_PATH);
    const { snapshot } = await parseWorkbookBuffer(fixtureBuffer, "fixture.xlsx");
    const html = await renderReportHtml(snapshot, {
      month: "2026-06",
      initialPageId: "p-gantt",
      showAllPages: false,
      hideChrome: true,
    });

    expect(html).toContain("TeacherActive");
    expect(html).toContain("June 2026");
    expect(html).toContain("Portfolio Gantt");
    expect(html).toContain("12-Week Rolling Portfolio View");
    expect(html).toContain("gantt-svg");
    expect(html).toContain("window.__REPORT_READY = true;");
    expect(html).not.toContain("__REPORT_DATA__");
    expect(html).not.toContain("__REPORT_RUNTIME__");
  });

  it("preserves page tab bootstrapping and support slide ids for tabbed sections", async () => {
    const fixtureBuffer = await readFile(FIXTURE_PATH);
    const { snapshot } = await parseWorkbookBuffer(fixtureBuffer, "fixture.xlsx");
    const html = await renderReportHtml(snapshot, {
      month: "2026-06",
      initialPageId: "p-support",
      initialTabId: "detail",
      showAllPages: false,
      hideChrome: true,
    });

    expect(html).toContain('const INITIAL_PAGE_ID = \'p-support\'');
    expect(html).toContain('const INITIAL_TAB_ID = "detail"');
    expect(html).toContain('id="p-support-overview"');
    expect(html).toContain('id="p-support-detail"');
    expect(html).toContain('id="p-network-map"');
    expect(html).toContain('id="p-network-detail"');
  });

  it("inlines exec summary content when provided for export rendering", async () => {
    const fixtureBuffer = await readFile(FIXTURE_PATH);
    const { snapshot } = await parseWorkbookBuffer(fixtureBuffer, "fixture.xlsx");
    const html = await renderReportHtml(snapshot, {
      month: "2026-06",
      initialPageId: "p-summary",
      showAllPages: false,
      hideChrome: true,
      execSummary: createDemoExecSummary("2026-06"),
    });

    expect(html).toContain("Exec Summary");
    expect(html).toContain("executive narrative");
  });
});
