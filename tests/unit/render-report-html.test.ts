import { readFile } from "node:fs/promises";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { renderReportHtml } from "@/lib/report/render-report-html";
import { parseWorkbookBuffer } from "@/lib/workbook/parser";

const FIXTURE_PATH = path.resolve(process.cwd(), "fixtures", "IT_Exec_Reporting_Ingestion_Template_v3_dummy_data.xlsx");

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
});
