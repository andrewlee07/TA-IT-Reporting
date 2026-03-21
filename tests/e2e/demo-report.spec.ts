import { readFileSync } from "node:fs";
import path from "node:path";

import { expect, test } from "@playwright/test";
import JSZip from "jszip";
import { PDFDocument } from "pdf-lib";

import { getReportSlides } from "../../src/lib/report/blocks";
import { createLocalReport } from "../../src/lib/reports/local-report-store";
import { parseWorkbookBuffer } from "../../src/lib/workbook/parser";
import type { NormalizedReportSnapshot } from "../../src/lib/workbook/types";

const DEMO_SNAPSHOT_PATH = path.resolve(process.cwd(), "fixtures", "demo-snapshot.json");
const V3_FIXTURE_PATH = path.resolve(process.cwd(), "fixtures", "IT_Exec_Reporting_Ingestion_Template_v3_dummy_data.xlsx");

async function seedLegacyV2Report() {
  const snapshot = JSON.parse(readFileSync(DEMO_SNAPSHOT_PATH, "utf8")) as NormalizedReportSnapshot;
  const legacySnapshot = {
    ...snapshot,
    metadata: {
      ...snapshot.metadata,
      templateKey: "IT_EXEC_TEMPLATE_V2",
      templateVersion: 2,
      sourceFilename: "IT_Exec_Reporting_Ingestion_Template_v2_dummy_data.xlsx",
    },
    periods: snapshot.periods.map((period) => ({
      ...period,
      reportCutOffDate: period.monthEndDate,
    })),
    portfolioGanttWorkstreams: [],
    portfolioGanttMilestones: [],
  };

  const report = await createLocalReport({
    title: "IT_Exec_Reporting_Ingestion_Template_v2_dummy_data · 2026-06",
    originalFilename: "IT_Exec_Reporting_Ingestion_Template_v2_dummy_data.xlsx",
    templateKey: "IT_EXEC_TEMPLATE_V2",
    templateVersion: 2,
    currentMonth: legacySnapshot.currentMonth,
    availableMonths: legacySnapshot.availableMonths,
    snapshot: legacySnapshot,
    workbookObjectKey: "workbooks/legacy-v2-dummy.xlsx",
  });

  return report.id;
}

async function seedDemoReport(originalFilename = "IT_Exec_Reporting_Ingestion_Template_v4_dummy_data.xlsx") {
  const snapshot = JSON.parse(readFileSync(DEMO_SNAPSHOT_PATH, "utf8")) as NormalizedReportSnapshot;
  const report = await createLocalReport({
    title: `${originalFilename.replace(/\.xlsx$/, "")} · ${snapshot.currentMonth}`,
    originalFilename,
    templateKey: snapshot.metadata.templateKey,
    templateVersion: snapshot.metadata.templateVersion,
    currentMonth: snapshot.currentMonth,
    availableMonths: snapshot.availableMonths,
    snapshot,
    workbookObjectKey: `workbooks/${originalFilename}`,
  });

  return report.id;
}

async function seedLegacyV3Report(originalFilename = "IT_Exec_Reporting_Ingestion_Template_v3_dummy_data.xlsx") {
  const workbook = readFileSync(V3_FIXTURE_PATH);
  const { snapshot } = await parseWorkbookBuffer(workbook, originalFilename);
  const report = await createLocalReport({
    title: `${originalFilename.replace(/\.xlsx$/, "")} · ${snapshot.currentMonth}`,
    originalFilename,
    templateKey: snapshot.metadata.templateKey,
    templateVersion: snapshot.metadata.templateVersion,
    currentMonth: snapshot.currentMonth,
    availableMonths: snapshot.availableMonths,
    snapshot,
    workbookObjectKey: `workbooks/${originalFilename}`,
  });

  return report.id;
}

test("bundled demo report renders directly in the app shell", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-exec");

  await expect(page.locator("iframe")).toHaveCount(0);
  await expect(page.locator(".sidebar")).toBeVisible();
  await expect(page.locator("#report-month-select")).toHaveValue("2026-06");
  await expect(page.locator(".nav-link.active")).toContainText("Executive Scorecard");
  await expect(page.locator(".report-page.active .ph-title")).toHaveText("Executive IT Scorecard");
  await expect(page.locator(".report-page.active .ph-title-accent")).toHaveText("Overview");
  await expect(page.locator(".report-page.active .ph-meta-line")).toContainText("Reporting Period");
  await expect(page.locator(".report-page.active .ph-brand")).toHaveCount(0);
  await expect(page.locator("#exec-svc-grid .svc-tile")).toHaveCount(6);
  await expect(page.locator("#exec-kpis .kc-spark")).toHaveCount(5);
  await expect(page.locator(".report-page.active .sl-title")).toBeHidden();

  const sidebarWidth = await page.locator(".sidebar").evaluate((element) => Math.round(element.getBoundingClientRect().width));
  expect(sidebarWidth).toBeGreaterThanOrEqual(372);

  const boxShadow = await page.locator(".report-page.active").evaluate((element) => getComputedStyle(element).boxShadow);
  expect(boxShadow).toBe("none");
});

test("major KPI strips render sparkline trends derived from workbook history", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-support&tab=overview");

  await expect(page.locator("#support-kpis .kc-spark")).toHaveCount(4);
  expect(await page.locator("#support-kpi-backlog .kc-spark line").count()).toBeGreaterThan(0);

  await page.goto("/?report=demo&month=2026-06&page=p-security");
  await expect(page.locator("#sec-kpis .kc-spark")).toHaveCount(4);
});

test("support ticket volumes renders the workbook-driven close-balance overlay for v4 reports", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-support&tab=overview");

  await expect(page.locator("#support-vol-legend")).toContainText("3M close balance %");
  await expect(page.locator("#support-vol-legend .legend-item")).toHaveCount(3);
});

test("legacy v3 reports keep the original support volume chart without the overlay", async ({ page }) => {
  const reportId = await seedLegacyV3Report();

  await page.goto(`/?report=${reportId}&month=2026-06&page=p-support&tab=overview`);
  await expect(page.locator("#support-vol-legend")).not.toContainText("close balance");
  await expect(page.locator("#support-vol-legend .legend-item")).toHaveCount(2);
});

test("support operations uses section tabs as separate slides and syncs them into the URL", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-support&tab=overview");

  const activeSlide = page.locator(".report-page.active");
  await expect(activeSlide).toHaveAttribute("id", "p-support-overview");
  await expect(activeSlide.locator(".slide-tab.active")).toHaveText("Overview");
  await expect(activeSlide.locator(".ph-title")).toHaveText("Support Operations");
  await expect(activeSlide.locator(".ph-title-accent")).toHaveText("Overview");
  await expect(activeSlide.locator("#support-hero")).toBeVisible();
  await expect(activeSlide.locator("#support-tickets-block")).toHaveCount(0);

  await activeSlide.locator(".slide-tab", { hasText: "Ticket Detail" }).click();

  await expect(page).toHaveURL(/page=p-support&tab=detail$/);
  const detailSlide = page.locator(".report-page.active");
  await expect(detailSlide).toHaveAttribute("id", "p-support-detail");
  await expect(detailSlide.locator(".slide-tab.active")).toHaveText("Ticket Detail");
  await expect(detailSlide.locator(".ph-title-accent")).toHaveText("Ticket Detail");
  await expect(detailSlide.locator("#support-tickets-block")).toBeVisible();
  await expect(detailSlide.locator("#support-hero")).toHaveCount(0);

  await page.goBack();
  await expect(page).toHaveURL(/page=p-support&tab=overview$/);
  await expect(page.locator(".report-page.active")).toHaveAttribute("id", "p-support-overview");
});

test("root route prefers demo or newer reports over a legacy v2 saved upload", async ({ page }) => {
  const legacyReportId = await seedLegacyV2Report();

  await page.goto("/");
  await expect(page).not.toHaveURL(new RegExp(`report=${legacyReportId}$`));
  await expect(page).not.toHaveURL(new RegExp(`report=${legacyReportId}&`));
  await expect(page.locator(".report-page.active .ph-title")).toHaveText("Exec Summary");
  await expect(page.locator(".nav-link.active")).toContainText("Exec Summary");
});

test("bundled demo renders a read-only exec summary as the first page", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-summary");

  await expect(page.locator(".report-page.active .ph-title")).toHaveText("Exec Summary");
  await expect(page.locator("#summary-state-badge")).toContainText("Bundled example");
  await expect(page.locator("#summary-content")).toContainText("executive narrative");
  await expect(page.locator(".summary-readonly-pill")).toContainText("read only");
});

test("saved reports can add and persist an exec summary through the UI", async ({ page }) => {
  const reportId = await seedDemoReport(`exec-summary-ui-seed-${Date.now()}.xlsx`);

  await page.goto(`/?report=${reportId}&month=2026-06&page=p-summary`);

  const actionButton = page
    .getByRole("button", { name: /Add exec summary|Review inherited draft|Edit summary/ })
    .first();
  await actionButton.click();
  await page.locator(".summary-editor").click();
  await page.keyboard.press("Meta+A");
  await page.keyboard.press("Backspace");
  await page.keyboard.type("Executive summary drafted in the app.");
  await page.getByRole("button", { name: "Save" }).click();

  await expect(page.locator("#summary-content")).toContainText("Executive summary drafted in the app.");
  await expect(page.locator("#summary-state-badge")).toContainText("Saved summary");

  await page.reload();
  await expect(page.locator("#summary-content")).toContainText("Executive summary drafted in the app.");
});

test("exec summaries carry forward for refreshed uploads in the same report family", async ({ request }) => {
  const originalFilename = "carry-forward-pack.xlsx";
  const reportId = await seedDemoReport(originalFilename);

  const saveResponse = await request.put(`/api/reports/${reportId}/exec-summary?month=2026-06`, {
    data: {
      contentHtml: "<p><strong>Carry-forward</strong> candidate narrative.</p>",
    },
  });

  expect(saveResponse.ok()).toBeTruthy();

  const refreshedReportId = await seedDemoReport(originalFilename);
  const summaryResponse = await request.get(`/api/reports/${refreshedReportId}/exec-summary?month=2026-06`);

  expect(summaryResponse.ok()).toBeTruthy();
  const payload = (await summaryResponse.json()) as { summary: { mode: string; contentHtml: string } };
  expect(payload.summary.mode).toBe("carried-forward");
  expect(payload.summary.contentHtml).toContain("Carry-forward");
});

test("sidebar collapses into an icon rail and persists across refresh", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-network");

  const expandedWidthBeforeCollapse = await page.locator(".sidebar").evaluate((element) => Math.round(element.getBoundingClientRect().width));
  await page.getByRole("button", { name: "Collapse sidebar" }).click();
  await expect(page.locator(".shell")).toHaveClass(/sidebar-collapsed/);
  await page.waitForTimeout(350);

  const collapsedWidth = await page.locator(".sidebar").evaluate((element) => Math.round(element.getBoundingClientRect().width));
  expect(collapsedWidth).toBeLessThan(expandedWidthBeforeCollapse / 2);

  await expect(page.locator(".nav-link.active")).toHaveAttribute("title", "Network & Offices");
  await expect(page.locator(".nav-text").first()).toBeHidden();
  await expect(page.locator(".shell")).toHaveClass(/sidebar-use-icons/);
  await expect(page.locator(".nav-link.active .nav-icon-glyph")).toBeVisible();
  await expect(page.locator(".nav-link.active .nav-icon-label")).toBeHidden();
  await page.locator(".nav-link.active").hover();
  await expect(page.locator(".nav-link.active .nav-tooltip")).toBeVisible();
  await expect(page.locator(".nav-link.active .nav-tooltip")).toHaveText("Network & Offices");
  const sidebarBox = await page.locator(".sidebar").boundingBox();
  const tooltipBox = await page.locator(".nav-link.active .nav-tooltip").boundingBox();
  expect(sidebarBox).not.toBeNull();
  expect(tooltipBox).not.toBeNull();
  expect(tooltipBox!.x).toBeGreaterThan(sidebarBox!.x + sidebarBox!.width - 2);

  await page.getByRole("button", { name: "Expand sidebar" }).click();
  await page.getByRole("button", { name: "Initials" }).click();
  await page.getByRole("button", { name: "Collapse sidebar" }).click();
  await expect(page.locator(".shell")).toHaveClass(/sidebar-use-initials/);
  await expect(page.locator(".nav-link.active .nav-icon-label")).toBeVisible();
  await expect(page.locator(".nav-link.active .nav-icon-glyph")).toBeHidden();

  await page.reload();
  await expect(page.locator(".shell")).toHaveClass(/sidebar-collapsed/);
  await expect(page.locator(".shell")).toHaveClass(/sidebar-use-initials/);
  await page.waitForTimeout(350);

  const persistedWidth = await page.locator(".sidebar").evaluate((element) => Math.round(element.getBoundingClientRect().width));
  expect(persistedWidth).toBeLessThan(expandedWidthBeforeCollapse / 2);
  await expect(page.locator(".nav-link.active .nav-icon-label")).toBeVisible();

  await page.getByRole("button", { name: "Expand sidebar" }).click();
  await page.waitForTimeout(350);
  const expandedWidth = await page.locator(".sidebar").evaluate((element) => Math.round(element.getBoundingClientRect().width));
  expect(expandedWidth).toBeGreaterThanOrEqual(372);
});

test("legacy report route redirects into the canonical root route", async ({ page }) => {
  await page.goto("/reports/demo?month=2026-06&page=p-network");

  await expect(page).toHaveURL(/\/\?report=demo&month=2026-06&page=p-network&tab=map$/);
  await expect(page.locator(".report-page.active .ph-title")).toHaveText("Network & Office Availability");
});

test("network & offices uses section tabs as separate slides and syncs them into the URL", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-network&tab=map");

  const activeSlide = page.locator(".report-page.active");
  await expect(activeSlide).toHaveAttribute("id", "p-network-map");
  await expect(activeSlide.locator(".slide-tab.active")).toHaveText("Map View");
  await expect(activeSlide.locator("#network-map-block")).toBeVisible();
  await expect(activeSlide.locator("#office-list-block")).toHaveCount(0);

  await activeSlide.locator(".slide-tab", { hasText: "Office Detail" }).click();

  await expect(page).toHaveURL(/page=p-network&tab=detail$/);
  const detailSlide = page.locator(".report-page.active");
  await expect(detailSlide).toHaveAttribute("id", "p-network-detail");
  await expect(detailSlide.locator(".slide-tab.active")).toHaveText("Office Detail");
  await expect(detailSlide.locator("#office-list-block")).toBeVisible();
  await expect(detailSlide.locator("#network-map-block")).toHaveCount(0);
});

test("portfolio gantt renders as a first-class report page with summary cards", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-gantt");

  await expect(page.locator(".nav-link.active")).toContainText("Portfolio Gantt");
  await expect(page.locator(".report-page.active .ph-title")).toHaveText("Portfolio Gantt");
  await expect(page.locator("#gantt-svg")).toBeVisible();
  await expect(page.locator("#gantt-summary .kc")).toHaveCount(4);
  await expect(page.locator("#gantt-sub")).toContainText("active workstreams");
  await expect(page.locator("#gantt-period-label")).toContainText("2026");
});

test("portfolio gantt exposes hover details for workstreams and milestones", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-gantt");

  const workstreamTarget = page.locator('#gantt-svg .gantt-hover-target[data-hover-id^="gantt-workstream-"]').first();
  await workstreamTarget.evaluate((element) => {
    element.dispatchEvent(new MouseEvent("mouseenter", { bubbles: true }));
    element.dispatchEvent(new MouseEvent("mouseover", { bubbles: true }));
  });

  await expect(page.locator("#gantt-tooltip")).toHaveClass(/active/);
  await expect(page.locator("#gantt-tooltip")).toContainText("WAN Resilience Uplift");
  await expect(page.locator("#gantt-tooltip")).toContainText("Head of IT");

  const milestoneTarget = page.locator('#gantt-svg .gantt-hover-target[data-hover-id*="-milestone-"]').first();
  await milestoneTarget.evaluate((element) => {
    element.dispatchEvent(new MouseEvent("mouseenter", { bubbles: true }));
    element.dispatchEvent(new MouseEvent("mouseover", { bubbles: true }));
  });

  await expect(page.locator("#gantt-tooltip")).toContainText("Milestone");
});

test("legacy v2 reports show a compatibility empty state on portfolio gantt", async ({ page }) => {
  const legacyReportId = await seedLegacyV2Report();

  await page.goto(`/?report=${legacyReportId}&month=2026-06&page=p-gantt`);

  await expect(page.locator("#gantt-empty-state")).toHaveClass(/active/);
  await expect(page.locator("#gantt-empty-copy")).toContainText("legacy workbook");
  await expect(page.locator("#gantt-open-demo-link")).toHaveAttribute("href", /report=demo/);
  await expect(page.getByRole("button", { name: "Upload v4 workbook" })).toBeVisible();
  await expect(page.locator("#gantt-summary .kc")).toHaveCount(0);
});

test("prototype export mode is integrated into the report shell and clears selection on page change", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-support&tab=overview");

  await page.getByRole("button", { name: "Select to Export" }).click();
  await expect(page.locator(".shell")).toHaveClass(/export-mode/);

  await page.locator("#support-kpi-opened").click();
  await page.locator("#support-kpi-closed").click();

  await expect(page.getByText("2 items selected")).toBeVisible();
  await expect(page.locator(".exportable.selected")).toHaveCount(2);

  await page.locator(".nav-link", { hasText: "Service Availability" }).click();

  await expect(page.getByText("0 items selected")).toBeVisible();
  await expect(page.locator(".report-page.active .ph-title")).toHaveText("Service Availability");
  await expect(page.locator(".exportable.selected")).toHaveCount(0);
});

test("client-side export downloads a single panel and a multi-select composite", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-support&tab=overview");

  await page.locator("#support-kpi-opened").hover();
  const [singleDownload] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#support-kpi-opened .export-icon").click(),
  ]);

  expect(singleDownload.suggestedFilename()).toMatch(/\.png$/);

  await page.goto("/?report=demo&month=2026-06&page=p-support&tab=volumes");
  await page.getByRole("button", { name: "Select to Export" }).click();
  await page.locator("#support-vol-block").click();
  await page.locator("#support-detail-note-block").click();

  const [multiDownload] = await Promise.all([
    page.waitForEvent("download"),
    page.getByRole("button", { name: "Export Selected" }).click(),
  ]);

  expect(multiDownload.suggestedFilename()).toMatch(/p-support-volumes-selection\.png$/);
});

test("network page renders a neutral gb map asset with stable office plotting", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-network&tab=map");

  await expect(page.locator("#network-map-block svg image")).toHaveAttribute("href", "/maps/great-britain-locator.svg");
  await expect(page.locator("#office-dots circle")).toHaveCount(22);

  const mapSvg = await page.evaluate(async () => {
    const response = await fetch("/maps/great-britain-locator.svg");
    return response.text();
  });

  expect(mapSvg).not.toContain("#ff1d1d");

  const newcastle = page.locator('#office-dots circle[data-office-name="Newcastle"]');
  await expect(newcastle).toHaveAttribute("cx", /48[0-9](\.\d)?/);
  await expect(newcastle).toHaveAttribute("cy", /57[0-9](\.\d)?/);

  const cardiff = page.locator('#office-dots circle[data-office-name="Cardiff"]');
  await expect(cardiff).toHaveAttribute("cx", /39[0-9](\.\d)?/);
  await expect(cardiff).toHaveAttribute("cy", /90[0-9](\.\d)?/);

  const norwich = page.locator('#office-dots circle[data-office-name="Norwich"]');
  await expect(norwich).toHaveAttribute("cx", /58[0-9](\.\d)?/);
  await expect(norwich).toHaveAttribute("cy", /77[0-9](\.\d)?/);
});

test("every concrete slide fits within a single slide canvas", async ({ page }) => {
  for (const slide of getReportSlides()) {
    await test.step(slide.slideLabel, async () => {
      const params = new URLSearchParams({
        report: "demo",
        month: "2026-06",
        page: slide.pageId,
      });

      if (slide.tabId) {
        params.set("tab", slide.tabId);
      }

      await page.goto(`/?${params.toString()}`);
      const activeSlide = page.locator(".report-page.active");
      await expect(activeSlide).toBeVisible();

      const overflow = await activeSlide.evaluate((element) => ({
        slide: element.getBoundingClientRect().toJSON(),
        contentBounds: Array.from(element.querySelectorAll<HTMLElement>("*"))
          .filter((child) => {
            const style = window.getComputedStyle(child);
            if (style.display === "none" || style.visibility === "hidden") {
              return false;
            }

            const rect = child.getBoundingClientRect();
            return rect.width > 0 && rect.height > 0;
          })
          .reduce(
            (bounds, child) => {
              const rect = child.getBoundingClientRect();
              return {
                right: Math.max(bounds.right, rect.right),
                bottom: Math.max(bounds.bottom, rect.bottom),
              };
            },
            {
              right: element.getBoundingClientRect().left,
              bottom: element.getBoundingClientRect().top,
            },
          ),
      }));

      expect(overflow.contentBounds.bottom).toBeLessThanOrEqual(overflow.slide.bottom + 2);
      expect(overflow.contentBounds.right).toBeLessThanOrEqual(overflow.slide.right + 2);
    });
  }
});

test("demo export endpoints return binary artifacts", async ({ request }) => {
  const pngResponse = await request.post("/api/reports/demo/exports", {
    data: {
      exportType: "page-png",
      month: "2026-06",
      pageId: "p-exec",
    },
  });

  expect(pngResponse.ok()).toBeTruthy();
  expect(pngResponse.headers()["content-type"]).toContain("image/png");

  const pngBuffer = await pngResponse.body();
  expect(pngBuffer.byteLength).toBeGreaterThan(50_000);

  const pdfResponse = await request.post("/api/reports/demo/exports", {
    data: {
      exportType: "full-pdf",
      month: "2026-06",
    },
  });

  expect(pdfResponse.ok()).toBeTruthy();
  expect(pdfResponse.headers()["content-type"]).toContain("application/pdf");

  const pdfBuffer = await pdfResponse.body();
  expect(pdfBuffer.byteLength).toBeGreaterThan(50_000);

  const pdf = await PDFDocument.load(pdfBuffer);
  expect(pdf.getPageCount()).toBe(getReportSlides().length);

  const pptxResponse = await request.post("/api/reports/demo/exports", {
    data: {
      exportType: "full-pptx",
      month: "2026-06",
    },
  });

  expect(pptxResponse.ok()).toBeTruthy();
  expect(pptxResponse.headers()["content-type"]).toContain(
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  );
  expect(pptxResponse.headers()["content-disposition"]).toContain(".pptx");

  const pptxBuffer = await pptxResponse.body();
  expect(pptxBuffer.byteLength).toBeGreaterThan(100_000);

  const pptxZip = await JSZip.loadAsync(pptxBuffer);
  const slideEntries = Object.keys(pptxZip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
  expect(slideEntries).toHaveLength(getReportSlides().length);
});
