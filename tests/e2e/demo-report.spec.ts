import { readFileSync } from "node:fs";
import path from "node:path";

import { expect, test } from "@playwright/test";

import { createLocalReport } from "../../src/lib/reports/local-report-store";
import type { NormalizedReportSnapshot } from "../../src/lib/workbook/types";

const DEMO_SNAPSHOT_PATH = path.resolve(process.cwd(), "fixtures", "demo-snapshot.json");

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

test("bundled demo report renders directly in the app shell", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-exec");

  await expect(page.locator("iframe")).toHaveCount(0);
  await expect(page.locator(".sidebar")).toBeVisible();
  await expect(page.locator("#report-month-select")).toHaveValue("2026-06");
  await expect(page.locator(".nav-link.active")).toContainText("Executive Scorecard");
  await expect(page.locator(".report-page.active .ph-title")).toHaveText("Executive IT Scorecard");
  await expect(page.locator("#exec-svc-grid .svc-tile")).toHaveCount(6);

  const sidebarWidth = await page.locator(".sidebar").evaluate((element) => Math.round(element.getBoundingClientRect().width));
  expect(sidebarWidth).toBeGreaterThanOrEqual(372);

  const boxShadow = await page.locator(".report-page.active").evaluate((element) => getComputedStyle(element).boxShadow);
  expect(boxShadow).toBe("none");
});

test("root route prefers demo or newer reports over a legacy v2 saved upload", async ({ page }) => {
  const legacyReportId = await seedLegacyV2Report();

  await page.goto("/");
  await expect(page).not.toHaveURL(new RegExp(`report=${legacyReportId}$`));
  await expect(page).not.toHaveURL(new RegExp(`report=${legacyReportId}&`));
  await expect(page.locator(".report-page.active .ph-title")).toHaveText("Executive IT Scorecard");
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

  await expect(page).toHaveURL(/\/\?report=demo&month=2026-06&page=p-network$/);
  await expect(page.locator(".report-page.active .ph-title")).toHaveText("Network & Office Availability");
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
  await workstreamTarget.hover();

  await expect(page.locator("#gantt-tooltip")).toHaveClass(/active/);
  await expect(page.locator("#gantt-tooltip")).toContainText("WAN Resilience Uplift");
  await expect(page.locator("#gantt-tooltip")).toContainText("Head of IT");

  const milestoneTarget = page.locator('#gantt-svg .gantt-hover-target[data-hover-id*="-milestone-"]').first();
  await milestoneTarget.hover();

  await expect(page.locator("#gantt-tooltip")).toContainText("Milestone");
});

test("legacy v2 reports show a compatibility empty state on portfolio gantt", async ({ page }) => {
  const legacyReportId = await seedLegacyV2Report();

  await page.goto(`/?report=${legacyReportId}&month=2026-06&page=p-gantt`);

  await expect(page.locator("#gantt-empty-state")).toHaveClass(/active/);
  await expect(page.locator("#gantt-empty-copy")).toContainText("legacy workbook");
  await expect(page.locator("#gantt-open-demo-link")).toHaveAttribute("href", /report=demo/);
  await expect(page.getByRole("button", { name: "Upload v3 workbook" })).toBeVisible();
  await expect(page.locator("#gantt-summary .kc")).toHaveCount(0);
});

test("prototype export mode is integrated into the report shell and clears selection on page change", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-support");

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
  await page.goto("/?report=demo&month=2026-06&page=p-support");

  await page.locator("#support-kpi-opened").hover();
  const [singleDownload] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#support-kpi-opened .export-icon").click(),
  ]);

  expect(singleDownload.suggestedFilename()).toMatch(/\.png$/);

  await page.getByRole("button", { name: "Select to Export" }).click();
  await page.locator("#support-kpi-opened").click();
  await page.locator("#support-cats-block").click();

  const [multiDownload] = await Promise.all([
    page.waitForEvent("download"),
    page.getByRole("button", { name: "Export Selected" }).click(),
  ]);

  expect(multiDownload.suggestedFilename()).toMatch(/p-support-selection\.png$/);
});

test("network page renders a neutral gb map asset with stable office plotting", async ({ page }) => {
  await page.goto("/?report=demo&month=2026-06&page=p-network");

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
});
