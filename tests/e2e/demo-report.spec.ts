import { expect, test } from "@playwright/test";

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

  await page.reload();
  await expect(page.locator(".shell")).toHaveClass(/sidebar-collapsed/);
  await page.waitForTimeout(350);

  const persistedWidth = await page.locator(".sidebar").evaluate((element) => Math.round(element.getBoundingClientRect().width));
  expect(persistedWidth).toBeLessThan(expandedWidthBeforeCollapse / 2);

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
