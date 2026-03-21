import { readFile } from "node:fs/promises";
import path from "node:path";

import * as XLSX from "xlsx";
import { describe, expect, it } from "vitest";

import {
  CHART_SETTINGS_SHEET_NAME,
  OFFICE_NETWORK_SHEET_NAME,
  PORTFOLIO_GANTT_MILESTONES_SHEET_NAME,
  PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME,
} from "@/lib/workbook/contracts";
import { parseWorkbookBuffer } from "@/lib/workbook/parser";
import { WorkbookValidationError } from "@/lib/workbook/types";

const V4_FIXTURE_PATH = path.resolve(process.cwd(), "fixtures", "IT_Exec_Reporting_Ingestion_Template_v4_dummy_data.xlsx");
const V3_FIXTURE_PATH = path.resolve(process.cwd(), "fixtures", "IT_Exec_Reporting_Ingestion_Template_v3_dummy_data.xlsx");

async function loadFixtureBuffer(fixturePath = V4_FIXTURE_PATH): Promise<Buffer> {
  return readFile(fixturePath);
}

function getSheetRows(workbook: XLSX.WorkBook, sheetName: string): string[][] {
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    raw: false,
    defval: "",
    blankrows: true,
  }) as string[][];
}

function setSheetRows(workbook: XLSX.WorkBook, sheetName: string, rows: string[][]): void {
  workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(rows);
}

async function createMutatedWorkbook(
  mutate: (workbook: XLSX.WorkBook) => void,
  fixturePath = V4_FIXTURE_PATH,
): Promise<Buffer> {
  const fixtureBuffer = await loadFixtureBuffer(fixturePath);
  const workbook = XLSX.read(fixtureBuffer, { type: "buffer", raw: false });
  mutate(workbook);
  return Buffer.from(XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }));
}

describe("parseWorkbookBuffer", () => {
  it("parses the bundled v4 workbook and derives office network metrics", async () => {
    const result = await parseWorkbookBuffer(await loadFixtureBuffer(), "fixture.xlsx");
    const juneMetric = result.snapshot.derivedNetworkMetrics.find((row) => row.reportingMonth === "2026-06");
    const networkServiceJune = result.snapshot.serviceAvailability.find(
      (row) => row.reportingMonth === "2026-06" && row.serviceName === "Network",
    );

    expect(result.snapshot.metadata.templateKey).toBe("IT_EXEC_TEMPLATE_V4");
    expect(result.snapshot.metadata.templateVersion).toBe(4);
    expect(result.snapshot.currentMonth).toBe("2026-06");
    expect(result.snapshot.availableMonths).toEqual([
      "2026-01",
      "2026-02",
      "2026-03",
      "2026-04",
      "2026-05",
      "2026-06",
    ]);
    expect(result.snapshot.periods.find((row) => row.reportingMonth === "2026-06")?.reportCutOffDate).toBe("2026-06-19");
    expect(result.snapshot.portfolioGanttWorkstreams).toHaveLength(78);
    expect(result.snapshot.portfolioGanttMilestones).toHaveLength(48);
    expect(result.snapshot.chartSettings).toHaveLength(6);
    expect(result.snapshot.chartSettings[0]).toMatchObject({
      page: "Support Operations",
      chartKey: "support_ticket_volumes",
      overlayEnabled: true,
      overlayMetric: "Close Balance %",
      rollingWindow: 3,
      healthyMin: 100,
      amberMin: 97,
    });
    expect(juneMetric).toMatchObject({
      availabilityPct: 99.97,
      outageMinutes: 303,
      majorIncidents: 0,
      perfectOffices: 19,
      below99_9Offices: 2,
      below99Offices: 0,
      worstOffice: "Cardiff",
      worstAvailabilityPct: 99.62,
    });
    expect(networkServiceJune).toMatchObject({
      targetPct: 99.9,
      availabilityPct: 99.97,
      outageMinutes: 303,
    });
  });

  it("continues to parse a legacy v3 workbook without chart settings", async () => {
    const result = await parseWorkbookBuffer(await loadFixtureBuffer(V3_FIXTURE_PATH), "legacy-v3.xlsx");

    expect(result.snapshot.metadata.templateKey).toBe("IT_EXEC_TEMPLATE_V3");
    expect(result.snapshot.metadata.templateVersion).toBe(3);
    expect(result.snapshot.chartSettings).toEqual([]);
    expect(result.snapshot.portfolioGanttWorkstreams).toHaveLength(78);
  });

  it("rejects an invalid template version", async () => {
    const buffer = await createMutatedWorkbook((workbook) => {
      const rows = getSheetRows(workbook, "README");
      const versionRowIndex = rows.findIndex((row) => row[0] === "Template Version");
      rows[versionRowIndex][1] = "1";
      setSheetRows(workbook, "README", rows);
    });

    await expect(parseWorkbookBuffer(buffer, "bad-version.xlsx", { skipTableValidation: true })).rejects.toMatchObject({
      name: "WorkbookValidationError",
      issues: expect.arrayContaining(["README Template Version must be one of 3, 4."]),
    } satisfies Partial<WorkbookValidationError>);
  });

  it("rejects a header mismatch on row 3", async () => {
    const buffer = await createMutatedWorkbook((workbook) => {
      const rows = getSheetRows(workbook, "Periods");
      rows[2][0] = "Month";
      setSheetRows(workbook, "Periods", rows);
    });

    await expect(parseWorkbookBuffer(buffer, "bad-header.xlsx", { skipTableValidation: true })).rejects.toThrow(
      'Periods header mismatch at column 1. Expected "Reporting Month", received "Month".',
    );
  });

  it("rejects manual Network rows in INPUT_Service_Availability", async () => {
    const buffer = await createMutatedWorkbook((workbook) => {
      const rows = getSheetRows(workbook, "INPUT_Service_Availability");
      rows.push([
        "2026-06",
        "Network",
        "Network",
        "99.95%",
        "99.90%",
        "4",
        "0",
        "",
        "",
        "Should not be here.",
      ]);
      setSheetRows(workbook, "INPUT_Service_Availability", rows);
    });

    await expect(parseWorkbookBuffer(buffer, "manual-network.xlsx", { skipTableValidation: true })).rejects.toMatchObject({
      name: "WorkbookValidationError",
      issues: expect.arrayContaining([
        'INPUT_Service_Availability must not contain manual "Network" rows in supported workbook templates.',
      ]),
    } satisfies Partial<WorkbookValidationError>);
  });

  it("rejects a missing office network row for an in-scope office", async () => {
    const buffer = await createMutatedWorkbook((workbook) => {
      const rows = getSheetRows(workbook, OFFICE_NETWORK_SHEET_NAME);
      const filtered = rows.filter((row, index) => {
        if (index < 3) {
          return true;
        }

        return !(row[0] === "2026-06" && row[1] === "Cardiff");
      });
      setSheetRows(workbook, OFFICE_NETWORK_SHEET_NAME, filtered);
    });

    await expect(parseWorkbookBuffer(buffer, "missing-office.xlsx", { skipTableValidation: true })).rejects.toMatchObject({
      name: "WorkbookValidationError",
      issues: expect.arrayContaining(["Missing office network row for Cardiff in 2026-06."]),
    } satisfies Partial<WorkbookValidationError>);
  });

  it("rejects duplicate office network rows for the same office-month", async () => {
    const buffer = await createMutatedWorkbook((workbook) => {
      const rows = getSheetRows(workbook, OFFICE_NETWORK_SHEET_NAME);
      const duplicate = rows.find((row, index) => index >= 3 && row[0] === "2026-06" && row[1] === "Cardiff");

      if (!duplicate) {
        throw new Error("Unable to locate Cardiff June row in fixture.");
      }

      rows.push([...duplicate]);
      setSheetRows(workbook, OFFICE_NETWORK_SHEET_NAME, rows);
    });

    await expect(parseWorkbookBuffer(buffer, "duplicate-office.xlsx", { skipTableValidation: true })).rejects.toMatchObject({
      name: "WorkbookValidationError",
      issues: expect.arrayContaining(["Duplicate office network rows found for Cardiff in 2026-06."]),
    } satisfies Partial<WorkbookValidationError>);
  });

  it("rejects invalid percentage values", async () => {
    const buffer = await createMutatedWorkbook((workbook) => {
      const rows = getSheetRows(workbook, OFFICE_NETWORK_SHEET_NAME);
      const rowIndex = rows.findIndex((row, index) => index >= 3 && row[0] === "2026-06" && row[1] === "Cardiff");
      rows[rowIndex][2] = "N/A";
      setSheetRows(workbook, OFFICE_NETWORK_SHEET_NAME, rows);
    });

    await expect(parseWorkbookBuffer(buffer, "invalid-pct.xlsx", { skipTableValidation: true })).rejects.toThrow(
      `${OFFICE_NETWORK_SHEET_NAME}.Availability % must be a percentage.`,
    );
  });

  it("rejects an invalid portfolio gantt domain", async () => {
    const buffer = await createMutatedWorkbook((workbook) => {
      const rows = getSheetRows(workbook, PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME);
      rows[3][3] = "Unknown domain";
      setSheetRows(workbook, PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME, rows);
    });

    await expect(parseWorkbookBuffer(buffer, "invalid-gantt-domain.xlsx", { skipTableValidation: true })).rejects.toThrow(
      `${PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME}.Domain must be one of:`,
    );
  });

  it("rejects a gantt milestone without a matching workstream in the same month", async () => {
    const buffer = await createMutatedWorkbook((workbook) => {
      const rows = getSheetRows(workbook, PORTFOLIO_GANTT_MILESTONES_SHEET_NAME);
      rows[3][1] = "Nonexistent Workstream";
      setSheetRows(workbook, PORTFOLIO_GANTT_MILESTONES_SHEET_NAME, rows);
    });

    await expect(parseWorkbookBuffer(buffer, "orphan-milestone.xlsx", { skipTableValidation: true })).rejects.toMatchObject({
      name: "WorkbookValidationError",
      issues: expect.arrayContaining([
        'Portfolio Gantt milestone "Secondary path live" does not match a workstream named "Nonexistent Workstream" in 2026-01.',
      ]),
    } satisfies Partial<WorkbookValidationError>);
  });

  it("rejects a gantt workstream with an invalid date range", async () => {
    const buffer = await createMutatedWorkbook((workbook) => {
      const rows = getSheetRows(workbook, PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME);
      rows[3][5] = "2026-03-01";
      rows[3][6] = "2026-02-01";
      setSheetRows(workbook, PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME, rows);
    });

    await expect(parseWorkbookBuffer(buffer, "bad-gantt-dates.xlsx", { skipTableValidation: true })).rejects.toThrow(
      `${PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME}.Start Date must be on or before End Date.`,
    );
  });

  it("rejects a missing report cut-off date", async () => {
    const buffer = await createMutatedWorkbook((workbook) => {
      const rows = getSheetRows(workbook, "Periods");
      rows[3][5] = "";
      setSheetRows(workbook, "Periods", rows);
    });

    await expect(parseWorkbookBuffer(buffer, "missing-cutoff.xlsx", { skipTableValidation: true })).rejects.toThrow(
      "Periods.Report Cut-Off Date is required.",
    );
  });

  it("rejects an unsupported chart overlay metric", async () => {
    const buffer = await createMutatedWorkbook((workbook) => {
      const rows = getSheetRows(workbook, CHART_SETTINGS_SHEET_NAME);
      rows[3][4] = "Ticket Average";
      setSheetRows(workbook, CHART_SETTINGS_SHEET_NAME, rows);
    });

    await expect(parseWorkbookBuffer(buffer, "bad-chart-metric.xlsx", { skipTableValidation: true })).rejects.toMatchObject({
      name: "WorkbookValidationError",
      issues: expect.arrayContaining([`${CHART_SETTINGS_SHEET_NAME}.Overlay Metric must be one of: Close Balance %.`]),
    } satisfies Partial<WorkbookValidationError>);
  });

  it("rejects a chart overlay with an invalid rolling window", async () => {
    const buffer = await createMutatedWorkbook((workbook) => {
      const rows = getSheetRows(workbook, CHART_SETTINGS_SHEET_NAME);
      rows[3][5] = "0";
      setSheetRows(workbook, CHART_SETTINGS_SHEET_NAME, rows);
    });

    await expect(parseWorkbookBuffer(buffer, "bad-chart-window.xlsx", { skipTableValidation: true })).rejects.toMatchObject({
      name: "WorkbookValidationError",
      issues: expect.arrayContaining([`${CHART_SETTINGS_SHEET_NAME}.Rolling Window must be at least 1.`]),
    } satisfies Partial<WorkbookValidationError>);
  });

  it("rejects chart thresholds where healthy is below amber", async () => {
    const buffer = await createMutatedWorkbook((workbook) => {
      const rows = getSheetRows(workbook, CHART_SETTINGS_SHEET_NAME);
      rows[3][6] = "96";
      rows[3][7] = "98";
      setSheetRows(workbook, CHART_SETTINGS_SHEET_NAME, rows);
    });

    await expect(parseWorkbookBuffer(buffer, "bad-chart-thresholds.xlsx", { skipTableValidation: true })).rejects.toMatchObject({
      name: "WorkbookValidationError",
      issues: expect.arrayContaining([`${CHART_SETTINGS_SHEET_NAME}.Healthy Min must be greater than or equal to Amber Min.`]),
    } satisfies Partial<WorkbookValidationError>);
  });
});
