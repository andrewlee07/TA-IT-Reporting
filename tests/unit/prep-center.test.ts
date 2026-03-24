import { readFile } from "node:fs/promises";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { buildReportPrepView, filterAcknowledgeableCheckIds } from "@/lib/reports/prep-center";
import type { ExecSummaryState } from "@/lib/reports/exec-summary";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

const SNAPSHOT_PATH = path.resolve(process.cwd(), "fixtures", "demo-snapshot.json");

async function loadSnapshot(): Promise<NormalizedReportSnapshot> {
  return JSON.parse(await readFile(SNAPSHOT_PATH, "utf8")) as NormalizedReportSnapshot;
}

describe("prep center view model", () => {
  it("flags a missing exec summary as a blocking readiness issue", async () => {
    const snapshot = await loadSnapshot();
    const prep = buildReportPrepView({
      snapshot,
      reportingMonth: "2026-06",
      currentSummary: {
        mode: "empty",
        contentHtml: "",
        excerpt: "",
        updatedAt: null,
        sourceReportId: null,
      },
      previousMonthSummary: {
        mode: "explicit",
        contentHtml: "<p>May summary.</p>",
        excerpt: "May summary.",
        updatedAt: "2026-05-18T09:00:00.000Z",
        sourceReportId: null,
      },
      mode: "editable",
    });

    expect(prep.readiness.summary.blockingCount).toBeGreaterThan(0);
    expect(prep.readiness.checks.some((check) => check.id === "summary-missing")).toBe(true);
    expect(prep.rollover.previousExecSummary?.available).toBe(true);
  });

  it("only persists acknowledgeable warning ids", async () => {
    const snapshot = await loadSnapshot();
    const currentSummary: ExecSummaryState = {
      mode: "carried-forward",
      contentHtml: "<p>Inherited.</p>",
      excerpt: "Inherited.",
      updatedAt: "2026-06-18T09:00:00.000Z",
      sourceReportId: "prior-report",
    };

    const prep = buildReportPrepView({
      snapshot,
      reportingMonth: "2026-06",
      currentSummary,
      previousMonthSummary: {
        mode: "explicit",
        contentHtml: "<p>Previous.</p>",
        excerpt: "Previous.",
        updatedAt: "2026-05-18T09:00:00.000Z",
        sourceReportId: null,
      },
      mode: "editable",
      acknowledgedCheckIds: ["summary-carried-forward", "summary-missing", "rollover-no-previous-month"],
    });

    expect(filterAcknowledgeableCheckIds([...prep.readiness.checks, ...prep.readiness.reviewedChecks], prep.acknowledgedCheckIds)).toEqual([
      "summary-carried-forward",
    ]);
    expect(prep.readiness.summary.warningCount).toBeGreaterThanOrEqual(0);
  });

  it("shows a clean rollover empty state when there is no previous month", async () => {
    const snapshot = await loadSnapshot();
    const singleMonthSnapshot: NormalizedReportSnapshot = {
      ...snapshot,
      availableMonths: ["2026-06"],
    };

    const prep = buildReportPrepView({
      snapshot: singleMonthSnapshot,
      reportingMonth: "2026-06",
      currentSummary: {
        mode: "explicit",
        contentHtml: "<p>Current month summary.</p>",
        excerpt: "Current month summary.",
        updatedAt: "2026-06-20T09:00:00.000Z",
        sourceReportId: null,
      },
      previousMonthSummary: null,
      mode: "editable",
    });

    expect(prep.rollover.previousMonth).toBeNull();
    expect(prep.rollover.reviewQueue).toHaveLength(0);
    expect(prep.readiness.checks.some((check) => check.id === "rollover-no-previous-month")).toBe(true);
  });
});
