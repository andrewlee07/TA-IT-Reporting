import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const { renderWorkbookFromSnapshotMock } = vi.hoisted(() => ({
  renderWorkbookFromSnapshotMock: vi.fn(async () => Buffer.from("mock-workbook-binary")),
}));

vi.mock("@/lib/workbook/serialize-workbook", () => ({
  renderWorkbookFromSnapshot: renderWorkbookFromSnapshotMock,
}));

interface LoadedEditorModules {
  service: typeof import("@/lib/reports/service");
  sections: typeof import("@/lib/editor/sections");
}

const TEST_ACTOR = {
  id: "user-editor",
  name: "Editor User",
  email: "editor@example.com",
};

async function loadModules(): Promise<LoadedEditorModules> {
  const [service, sections] = await Promise.all([import("@/lib/reports/service"), import("@/lib/editor/sections")]);
  return { service, sections };
}

describe("editor draft lifecycle", () => {
  let tempStorageDir = "";

  beforeEach(async () => {
    tempStorageDir = await fs.mkdtemp(path.join(os.tmpdir(), "ta-it-reporting-tests-"));
    vi.resetModules();
    vi.clearAllMocks();
    process.env.AUTH_MODE = "development";
    process.env.STORAGE_MODE = "local";
    process.env.LOCAL_STORAGE_DIR = tempStorageDir;
    process.env.GRAPH_AUTH_MODE = "disabled";
    delete process.env.DATABASE_URL;
  });

  afterEach(async () => {
    await fs.rm(tempStorageDir, { recursive: true, force: true });
    delete process.env.AUTH_MODE;
    delete process.env.STORAGE_MODE;
    delete process.env.LOCAL_STORAGE_DIR;
    delete process.env.GRAPH_AUTH_MODE;
    delete process.env.DATABASE_URL;
    vi.resetModules();
  });

  it("creates a blank draft, persists revision metadata, and syncs workbook/json artifacts", async () => {
    const { service } = await loadModules();

    const report = await service.createBlankReportDraft({
      title: "Operations Rollup",
      initialMonth: "2026-03",
      actor: TEST_ACTOR,
    });

    await service.syncDraftArtifacts(report.id);

    const draft = await service.getEditableReportDraft(report.id, "2026-03");
    const jsonArtifact = await service.getCurrentDraftJsonArtifact(report.id);
    const workbookArtifact = await service.getCurrentDraftWorkbookArtifact(report.id);

    expect(draft.manifest.title).toBe("Operations Rollup");
    expect(draft.manifest.currentRevision.revisionNumber).toBe(1);
    expect(draft.manifest.currentRevision.changedSections).toHaveLength(8);
    expect(draft.manifest.artifactSyncStatus.state).toBe("current");
    expect(draft.snapshot.currentMonth).toBe("2026-03");
    expect(draft.snapshot.availableMonths).toEqual(["2026-03"]);
    expect(draft.snapshot.supportOperations).toHaveLength(1);
    expect(draft.snapshot.securityPatching).toHaveLength(1);
    expect(draft.snapshot.changeRelease).toHaveLength(1);
    expect(draft.snapshot.devDelivery).toHaveLength(1);
    expect(draft.snapshot.assetsLifecycle.map((row) => row.assetType)).toEqual(["Laptop", "Mobile", "Monitor"]);
    expect(JSON.parse(jsonArtifact.buffer.toString("utf8"))).toMatchObject({
      currentMonth: "2026-03",
      metadata: {
        templateVersion: 4,
      },
    });
    expect(jsonArtifact.filename).toBe("operations-rollup-2026-03-current.json");
    expect(workbookArtifact.filename).toBe("operations-rollup-2026-03.xlsx");
    expect(workbookArtifact.buffer.toString("utf8")).toBe("mock-workbook-binary");
    expect(renderWorkbookFromSnapshotMock).toHaveBeenCalled();
  });

  it("accepts stale saves for different sections and rejects stale saves for the same section", async () => {
    const { service, sections } = await loadModules();

    const report = await service.createBlankReportDraft({
      title: "Portfolio Governance",
      initialMonth: "2026-03",
      actor: TEST_ACTOR,
    });

    const initial = await service.getEditableReportDraft(report.id, "2026-03");
    const baseRevisionId = initial.manifest.currentRevision.revisionId;

    const overviewPayload = sections.getSectionPayload(
      initial.snapshot,
      "overview-setup",
      initial.manifest.title,
      initial.manifest.reportSeriesKey,
    );
    overviewPayload.title = "Portfolio Governance Updated";

    const afterOverviewSave = await service.saveEditorSection({
      reportId: report.id,
      reportingMonth: "2026-03",
      sectionId: "overview-setup",
      payload: overviewPayload,
      baseRevisionId,
      actor: TEST_ACTOR,
    });

    const financePayload = sections.getSectionPayload(
      afterOverviewSave.snapshot,
      "finance-risks",
      afterOverviewSave.manifest.title,
      afterOverviewSave.manifest.reportSeriesKey,
    );
    financePayload.topRisks = [
      {
        reportingMonth: "2026-03",
        riskIssue: "Capacity pressure",
        type: "Risk",
        owner: "Operations",
        impact: "High",
        likelihood: "Medium",
        ratingRag: "Amber",
        currentControlMitigation: "Weekly review",
        targetDate: "2026-04-30",
        decisionRequired: false,
        commentary: "Monitor hiring plan",
      },
    ];

    const afterFinanceSave = await service.saveEditorSection({
      reportId: report.id,
      reportingMonth: "2026-03",
      sectionId: "finance-risks",
      payload: financePayload,
      baseRevisionId,
      actor: TEST_ACTOR,
    });

    const staleOverviewPayload = sections.getSectionPayload(
      afterFinanceSave.snapshot,
      "overview-setup",
      afterFinanceSave.manifest.title,
      afterFinanceSave.manifest.reportSeriesKey,
    );
    staleOverviewPayload.title = "This stale title should conflict";

    await expect(
      service.saveEditorSection({
        reportId: report.id,
        reportingMonth: "2026-03",
        sectionId: "overview-setup",
        payload: staleOverviewPayload,
        baseRevisionId,
        actor: TEST_ACTOR,
      }),
    ).rejects.toThrow(/Conflict:/);

    expect(afterOverviewSave.manifest.currentRevision.revisionNumber).toBe(2);
    expect(afterOverviewSave.manifest.title).toBe("Portfolio Governance Updated");
    expect(afterFinanceSave.manifest.currentRevision.revisionNumber).toBe(3);
    expect(afterFinanceSave.manifest.currentRevision.changedSections).toEqual(["finance-risks"]);
    expect(afterFinanceSave.snapshot.topRisks).toHaveLength(1);
  });
});
