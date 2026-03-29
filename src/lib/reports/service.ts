import { promises as fs } from "node:fs";
import path from "node:path";

import { nanoid } from "nanoid";
import { Prisma } from "@/generated/prisma/client";

import { getPrisma } from "@/lib/prisma";
import { getObjectStorage } from "@/lib/storage";
import { logger } from "@/lib/logger";
import {
  getCurrentDraftSnapshot,
  getDraftManifest,
  listDraftManifests,
  upsertDraftManifest,
  writeCurrentDraftSnapshot,
  saveDraftRevision,
  listDraftRevisions,
  listDraftPresence,
  upsertDraftPresence,
} from "@/lib/drafts/store";
import { EDITOR_SECTIONS, type ArtifactSyncStatus, type DraftManifest, type DraftUserRef, type EditableReportDraft, type EditorPresence, type ReportRevisionMeta, type SectionId } from "@/lib/drafts/types";
import { applySectionPayload, type SectionPayloadMap } from "@/lib/editor/sections";
import type { AppUser } from "@/lib/auth/app-user";
import { createBlankSnapshot } from "@/lib/workbook/blank-snapshot";
import {
  createLocalReport,
  findLocalCarryForwardExecSummary,
  getLocalExecSummary,
  getLocalPrepState,
  getLocalReport,
  listLocalReports,
  saveLocalExport,
  upsertLocalReport,
  upsertLocalPrepState,
  upsertLocalExecSummary,
} from "@/lib/reports/local-report-store";
import {
  buildExecSummaryExcerpt,
  createDemoExecSummary,
  deriveReportSeriesKey,
  sanitizeExecSummaryHtml,
  type ExecSummaryState,
} from "@/lib/reports/exec-summary";
import {
  buildReportPrepView,
  filterAcknowledgeableCheckIds,
  type ReportPrepView,
} from "@/lib/reports/prep-center";
import { parseWorkbookBuffer } from "@/lib/workbook/parser";
import { buildDerivedNetworkServiceRows, deriveNetworkMetrics, NETWORK_SERVICE_NAME } from "@/lib/workbook/derived-network";
import { renderWorkbookFromSnapshot } from "@/lib/workbook/serialize-workbook";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

export interface ReportListItem {
  id: string;
  title: string;
  originalFilename: string;
  reportSeriesKey: string;
  templateKey: string;
  templateVersion: number;
  currentMonth: string;
  availableMonths: string[];
  createdAt: Date;
  updatedAt: Date;
}

export interface StoredReport {
  id: string;
  title: string;
  originalFilename: string;
  reportSeriesKey: string;
  templateKey: string;
  templateVersion: number;
  currentMonth: string;
  availableMonths: string[];
  createdAt: Date;
  updatedAt: Date;
  snapshot: NormalizedReportSnapshot;
  workbookObjectKey: string;
}

interface StoredReportInput {
  id: string;
  title: string;
  originalFilename: string;
  reportSeriesKey: string;
  templateKey: string;
  templateVersion: number;
  currentMonth: string;
  availableMonths: string[];
  createdAt: string;
  updatedAt: string;
  snapshot: NormalizedReportSnapshot;
  workbookObjectKey: string;
}

function normalizeSnapshot(snapshot: unknown): NormalizedReportSnapshot {
  const rawSnapshot = snapshot as Partial<NormalizedReportSnapshot>;

  return {
    ...rawSnapshot,
    periods: (rawSnapshot.periods ?? []).map((period) => ({
      ...period,
      reportCutOffDate: period.reportCutOffDate ?? period.monthEndDate ?? "",
    })),
    portfolioGanttWorkstreams: rawSnapshot.portfolioGanttWorkstreams ?? [],
    portfolioGanttMilestones: rawSnapshot.portfolioGanttMilestones ?? [],
    chartSettings: rawSnapshot.chartSettings ?? [],
  } as NormalizedReportSnapshot;
}

function sanitizeFilename(filename: string): string {
  return filename.replace(/[^a-zA-Z0-9._-]+/g, "-");
}

function slugifyLabel(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .replace(/-{2,}/g, "-");
}

function toJsonValue(value: unknown): Prisma.InputJsonValue {
  return JSON.parse(JSON.stringify(value)) as Prisma.InputJsonValue;
}

function createReportTitle(filename: string, snapshot: NormalizedReportSnapshot): string {
  const monthLabel = snapshot.currentMonth || snapshot.availableMonths.at(-1) || "report";
  const baseName = filename.replace(/\.[^.]+$/, "");
  return `${baseName} · ${monthLabel}`;
}

function toReportListItem(report: {
  id: string;
  title: string;
  originalFilename: string;
  reportSeriesKey?: string;
  templateKey: string;
  templateVersion: number;
  currentMonth: string;
  availableMonths: unknown;
  createdAt: Date;
  updatedAt: Date;
}): ReportListItem {
  return {
    id: report.id,
    title: report.title,
    originalFilename: report.originalFilename,
    reportSeriesKey: report.reportSeriesKey ?? deriveReportSeriesKey(report.originalFilename),
    templateKey: report.templateKey,
    templateVersion: report.templateVersion,
    currentMonth: report.currentMonth,
    availableMonths: report.availableMonths as string[],
    createdAt: report.createdAt,
    updatedAt: report.updatedAt,
  };
}

function toStoredReport(report: {
  id: string;
  title: string;
  originalFilename: string;
  reportSeriesKey?: string;
  templateKey: string;
  templateVersion: number;
  currentMonth: string;
  availableMonths: unknown;
  snapshot: unknown;
  createdAt: Date;
  updatedAt: Date;
  workbookObjectKey: string;
}): StoredReport {
  return {
    id: report.id,
    title: report.title,
    originalFilename: report.originalFilename,
    reportSeriesKey: report.reportSeriesKey ?? deriveReportSeriesKey(report.originalFilename),
    templateKey: report.templateKey,
    templateVersion: report.templateVersion,
    currentMonth: report.currentMonth,
    availableMonths: report.availableMonths as string[],
    snapshot: normalizeSnapshot(report.snapshot),
    createdAt: report.createdAt,
    updatedAt: report.updatedAt,
    workbookObjectKey: report.workbookObjectKey,
  };
}

function toStoredReportFromInput(report: StoredReportInput): StoredReport {
  return {
    id: report.id,
    title: report.title,
    originalFilename: report.originalFilename,
    reportSeriesKey: report.reportSeriesKey,
    templateKey: report.templateKey,
    templateVersion: report.templateVersion,
    currentMonth: report.currentMonth,
    availableMonths: report.availableMonths,
    snapshot: normalizeSnapshot(report.snapshot),
    createdAt: new Date(report.createdAt),
    updatedAt: new Date(report.updatedAt),
    workbookObjectKey: report.workbookObjectKey,
  };
}

function toReportListItemFromManifest(manifest: DraftManifest): ReportListItem {
  return {
    id: manifest.reportId,
    title: manifest.title,
    originalFilename: manifest.originalFilename,
    reportSeriesKey: manifest.reportSeriesKey,
    templateKey: manifest.templateKey,
    templateVersion: manifest.templateVersion,
    currentMonth: manifest.currentMonth,
    availableMonths: manifest.availableMonths,
    createdAt: new Date(manifest.createdAt),
    updatedAt: new Date(manifest.updatedAt),
  };
}

function toDraftUserRef(user: AppUser): DraftUserRef {
  return {
    id: user.id,
    name: user.name,
    email: user.email,
  };
}

function hydrateDerivedSnapshot(snapshot: NormalizedReportSnapshot): NormalizedReportSnapshot {
  const derivedNetworkMetrics = deriveNetworkMetrics(snapshot);
  const networkRows = buildDerivedNetworkServiceRows(derivedNetworkMetrics);

  return {
    ...snapshot,
    derivedNetworkMetrics,
    serviceAvailability: [
      ...snapshot.serviceAvailability.filter((row) => row.serviceName !== NETWORK_SERVICE_NAME),
      ...networkRows,
    ].sort((left, right) => left.reportingMonth.localeCompare(right.reportingMonth)),
  };
}

function createArtifactSyncStatus(workbookObjectKey: string | null): ArtifactSyncStatus {
  return {
    state: workbookObjectKey ? "current" : "pending",
    revisionId: null,
    updatedAt: null,
    error: null,
    workbookObjectKey,
    jsonObjectKey: null,
  };
}

async function getDraftStoredReport(id: string): Promise<StoredReport | null> {
  const manifest = await getDraftManifest(id);
  if (!manifest) {
    return null;
  }

  const snapshot = await getCurrentDraftSnapshot(id);
  if (!snapshot) {
    return null;
  }

  return toStoredReportFromInput({
    id: manifest.reportId,
    title: manifest.title,
    originalFilename: manifest.originalFilename,
    reportSeriesKey: manifest.reportSeriesKey,
    templateKey: manifest.templateKey,
    templateVersion: manifest.templateVersion,
    currentMonth: manifest.currentMonth,
    availableMonths: manifest.availableMonths,
    createdAt: manifest.createdAt,
    updatedAt: manifest.updatedAt,
    snapshot,
    workbookObjectKey: manifest.workbookObjectKey,
  });
}

async function upsertMirrorReportRecord(input: StoredReportInput): Promise<void> {
  await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      await prisma.report.upsert({
        where: { id: input.id },
        update: {
          title: input.title,
          originalFilename: input.originalFilename,
          templateKey: input.templateKey,
          templateVersion: input.templateVersion,
          currentMonth: input.currentMonth,
          availableMonths: toJsonValue(input.availableMonths),
          metadata: toJsonValue(input.snapshot.metadata),
          snapshot: toJsonValue(input.snapshot),
          workbookObjectKey: input.workbookObjectKey,
        },
        create: {
          id: input.id,
          title: input.title,
          originalFilename: input.originalFilename,
          templateKey: input.templateKey,
          templateVersion: input.templateVersion,
          validationStatus: "VALID",
          workbookObjectKey: input.workbookObjectKey,
          availableMonths: toJsonValue(input.availableMonths),
          currentMonth: input.currentMonth,
          metadata: toJsonValue(input.snapshot.metadata),
          snapshot: toJsonValue(input.snapshot),
        },
      });
    },
    async () => {
      await upsertLocalReport({
        id: input.id,
        title: input.title,
        originalFilename: input.originalFilename,
        reportSeriesKey: input.reportSeriesKey,
        templateKey: input.templateKey,
        templateVersion: input.templateVersion,
        currentMonth: input.currentMonth,
        availableMonths: input.availableMonths,
        snapshot: input.snapshot,
        workbookObjectKey: input.workbookObjectKey,
      });
    },
  );
}

function normalizePresence(records: EditorPresence[]): EditorPresence[] {
  const cutoff = Date.now() - 2 * 60 * 1000;
  return records
    .filter((record) => Date.parse(record.lastSeenAt) >= cutoff)
    .sort((left, right) => Date.parse(right.lastSeenAt) - Date.parse(left.lastSeenAt));
}

function buildRevisionMeta(input: {
  previousRevision: ReportRevisionMeta | null;
  changedSections: SectionId[];
  actor: DraftUserRef;
}): ReportRevisionMeta {
  return {
    revisionId: nanoid(),
    revisionNumber: (input.previousRevision?.revisionNumber ?? 0) + 1,
    updatedAt: new Date().toISOString(),
    updatedBy: input.actor,
    changedSections: input.changedSections,
  };
}

async function persistDraftSnapshot(input: {
  reportId: string;
  title: string;
  originalFilename: string;
  reportSeriesKey: string;
  templateKey: string;
  templateVersion: number;
  snapshot: NormalizedReportSnapshot;
  workbookObjectKey: string;
  actor: DraftUserRef;
  changedSections: SectionId[];
  createdAt?: string;
  artifactSyncStatus?: ArtifactSyncStatus;
}): Promise<StoredReport> {
  const previousManifest = await getDraftManifest(input.reportId);
  const normalizedSnapshot = hydrateDerivedSnapshot(normalizeSnapshot(input.snapshot));
  const revision = buildRevisionMeta({
    previousRevision: previousManifest?.currentRevision ?? null,
    changedSections: input.changedSections,
    actor: input.actor,
  });
  const artifactSyncStatus: ArtifactSyncStatus = {
    ...(input.artifactSyncStatus ?? previousManifest?.artifactSyncStatus ?? createArtifactSyncStatus(input.workbookObjectKey)),
    state: "pending",
    revisionId: revision.revisionId,
    updatedAt: revision.updatedAt,
    error: null,
    workbookObjectKey: input.workbookObjectKey,
  };

  const manifest: DraftManifest = {
    reportId: input.reportId,
    title: input.title,
    originalFilename: input.originalFilename,
    reportSeriesKey: input.reportSeriesKey,
    templateKey: input.templateKey,
    templateVersion: input.templateVersion,
    currentMonth: normalizedSnapshot.currentMonth,
    availableMonths: normalizedSnapshot.availableMonths,
    createdAt: previousManifest?.createdAt ?? input.createdAt ?? revision.updatedAt,
    updatedAt: revision.updatedAt,
    workbookObjectKey: input.workbookObjectKey,
    artifactSyncStatus,
    currentRevision: revision,
  };

  await Promise.all([
    saveDraftRevision(input.reportId, {
      meta: revision,
      snapshot: normalizedSnapshot,
    }),
    writeCurrentDraftSnapshot(input.reportId, normalizedSnapshot),
    upsertDraftManifest(manifest),
    upsertMirrorReportRecord({
      id: input.reportId,
      title: input.title,
      originalFilename: input.originalFilename,
      reportSeriesKey: input.reportSeriesKey,
      templateKey: input.templateKey,
      templateVersion: input.templateVersion,
      currentMonth: input.snapshot.currentMonth,
      availableMonths: input.snapshot.availableMonths,
      createdAt: previousManifest?.createdAt ?? input.createdAt ?? revision.updatedAt,
      updatedAt: revision.updatedAt,
      snapshot: normalizedSnapshot,
      workbookObjectKey: input.workbookObjectKey,
    }),
  ]);

  return toStoredReportFromInput({
    id: input.reportId,
    title: input.title,
    originalFilename: input.originalFilename,
    reportSeriesKey: input.reportSeriesKey,
    templateKey: input.templateKey,
    templateVersion: input.templateVersion,
    currentMonth: input.snapshot.currentMonth,
    availableMonths: input.snapshot.availableMonths,
    createdAt: previousManifest?.createdAt ?? input.createdAt ?? revision.updatedAt,
    updatedAt: revision.updatedAt,
    snapshot: normalizedSnapshot,
    workbookObjectKey: input.workbookObjectKey,
  });
}

export async function createBlankReportDraft(input: {
  title: string;
  initialMonth: string;
  actor?: AppUser;
}): Promise<StoredReport> {
  const reportId = nanoid();
  const reportSeriesKey = slugifyLabel(input.title);
  const snapshot = createBlankSnapshot(input.initialMonth, input.title);
  const actor = toDraftUserRef(
    input.actor ?? {
      id: "system-blank-draft",
      name: "Blank Draft",
      email: null,
    },
  );
  const originalFilename = `${reportSeriesKey || "report"}.xlsx`;
  const workbookObjectKey = path.posix.join("workbooks", reportId, "current.xlsx");

  const report = await persistDraftSnapshot({
    reportId,
    title: input.title,
    originalFilename,
    reportSeriesKey,
    templateKey: snapshot.metadata.templateKey,
    templateVersion: snapshot.metadata.templateVersion,
    snapshot,
    workbookObjectKey,
    actor,
    changedSections: [...EDITOR_SECTIONS],
  });

  void syncDraftArtifacts(report.id);
  return report;
}

export async function getEditableReportDraft(reportId: string, reportingMonth: string): Promise<EditableReportDraft> {
  const report = await getStoredReport(reportId);
  if (!report) {
    throw new Error("Report not found.");
  }

  if (!report.availableMonths.includes(reportingMonth)) {
    throw new Error("Invalid month.");
  }

  const manifest = await getDraftManifest(reportId);
  if (!manifest) {
    throw new Error("Editable draft not found.");
  }

  return {
    manifest,
    snapshot: report.snapshot,
    activePresence: normalizePresence(await listDraftPresence(reportId, reportingMonth)),
  };
}

export async function updateEditorPresence(reportId: string, reportingMonth: string, actor: AppUser): Promise<EditorPresence[]> {
  await upsertDraftPresence({
    reportId,
    reportingMonth,
    user: toDraftUserRef(actor),
    lastSeenAt: new Date().toISOString(),
  });

  return normalizePresence(await listDraftPresence(reportId, reportingMonth));
}

function getChangedSectionsSince(revisions: ReportRevisionMeta[], baseRevisionId: string | null): SectionId[] {
  if (!baseRevisionId) {
    return [];
  }

  const changed = new Set<SectionId>();
  for (const revision of revisions) {
    if (revision.revisionId === baseRevisionId) {
      break;
    }

    revision.changedSections.forEach((sectionId) => changed.add(sectionId));
  }

  return [...changed];
}

export async function saveEditorSection<S extends SectionId>(input: {
  reportId: string;
  reportingMonth: string;
  sectionId: S;
  payload: SectionPayloadMap[S];
  baseRevisionId: string | null;
  actor: AppUser;
}): Promise<EditableReportDraft> {
  const current = await getEditableReportDraft(input.reportId, input.reportingMonth);
  const revisions = await listDraftRevisions(input.reportId);
  const changedSectionsSinceBase = getChangedSectionsSince(revisions, input.baseRevisionId);

  if (input.baseRevisionId && changedSectionsSinceBase.includes(input.sectionId)) {
    throw new Error(`Conflict:${changedSectionsSinceBase.join(",")}`);
  }

  const applied = applySectionPayload(current.snapshot, input.sectionId, input.payload);
  const nextSnapshot = normalizeSnapshot(applied.snapshot);
  const nextTitle = applied.title ?? current.manifest.title;
  const nextReportSeriesKey = applied.reportSeriesKey ?? current.manifest.reportSeriesKey;

  await persistDraftSnapshot({
    reportId: input.reportId,
    title: nextTitle,
    originalFilename: current.manifest.originalFilename,
    reportSeriesKey: nextReportSeriesKey,
    templateKey: current.manifest.templateKey,
    templateVersion: current.manifest.templateVersion,
    snapshot: nextSnapshot,
    workbookObjectKey: current.manifest.workbookObjectKey,
    actor: toDraftUserRef(input.actor),
    changedSections: [input.sectionId],
    createdAt: current.manifest.createdAt,
    artifactSyncStatus: current.manifest.artifactSyncStatus,
  });

  await updateEditorPresence(input.reportId, input.reportingMonth, input.actor);
  void syncDraftArtifacts(input.reportId);
  return getEditableReportDraft(input.reportId, input.reportingMonth);
}

function isPersistenceFallbackError(error: unknown): boolean {
  if (!(error instanceof Error)) {
    return false;
  }

  return [
    "DATABASE_URL is required",
    "User was denied access on the database",
    "Can't reach database server",
    "Connection refused",
    "connection pool",
    "does not exist",
    "The table",
    "The column",
  ].some((message) => error.message.includes(message));
}

async function withLocalFallback<T>(action: () => Promise<T>, fallback: () => Promise<T>): Promise<T> {
  try {
    return await action();
  } catch (error) {
    if (!isPersistenceFallbackError(error)) {
      throw error;
    }

    logger.warn({ error }, "Database unavailable; falling back to local JSON report store");
    return fallback();
  }
}

export async function listReports(): Promise<ReportListItem[]> {
  const draftReports = (await listDraftManifests()).map(toReportListItemFromManifest);
  const legacyReports = await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const reports = await prisma.report.findMany({
        orderBy: { createdAt: "desc" },
        select: {
          id: true,
          title: true,
          originalFilename: true,
          templateKey: true,
          templateVersion: true,
          currentMonth: true,
          availableMonths: true,
          createdAt: true,
          updatedAt: true,
        },
      });

      return reports.map(toReportListItem);
    },
    async () =>
      (await listLocalReports()).map((report) => ({
        id: report.id,
        title: report.title,
        originalFilename: report.originalFilename,
        reportSeriesKey: report.reportSeriesKey ?? deriveReportSeriesKey(report.originalFilename),
        templateKey: report.templateKey,
        templateVersion: report.templateVersion,
        currentMonth: report.currentMonth,
        availableMonths: report.availableMonths,
        createdAt: new Date(report.createdAt),
        updatedAt: new Date(report.updatedAt),
      })),
  );

  const merged = [...draftReports, ...legacyReports.filter((report) => !draftReports.some((draft) => draft.id === report.id))];
  return merged.sort((left, right) => right.updatedAt.getTime() - left.updatedAt.getTime());
}

export async function getStoredReport(id: string): Promise<StoredReport | null> {
  const draftReport = await getDraftStoredReport(id);
  if (draftReport) {
    return draftReport;
  }

  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const report = await prisma.report.findUnique({
        where: { id },
        select: {
          id: true,
          title: true,
          originalFilename: true,
          templateKey: true,
          templateVersion: true,
          currentMonth: true,
          availableMonths: true,
          snapshot: true,
          createdAt: true,
          updatedAt: true,
          workbookObjectKey: true,
        },
      });

      return report ? toStoredReport(report) : null;
    },
    async () => {
      const report = await getLocalReport(id);
      return report
        ? {
            id: report.id,
            title: report.title,
            originalFilename: report.originalFilename,
            reportSeriesKey: report.reportSeriesKey ?? deriveReportSeriesKey(report.originalFilename),
            templateKey: report.templateKey,
            templateVersion: report.templateVersion,
            currentMonth: report.currentMonth,
            availableMonths: report.availableMonths,
            snapshot: normalizeSnapshot(report.snapshot),
            createdAt: new Date(report.createdAt),
            updatedAt: new Date(report.updatedAt),
            workbookObjectKey: report.workbookObjectKey,
          }
        : null;
    },
  );
}

export async function createReportFromWorkbookUpload(filename: string, buffer: Buffer, actor?: AppUser): Promise<StoredReport> {
  const parsed = await parseWorkbookBuffer(buffer, filename);
  const storage = getObjectStorage();
  const key = path.posix.join("workbooks", nanoid(), sanitizeFilename(filename));
  const reportSeriesKey = deriveReportSeriesKey(filename);

  await storage.putBuffer(key, buffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

  const title = createReportTitle(filename, parsed.snapshot);
  const report = await withLocalFallback<StoredReport>(
    async () => {
      const prisma = getPrisma();
      const created = await prisma.report.create({
        data: {
          title,
          originalFilename: filename,
          templateKey: parsed.snapshot.metadata.templateKey,
          templateVersion: parsed.snapshot.metadata.templateVersion,
          validationStatus: "VALID",
          workbookObjectKey: key,
          availableMonths: toJsonValue(parsed.snapshot.availableMonths),
          currentMonth: parsed.snapshot.currentMonth,
          metadata: toJsonValue(parsed.snapshot.metadata),
          snapshot: toJsonValue(parsed.snapshot),
        },
        select: {
          id: true,
          title: true,
          originalFilename: true,
          templateKey: true,
          templateVersion: true,
          currentMonth: true,
          availableMonths: true,
          snapshot: true,
          createdAt: true,
          updatedAt: true,
          workbookObjectKey: true,
        },
      });

      return toStoredReport(created);
    },
    async () => {
      const localReport = await createLocalReport({
        title,
        originalFilename: filename,
        reportSeriesKey,
        templateKey: parsed.snapshot.metadata.templateKey,
        templateVersion: parsed.snapshot.metadata.templateVersion,
        currentMonth: parsed.snapshot.currentMonth,
        availableMonths: parsed.snapshot.availableMonths,
        snapshot: parsed.snapshot,
        workbookObjectKey: key,
      });

      return {
        id: localReport.id,
        title: localReport.title,
        originalFilename: localReport.originalFilename,
        reportSeriesKey: localReport.reportSeriesKey ?? deriveReportSeriesKey(localReport.originalFilename),
        templateKey: localReport.templateKey,
        templateVersion: localReport.templateVersion,
        currentMonth: localReport.currentMonth,
        availableMonths: localReport.availableMonths,
        snapshot: normalizeSnapshot(localReport.snapshot),
        createdAt: new Date(localReport.createdAt),
        updatedAt: new Date(localReport.updatedAt),
        workbookObjectKey: localReport.workbookObjectKey,
      };
    },
  );

  const draftActor = toDraftUserRef(
    actor ?? {
      id: "system-workbook-import",
      name: "Workbook Import",
      email: null,
    },
  );

  await persistDraftSnapshot({
    reportId: report.id,
    title: report.title,
    originalFilename: report.originalFilename,
    reportSeriesKey: report.reportSeriesKey,
    templateKey: report.templateKey,
    templateVersion: report.templateVersion,
    snapshot: report.snapshot,
    workbookObjectKey: report.workbookObjectKey,
    actor: draftActor,
    changedSections: [...EDITOR_SECTIONS],
    createdAt: report.createdAt.toISOString(),
    artifactSyncStatus: {
      state: "current",
      revisionId: null,
      updatedAt: report.updatedAt.toISOString(),
      error: null,
      workbookObjectKey: report.workbookObjectKey,
      jsonObjectKey: null,
    },
  });

  void syncDraftArtifacts(report.id);
  logger.info({ reportId: report.id, filename }, "Stored workbook report");

  return (await getStoredReport(report.id)) ?? report;
}

async function updateDraftArtifactSyncStatus(
  reportId: string,
  updater: (status: ArtifactSyncStatus, manifest: DraftManifest) => ArtifactSyncStatus,
): Promise<void> {
  const manifest = await getDraftManifest(reportId);
  if (!manifest) {
    return;
  }

  await upsertDraftManifest({
    ...manifest,
    updatedAt: new Date().toISOString(),
    artifactSyncStatus: updater(manifest.artifactSyncStatus, manifest),
  });
}

export async function syncDraftArtifacts(reportId: string): Promise<void> {
  const report = await getStoredReport(reportId);
  const manifest = await getDraftManifest(reportId);
  if (!report || !manifest) {
    return;
  }

  try {
    const storage = getObjectStorage();
    const workbookBuffer = await renderWorkbookFromSnapshot(report.snapshot);
    const jsonBuffer = Buffer.from(`${JSON.stringify(report.snapshot, null, 2)}\n`, "utf8");
    const jsonObjectKey = path.posix.join("exports", reportId, "current.json");

    await Promise.all([
      storage.putBuffer(report.workbookObjectKey, workbookBuffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
      storage.putBuffer(jsonObjectKey, jsonBuffer, "application/json"),
    ]);

    await updateDraftArtifactSyncStatus(reportId, () => ({
      state: "current",
      revisionId: manifest.currentRevision.revisionId,
      updatedAt: new Date().toISOString(),
      error: null,
      workbookObjectKey: report.workbookObjectKey,
      jsonObjectKey,
    }));
  } catch (error) {
    await updateDraftArtifactSyncStatus(reportId, (status) => ({
      ...status,
      state: "failed",
      updatedAt: new Date().toISOString(),
      error: error instanceof Error ? error.message : "Artifact sync failed.",
    }));
  }
}

export async function getCurrentDraftJsonArtifact(reportId: string): Promise<{ buffer: Buffer; filename: string }> {
  const report = await getStoredReport(reportId);
  if (!report) {
    throw new Error("Report not found.");
  }

  const storage = getObjectStorage();
  const manifest = await getDraftManifest(reportId);
  const jsonObjectKey = manifest?.artifactSyncStatus.jsonObjectKey ?? path.posix.join("exports", reportId, "current.json");
  if (!manifest || manifest.artifactSyncStatus.state !== "current" || manifest.artifactSyncStatus.revisionId !== manifest.currentRevision.revisionId) {
    await syncDraftArtifacts(reportId);
  }

  const latestManifest = (await getDraftManifest(reportId)) ?? manifest;
  const objectKey = latestManifest?.artifactSyncStatus.jsonObjectKey ?? jsonObjectKey;
  const buffer = (await storage.exists(objectKey))
    ? await storage.getBuffer(objectKey)
    : Buffer.from(`${JSON.stringify(report.snapshot, null, 2)}\n`, "utf8");

  return {
    buffer,
    filename: `${slugifyLabel(report.title)}-${report.currentMonth}-current.json`,
  };
}

export async function getCurrentDraftWorkbookArtifact(reportId: string): Promise<{ buffer: Buffer; filename: string }> {
  const report = await getStoredReport(reportId);
  if (!report) {
    throw new Error("Report not found.");
  }

  const manifest = await getDraftManifest(reportId);
  if (!manifest || manifest.artifactSyncStatus.state !== "current" || manifest.artifactSyncStatus.revisionId !== manifest.currentRevision.revisionId) {
    await syncDraftArtifacts(reportId);
  }

  const storage = getObjectStorage();
  const buffer = await storage.getBuffer(report.workbookObjectKey);

  return {
    buffer,
    filename: `${slugifyLabel(report.title)}-${report.currentMonth}.xlsx`,
  };
}

export async function saveGeneratedExport(input: {
  reportId: string;
  exportType: string;
  month?: string;
  pageId?: string;
  blockId?: string;
  contentType: string;
  data: Buffer;
}): Promise<string> {
  const extension =
    input.contentType === "application/pdf"
      ? "pdf"
      : input.contentType === "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        ? "pptx"
        : "png";
  const key = path.posix.join("exports", input.reportId, `${input.exportType}-${nanoid()}.${extension}`);
  const storage = getObjectStorage();

  await storage.putBuffer(key, input.data, input.contentType);

  await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      await prisma.reportExport.create({
        data: {
          reportId: input.reportId,
          exportType: input.exportType,
          objectKey: key,
          contentType: input.contentType,
          month: input.month,
          pageId: input.pageId,
          blockId: input.blockId,
          metadata: toJsonValue({
            size: input.data.byteLength,
          }),
        },
      });
    },
    async () =>
      saveLocalExport({
        reportId: input.reportId,
        exportType: input.exportType,
        objectKey: key,
        contentType: input.contentType,
        month: input.month,
        pageId: input.pageId,
        blockId: input.blockId,
        metadata: {
          size: input.data.byteLength,
        },
      }),
  );

  return key;
}

let cachedDemoSnapshot: NormalizedReportSnapshot | null = null;

export async function getBundledDemoSnapshot(): Promise<NormalizedReportSnapshot> {
  if (cachedDemoSnapshot) {
    return cachedDemoSnapshot;
  }

  const workbookPath = path.resolve(process.cwd(), "fixtures", "IT_Exec_Reporting_Ingestion_Template_v4_dummy_data.xlsx");
  const workbookBuffer = await fs.readFile(workbookPath);
  const parsed = await parseWorkbookBuffer(workbookBuffer, path.basename(workbookPath));

  cachedDemoSnapshot = parsed.snapshot;
  return cachedDemoSnapshot;
}

async function getOwnedExecSummaryForMonth(reportId: string, reportingMonth: string): Promise<ExecSummaryState> {
  if (reportId === "demo") {
    return createDemoExecSummary(reportingMonth);
  }

  const report = await getStoredReport(reportId);
  if (!report) {
    throw new Error("Report not found.");
  }

  if (!report.availableMonths.includes(reportingMonth)) {
    throw new Error("Invalid month.");
  }

  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const explicit = await prisma.reportExecSummary.findUnique({
        where: {
          reportId_reportingMonth: {
            reportId,
            reportingMonth,
          },
        },
      });

      if (!explicit) {
        return {
          mode: "empty",
          contentHtml: "",
          excerpt: "",
          updatedAt: null,
          sourceReportId: null,
        } satisfies ExecSummaryState;
      }

      return {
        mode: "explicit",
        contentHtml: explicit.contentHtml,
        excerpt: explicit.excerpt,
        updatedAt: explicit.updatedAt.toISOString(),
        sourceReportId: explicit.sourceReportId ?? null,
      } satisfies ExecSummaryState;
    },
    async () => {
      const explicit = await getLocalExecSummary(reportId, reportingMonth);
      if (!explicit) {
        return {
          mode: "empty",
          contentHtml: "",
          excerpt: "",
          updatedAt: null,
          sourceReportId: null,
        } satisfies ExecSummaryState;
      }

      return {
        mode: "explicit",
        contentHtml: explicit.contentHtml,
        excerpt: explicit.excerpt,
        updatedAt: explicit.updatedAt,
        sourceReportId: explicit.sourceReportId ?? null,
      } satisfies ExecSummaryState;
    },
  );
}

export async function getExecSummaryState(reportId: string, reportingMonth: string): Promise<ExecSummaryState> {
  if (reportId === "demo") {
    return createDemoExecSummary(reportingMonth);
  }

  const report = await getStoredReport(reportId);
  if (!report) {
    throw new Error("Report not found.");
  }

  if (!report.availableMonths.includes(reportingMonth)) {
    throw new Error("Invalid month.");
  }

  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const explicit = await prisma.reportExecSummary.findUnique({
        where: {
          reportId_reportingMonth: {
            reportId,
            reportingMonth,
          },
        },
      });

      if (explicit) {
        return {
          mode: "explicit",
          contentHtml: explicit.contentHtml,
          excerpt: explicit.excerpt,
          updatedAt: explicit.updatedAt.toISOString(),
          sourceReportId: explicit.sourceReportId ?? null,
        } satisfies ExecSummaryState;
      }

      const carried = await prisma.reportExecSummary.findFirst({
        where: {
          reportSeriesKey: report.reportSeriesKey,
          reportingMonth,
          NOT: { reportId },
        },
        orderBy: { updatedAt: "desc" },
      });

      if (carried) {
        return {
          mode: "carried-forward",
          contentHtml: carried.contentHtml,
          excerpt: carried.excerpt,
          updatedAt: carried.updatedAt.toISOString(),
          sourceReportId: carried.reportId,
        } satisfies ExecSummaryState;
      }

      return {
        mode: "empty",
        contentHtml: "",
        excerpt: "",
        updatedAt: null,
        sourceReportId: null,
      } satisfies ExecSummaryState;
    },
    async () => {
      const explicit = await getLocalExecSummary(reportId, reportingMonth);
      if (explicit) {
        return {
          mode: "explicit",
          contentHtml: explicit.contentHtml,
          excerpt: explicit.excerpt,
          updatedAt: explicit.updatedAt,
          sourceReportId: explicit.sourceReportId ?? null,
        } satisfies ExecSummaryState;
      }

      const carried = await findLocalCarryForwardExecSummary({
        reportSeriesKey: report.reportSeriesKey,
        reportingMonth,
        excludeReportId: reportId,
      });

      if (carried) {
        return {
          mode: "carried-forward",
          contentHtml: carried.contentHtml,
          excerpt: carried.excerpt,
          updatedAt: carried.updatedAt,
          sourceReportId: carried.reportId,
        } satisfies ExecSummaryState;
      }

      return {
        mode: "empty",
        contentHtml: "",
        excerpt: "",
        updatedAt: null,
        sourceReportId: null,
      } satisfies ExecSummaryState;
    },
  );
}

export async function saveExecSummary(reportId: string, reportingMonth: string, rawContentHtml: string): Promise<ExecSummaryState> {
  if (reportId === "demo") {
    throw new Error("The bundled demo summary is read-only.");
  }

  const report = await getStoredReport(reportId);
  if (!report) {
    throw new Error("Report not found.");
  }

  if (!report.availableMonths.includes(reportingMonth)) {
    throw new Error("Invalid month.");
  }

  const contentHtml = sanitizeExecSummaryHtml(rawContentHtml);
  const excerpt = buildExecSummaryExcerpt(contentHtml);
  const existingState = await getExecSummaryState(reportId, reportingMonth);
  const sourceReportId = existingState.mode === "carried-forward" ? existingState.sourceReportId : null;

  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const summary = await prisma.reportExecSummary.upsert({
        where: {
          reportId_reportingMonth: {
            reportId,
            reportingMonth,
          },
        },
        update: {
          contentHtml,
          excerpt,
          sourceReportId,
        },
        create: {
          reportId,
          reportSeriesKey: report.reportSeriesKey,
          reportingMonth,
          contentHtml,
          excerpt,
          sourceReportId,
        },
      });

      return {
        mode: "explicit",
        contentHtml: summary.contentHtml,
        excerpt: summary.excerpt,
        updatedAt: summary.updatedAt.toISOString(),
        sourceReportId: summary.sourceReportId ?? null,
      } satisfies ExecSummaryState;
    },
    async () =>
      upsertLocalExecSummary({
        reportId,
        reportSeriesKey: report.reportSeriesKey,
        reportingMonth,
        contentHtml,
        excerpt,
        sourceReportId,
      }),
  );
}

export async function getReportPrepView(reportId: string, reportingMonth: string): Promise<ReportPrepView> {
  const mode = reportId === "demo" ? "demo-readonly" : "editable";
  const snapshot =
    reportId === "demo"
      ? await getBundledDemoSnapshot()
      : (await getStoredReport(reportId))?.snapshot;

  if (!snapshot) {
    throw new Error("Report not found.");
  }

  if (!snapshot.availableMonths.includes(reportingMonth)) {
    throw new Error("Invalid month.");
  }

  const [currentSummary, previousMonthState] = await Promise.all([
    getExecSummaryState(reportId, reportingMonth),
    (async () => {
      const currentIndex = snapshot.availableMonths.indexOf(reportingMonth);
      if (currentIndex <= 0) {
        return null;
      }

      const previousMonth = snapshot.availableMonths[currentIndex - 1];
      return getOwnedExecSummaryForMonth(reportId, previousMonth);
    })(),
  ]);

  if (reportId === "demo") {
    return buildReportPrepView({
      snapshot,
      reportingMonth,
      currentSummary,
      previousMonthSummary: previousMonthState,
      mode,
      acknowledgedCheckIds: [],
      updatedAt: null,
    });
  }

  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const prepState = await prisma.reportPrepState.findUnique({
        where: {
          reportId_reportingMonth: {
            reportId,
            reportingMonth,
          },
        },
      });

      return buildReportPrepView({
        snapshot,
        reportingMonth,
        currentSummary,
        previousMonthSummary: previousMonthState,
        mode,
        acknowledgedCheckIds: (prepState?.acknowledgedCheckIds as string[] | null | undefined) ?? [],
        updatedAt: prepState?.updatedAt.toISOString() ?? null,
      });
    },
    async () => {
      const prepState = await getLocalPrepState(reportId, reportingMonth);
      return buildReportPrepView({
        snapshot,
        reportingMonth,
        currentSummary,
        previousMonthSummary: previousMonthState,
        mode,
        acknowledgedCheckIds: prepState?.acknowledgedCheckIds ?? [],
        updatedAt: prepState?.updatedAt ?? null,
      });
    },
  );
}

export async function saveReportPrepAcknowledgements(
  reportId: string,
  reportingMonth: string,
  requestedCheckIds: string[],
): Promise<ReportPrepView> {
  if (reportId === "demo") {
    throw new Error("The bundled demo prep center is read-only.");
  }

  const report = await getStoredReport(reportId);
  if (!report) {
    throw new Error("Report not found.");
  }

  if (!report.availableMonths.includes(reportingMonth)) {
    throw new Error("Invalid month.");
  }

  const currentView = await getReportPrepView(reportId, reportingMonth);
  const acknowledgedCheckIds = filterAcknowledgeableCheckIds(
    [...currentView.readiness.checks, ...currentView.readiness.reviewedChecks],
    requestedCheckIds,
  );

  await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      await prisma.reportPrepState.upsert({
        where: {
          reportId_reportingMonth: {
            reportId,
            reportingMonth,
          },
        },
        update: {
          acknowledgedCheckIds: toJsonValue(acknowledgedCheckIds),
        },
        create: {
          reportId,
          reportingMonth,
          acknowledgedCheckIds: toJsonValue(acknowledgedCheckIds),
        },
      });
    },
    async () =>
      {
        await upsertLocalPrepState({
          reportId,
          reportingMonth,
          acknowledgedCheckIds,
        });
      },
  );

  return getReportPrepView(reportId, reportingMonth);
}
