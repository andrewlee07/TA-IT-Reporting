import { promises as fs } from "node:fs";
import path from "node:path";

import { nanoid } from "nanoid";

import type { ReportAnnotation } from "@/lib/annotations/types";
import { getEnv } from "@/lib/env";
import { deriveReportSeriesKey, type ExecSummaryState } from "@/lib/reports/exec-summary";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

interface LocalReportRecord {
  id: string;
  title: string;
  originalFilename: string;
  reportSeriesKey?: string;
  templateKey: string;
  templateVersion: number;
  currentMonth: string;
  availableMonths: string[];
  createdAt: string;
  updatedAt: string;
  snapshot: NormalizedReportSnapshot;
  workbookObjectKey: string;
}

interface LocalExportRecord {
  id: string;
  reportId: string;
  exportType: string;
  objectKey: string;
  contentType: string;
  month?: string;
  pageId?: string;
  blockId?: string;
  metadata?: Record<string, unknown>;
  createdAt: string;
}

interface LocalExecSummaryRecord {
  id: string;
  reportId: string;
  reportSeriesKey: string;
  reportingMonth: string;
  contentHtml: string;
  excerpt: string;
  sourceReportId?: string | null;
  createdAt: string;
  updatedAt: string;
}

interface LocalPrepStateRecord {
  id: string;
  reportId: string;
  reportingMonth: string;
  acknowledgedCheckIds: string[];
  createdAt: string;
  updatedAt: string;
}

interface LocalAnnotationStateRecord {
  id: string;
  reportId: string;
  reportingMonth: string;
  revisionId: string;
  annotations: ReportAnnotation[];
  createdAt: string;
  updatedAt: string;
}

function getRootDir(): string {
  return path.resolve(process.cwd(), getEnv().LOCAL_STORAGE_DIR, "report-store");
}

function getReportsPath(): string {
  return path.join(getRootDir(), "reports.json");
}

function getExportsPath(): string {
  return path.join(getRootDir(), "exports.json");
}

function getExecSummariesPath(): string {
  return path.join(getRootDir(), "exec-summaries.json");
}

function getPrepStatesPath(): string {
  return path.join(getRootDir(), "prep-states.json");
}

function getAnnotationsPath(): string {
  return path.join(getRootDir(), "annotations.json");
}

async function ensureDir(): Promise<void> {
  await fs.mkdir(getRootDir(), { recursive: true });
}

async function readJsonFile<T>(filePath: string, fallback: T): Promise<T> {
  try {
    const raw = await fs.readFile(filePath, "utf8");
    return JSON.parse(raw) as T;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") {
      return fallback;
    }

    throw error;
  }
}

async function writeJsonFile(filePath: string, value: unknown): Promise<void> {
  await ensureDir();
  const tempPath = `${filePath}.${process.pid}.${Date.now()}.${Math.random().toString(36).slice(2)}.tmp`;
  await fs.writeFile(tempPath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
  await fs.rename(tempPath, filePath);
}

export async function listLocalReports(): Promise<LocalReportRecord[]> {
  const reports = await readJsonFile<LocalReportRecord[]>(getReportsPath(), []);
  return reports
    .map((report) => ({
      ...report,
      reportSeriesKey: report.reportSeriesKey ?? deriveReportSeriesKey(report.originalFilename),
    }))
    .sort((left, right) => right.createdAt.localeCompare(left.createdAt));
}

export async function getLocalReport(id: string): Promise<LocalReportRecord | null> {
  const reports = await listLocalReports();
  return reports.find((report) => report.id === id) ?? null;
}

export async function createLocalReport(input: {
  id?: string;
  title: string;
  originalFilename: string;
  reportSeriesKey?: string;
  templateKey: string;
  templateVersion: number;
  currentMonth: string;
  availableMonths: string[];
  snapshot: NormalizedReportSnapshot;
  workbookObjectKey: string;
}): Promise<LocalReportRecord> {
  const reports = await listLocalReports();
  const now = new Date().toISOString();

  const report: LocalReportRecord = {
    id: input.id ?? nanoid(),
    title: input.title,
    originalFilename: input.originalFilename,
    reportSeriesKey: input.reportSeriesKey ?? deriveReportSeriesKey(input.originalFilename),
    templateKey: input.templateKey,
    templateVersion: input.templateVersion,
    currentMonth: input.currentMonth,
    availableMonths: input.availableMonths,
    createdAt: now,
    updatedAt: now,
    snapshot: input.snapshot,
    workbookObjectKey: input.workbookObjectKey,
  };

  await writeJsonFile(getReportsPath(), [report, ...reports]);
  return report;
}

export async function upsertLocalReport(input: {
  id: string;
  title: string;
  originalFilename: string;
  reportSeriesKey?: string;
  templateKey: string;
  templateVersion: number;
  currentMonth: string;
  availableMonths: string[];
  snapshot: NormalizedReportSnapshot;
  workbookObjectKey: string;
}): Promise<LocalReportRecord> {
  const reports = await listLocalReports();
  const existing = reports.find((report) => report.id === input.id);
  const now = new Date().toISOString();

  const record: LocalReportRecord = {
    id: input.id,
    title: input.title,
    originalFilename: input.originalFilename,
    reportSeriesKey: input.reportSeriesKey ?? deriveReportSeriesKey(input.originalFilename),
    templateKey: input.templateKey,
    templateVersion: input.templateVersion,
    currentMonth: input.currentMonth,
    availableMonths: input.availableMonths,
    createdAt: existing?.createdAt ?? now,
    updatedAt: now,
    snapshot: input.snapshot,
    workbookObjectKey: input.workbookObjectKey,
  };

  const filtered = reports.filter((report) => report.id !== input.id);
  await writeJsonFile(getReportsPath(), [record, ...filtered]);
  return record;
}

export async function listLocalExecSummaries(): Promise<LocalExecSummaryRecord[]> {
  const summaries = await readJsonFile<LocalExecSummaryRecord[]>(getExecSummariesPath(), []);
  return summaries.sort((left, right) => right.updatedAt.localeCompare(left.updatedAt));
}

export async function getLocalExecSummary(reportId: string, reportingMonth: string): Promise<LocalExecSummaryRecord | null> {
  const summaries = await listLocalExecSummaries();
  return summaries.find((summary) => summary.reportId === reportId && summary.reportingMonth === reportingMonth) ?? null;
}

export async function findLocalCarryForwardExecSummary(input: {
  reportSeriesKey: string;
  reportingMonth: string;
  excludeReportId: string;
}): Promise<LocalExecSummaryRecord | null> {
  const summaries = await listLocalExecSummaries();
  return (
    summaries.find(
      (summary) =>
        summary.reportSeriesKey === input.reportSeriesKey &&
        summary.reportingMonth === input.reportingMonth &&
        summary.reportId !== input.excludeReportId,
    ) ?? null
  );
}

export async function upsertLocalExecSummary(input: {
  reportId: string;
  reportSeriesKey: string;
  reportingMonth: string;
  contentHtml: string;
  excerpt: string;
  sourceReportId?: string | null;
}): Promise<ExecSummaryState> {
  const summaries = await listLocalExecSummaries();
  const existing = summaries.find((summary) => summary.reportId === input.reportId && summary.reportingMonth === input.reportingMonth);
  const now = new Date().toISOString();

  const nextRecord: LocalExecSummaryRecord = existing
    ? {
        ...existing,
        reportSeriesKey: input.reportSeriesKey,
        contentHtml: input.contentHtml,
        excerpt: input.excerpt,
        sourceReportId: input.sourceReportId ?? existing.sourceReportId ?? null,
        updatedAt: now,
      }
    : {
        id: nanoid(),
        reportId: input.reportId,
        reportSeriesKey: input.reportSeriesKey,
        reportingMonth: input.reportingMonth,
        contentHtml: input.contentHtml,
        excerpt: input.excerpt,
        sourceReportId: input.sourceReportId ?? null,
        createdAt: now,
        updatedAt: now,
      };

  const filtered = summaries.filter((summary) => !(summary.reportId === input.reportId && summary.reportingMonth === input.reportingMonth));
  await writeJsonFile(getExecSummariesPath(), [nextRecord, ...filtered]);

  return {
    mode: "explicit",
    contentHtml: nextRecord.contentHtml,
    excerpt: nextRecord.excerpt,
    updatedAt: nextRecord.updatedAt,
    sourceReportId: nextRecord.sourceReportId ?? null,
  };
}

export async function getLocalPrepState(reportId: string, reportingMonth: string): Promise<LocalPrepStateRecord | null> {
  const states = await readJsonFile<LocalPrepStateRecord[]>(getPrepStatesPath(), []);
  return states.find((state) => state.reportId === reportId && state.reportingMonth === reportingMonth) ?? null;
}

export async function upsertLocalPrepState(input: {
  reportId: string;
  reportingMonth: string;
  acknowledgedCheckIds: string[];
}): Promise<LocalPrepStateRecord> {
  const states = await readJsonFile<LocalPrepStateRecord[]>(getPrepStatesPath(), []);
  const existing = states.find((state) => state.reportId === input.reportId && state.reportingMonth === input.reportingMonth);
  const now = new Date().toISOString();

  const nextRecord: LocalPrepStateRecord = existing
    ? {
        ...existing,
        acknowledgedCheckIds: input.acknowledgedCheckIds,
        updatedAt: now,
      }
    : {
        id: nanoid(),
        reportId: input.reportId,
        reportingMonth: input.reportingMonth,
        acknowledgedCheckIds: input.acknowledgedCheckIds,
        createdAt: now,
        updatedAt: now,
      };

  const filtered = states.filter((state) => !(state.reportId === input.reportId && state.reportingMonth === input.reportingMonth));
  await writeJsonFile(getPrepStatesPath(), [nextRecord, ...filtered]);
  return nextRecord;
}

export async function getLocalAnnotationState(reportId: string, reportingMonth: string): Promise<LocalAnnotationStateRecord | null> {
  const states = await readJsonFile<LocalAnnotationStateRecord[]>(getAnnotationsPath(), []);
  return states.find((state) => state.reportId === reportId && state.reportingMonth === reportingMonth) ?? null;
}

export async function upsertLocalAnnotationState(input: {
  reportId: string;
  reportingMonth: string;
  revisionId: string;
  annotations: ReportAnnotation[];
}): Promise<LocalAnnotationStateRecord> {
  const states = await readJsonFile<LocalAnnotationStateRecord[]>(getAnnotationsPath(), []);
  const existing = states.find((state) => state.reportId === input.reportId && state.reportingMonth === input.reportingMonth);
  const now = new Date().toISOString();

  const nextRecord: LocalAnnotationStateRecord = existing
    ? {
        ...existing,
        revisionId: input.revisionId,
        annotations: input.annotations,
        updatedAt: now,
      }
    : {
        id: nanoid(),
        reportId: input.reportId,
        reportingMonth: input.reportingMonth,
        revisionId: input.revisionId,
        annotations: input.annotations,
        createdAt: now,
        updatedAt: now,
      };

  const filtered = states.filter((state) => !(state.reportId === input.reportId && state.reportingMonth === input.reportingMonth));
  await writeJsonFile(getAnnotationsPath(), [nextRecord, ...filtered]);
  return nextRecord;
}

export async function saveLocalExport(input: {
  reportId: string;
  exportType: string;
  objectKey: string;
  contentType: string;
  month?: string;
  pageId?: string;
  blockId?: string;
  metadata?: Record<string, unknown>;
}): Promise<void> {
  const exports = await readJsonFile<LocalExportRecord[]>(getExportsPath(), []);
  const record: LocalExportRecord = {
    id: nanoid(),
    reportId: input.reportId,
    exportType: input.exportType,
    objectKey: input.objectKey,
    contentType: input.contentType,
    month: input.month,
    pageId: input.pageId,
    blockId: input.blockId,
    metadata: input.metadata,
    createdAt: new Date().toISOString(),
  };

  await writeJsonFile(getExportsPath(), [record, ...exports]);
}
