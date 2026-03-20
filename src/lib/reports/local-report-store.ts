import { promises as fs } from "node:fs";
import path from "node:path";

import { nanoid } from "nanoid";

import { getEnv } from "@/lib/env";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

interface LocalReportRecord {
  id: string;
  title: string;
  originalFilename: string;
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

function getRootDir(): string {
  return path.resolve(process.cwd(), getEnv().LOCAL_STORAGE_DIR, "report-store");
}

function getReportsPath(): string {
  return path.join(getRootDir(), "reports.json");
}

function getExportsPath(): string {
  return path.join(getRootDir(), "exports.json");
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
  await fs.writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

export async function listLocalReports(): Promise<LocalReportRecord[]> {
  const reports = await readJsonFile<LocalReportRecord[]>(getReportsPath(), []);
  return reports.sort((left, right) => right.createdAt.localeCompare(left.createdAt));
}

export async function getLocalReport(id: string): Promise<LocalReportRecord | null> {
  const reports = await listLocalReports();
  return reports.find((report) => report.id === id) ?? null;
}

export async function createLocalReport(input: {
  title: string;
  originalFilename: string;
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
    id: nanoid(),
    title: input.title,
    originalFilename: input.originalFilename,
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
