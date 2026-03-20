import { promises as fs } from "node:fs";
import path from "node:path";

import { nanoid } from "nanoid";
import { Prisma } from "@/generated/prisma/client";

import { getPrisma } from "@/lib/prisma";
import { getObjectStorage } from "@/lib/storage";
import { logger } from "@/lib/logger";
import { createLocalReport, getLocalReport, listLocalReports, saveLocalExport } from "@/lib/reports/local-report-store";
import { parseWorkbookBuffer } from "@/lib/workbook/parser";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

export interface ReportListItem {
  id: string;
  title: string;
  originalFilename: string;
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
  templateKey: string;
  templateVersion: number;
  currentMonth: string;
  availableMonths: string[];
  createdAt: Date;
  updatedAt: Date;
  snapshot: NormalizedReportSnapshot;
  workbookObjectKey: string;
}

function sanitizeFilename(filename: string): string {
  return filename.replace(/[^a-zA-Z0-9._-]+/g, "-");
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
    templateKey: report.templateKey,
    templateVersion: report.templateVersion,
    currentMonth: report.currentMonth,
    availableMonths: report.availableMonths as string[],
    snapshot: report.snapshot as NormalizedReportSnapshot,
    createdAt: report.createdAt,
    updatedAt: report.updatedAt,
    workbookObjectKey: report.workbookObjectKey,
  };
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
  return withLocalFallback(
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
        ...report,
        createdAt: new Date(report.createdAt),
        updatedAt: new Date(report.updatedAt),
      })),
  );
}

export async function getStoredReport(id: string): Promise<StoredReport | null> {
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
            ...report,
            createdAt: new Date(report.createdAt),
            updatedAt: new Date(report.updatedAt),
          }
        : null;
    },
  );
}

export async function createReportFromWorkbookUpload(filename: string, buffer: Buffer): Promise<StoredReport> {
  const parsed = await parseWorkbookBuffer(buffer, filename);
  const storage = getObjectStorage();
  const key = path.posix.join("workbooks", nanoid(), sanitizeFilename(filename));

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
        templateKey: parsed.snapshot.metadata.templateKey,
        templateVersion: parsed.snapshot.metadata.templateVersion,
        currentMonth: parsed.snapshot.currentMonth,
        availableMonths: parsed.snapshot.availableMonths,
        snapshot: parsed.snapshot,
        workbookObjectKey: key,
      });

      return {
        ...localReport,
        createdAt: new Date(localReport.createdAt),
        updatedAt: new Date(localReport.updatedAt),
      };
    },
  );

  logger.info({ reportId: report.id, filename }, "Stored workbook report");

  return report;
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
  const extension = input.contentType === "application/pdf" ? "pdf" : "png";
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

  const workbookPath = path.resolve(process.cwd(), "fixtures", "IT_Exec_Reporting_Ingestion_Template_v2_dummy_data.xlsx");
  const workbookBuffer = await fs.readFile(workbookPath);
  const parsed = await parseWorkbookBuffer(workbookBuffer, path.basename(workbookPath));

  cachedDemoSnapshot = parsed.snapshot;
  return cachedDemoSnapshot;
}
