import { promises as fs } from "node:fs";
import path from "node:path";

import { nanoid } from "nanoid";
import { Prisma } from "@/generated/prisma/client";

import { getPrisma } from "@/lib/prisma";
import { getObjectStorage } from "@/lib/storage";
import { logger } from "@/lib/logger";
import {
  createLocalReport,
  findLocalCarryForwardExecSummary,
  getLocalExecSummary,
  getLocalReport,
  listLocalReports,
  saveLocalExport,
  upsertLocalExecSummary,
} from "@/lib/reports/local-report-store";
import {
  buildExecSummaryExcerpt,
  createDemoExecSummary,
  deriveReportSeriesKey,
  sanitizeExecSummaryHtml,
  type ExecSummaryState,
} from "@/lib/reports/exec-summary";
import { parseWorkbookBuffer } from "@/lib/workbook/parser";
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

export async function createReportFromWorkbookUpload(filename: string, buffer: Buffer): Promise<StoredReport> {
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
