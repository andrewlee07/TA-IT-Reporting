import { redirect } from "next/navigation";

import { ReportAppShell, type AppReportRecord } from "@/components/report-app-shell";
import { REPORT_PAGES, isValidPageId } from "@/lib/report/blocks";
import { loadTemplateBodyMarkup, loadTemplateStyles } from "@/lib/report/template-source";
import { getBundledDemoSnapshot, getStoredReport, listReports, type ReportListItem } from "@/lib/reports/service";

export const dynamic = "force-dynamic";

function getSingleValue(value: string | string[] | undefined): string | undefined {
  return Array.isArray(value) ? value[0] : value;
}

function normalizeListEntry(report: ReportListItem) {
  return {
    ...report,
    createdAt: report.createdAt.toISOString(),
    updatedAt: report.updatedAt.toISOString(),
  };
}

function resolvePage(page: string | undefined): string {
  return page && isValidPageId(page) ? page : REPORT_PAGES[0].id;
}

function resolveMonth(report: Pick<AppReportRecord, "availableMonths" | "currentMonth">, month: string | undefined): string {
  if (month && report.availableMonths.includes(month)) {
    return month;
  }

  return report.currentMonth;
}

function hasPortfolioGanttData(report: Pick<AppReportRecord, "templateVersion" | "snapshot">): boolean {
  return report.templateVersion >= 3 && report.snapshot.portfolioGanttWorkstreams.length > 0;
}

async function loadAppReport(id: string): Promise<AppReportRecord | null> {
  if (id === "demo") {
    const snapshot = await getBundledDemoSnapshot();
    return {
      id: "demo",
      title: "Bundled Demo Report",
      originalFilename: snapshot.metadata.sourceFilename,
      templateKey: snapshot.metadata.templateKey,
      templateVersion: snapshot.metadata.templateVersion,
      currentMonth: snapshot.currentMonth,
      availableMonths: snapshot.availableMonths,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      snapshot,
    };
  }

  const report = await getStoredReport(id);
  if (!report) {
    return null;
  }

  return {
    ...report,
    createdAt: report.createdAt.toISOString(),
    updatedAt: report.updatedAt.toISOString(),
  };
}

async function loadPreferredFallbackReport(reports: ReturnType<typeof normalizeListEntry>[]): Promise<AppReportRecord | null> {
  for (const report of reports) {
    if (report.templateVersion < 3) {
      continue;
    }

    const loadedReport = await loadAppReport(report.id);

    if (loadedReport && hasPortfolioGanttData(loadedReport)) {
      return loadedReport;
    }
  }

  return loadAppReport("demo");
}

interface HomePageProps {
  searchParams: Promise<{
    report?: string | string[];
    month?: string | string[];
    page?: string | string[];
  }>;
}

export default async function HomePage({ searchParams }: HomePageProps) {
  const query = await searchParams;
  const reports = await listReports();
  const normalizedReports = reports.map(normalizeListEntry);
  const requestedReportId = getSingleValue(query.report);
  const requestedMonth = getSingleValue(query.month);
  const requestedPageId = getSingleValue(query.page);

  let activeReport = requestedReportId
    ? await loadAppReport(requestedReportId)
    : await loadPreferredFallbackReport(normalizedReports);

  if (!activeReport) {
    activeReport = await loadPreferredFallbackReport(normalizedReports);
  }

  if (!activeReport) {
    throw new Error("Unable to load any report state.");
  }

  const canonicalReportId = activeReport.id;
  const canonicalMonth = resolveMonth(activeReport, requestedMonth);
  const canonicalPageId = resolvePage(requestedPageId);

  if (requestedReportId !== canonicalReportId || requestedMonth !== canonicalMonth || requestedPageId !== canonicalPageId) {
    redirect(`/?report=${canonicalReportId}&month=${canonicalMonth}&page=${canonicalPageId}`);
  }

  const [templateStyles, templateBody] = await Promise.all([loadTemplateStyles(), loadTemplateBodyMarkup()]);

  return (
    <>
      <style dangerouslySetInnerHTML={{ __html: templateStyles }} />
      <ReportAppShell
        initialMonth={canonicalMonth}
        initialPageId={canonicalPageId}
        initialReport={activeReport}
        initialReports={normalizedReports}
        templateBody={templateBody}
      />
    </>
  );
}
