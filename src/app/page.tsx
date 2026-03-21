import { redirect } from "next/navigation";

import { ReportAppShell, type AppReportRecord } from "@/components/report-app-shell";
import { REPORT_PAGES, hasPageTabs, isValidPageId, resolveTabId } from "@/lib/report/blocks";
import { loadTemplateBodyMarkup, loadTemplateStyles } from "@/lib/report/template-source";
import { getBundledDemoSnapshot, getExecSummaryState, getStoredReport, listReports, type ReportListItem } from "@/lib/reports/service";

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

function resolveTab(pageId: string, tab: string | undefined): string | null {
  return resolveTabId(pageId, tab);
}

function buildCanonicalUrl(reportId: string, month: string, pageId: string, tabId: string | null): string {
  const params = new URLSearchParams();
  params.set("report", reportId);
  params.set("month", month);
  params.set("page", pageId);

  if (tabId) {
    params.set("tab", tabId);
  }

  return `/?${params.toString()}`;
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
      reportSeriesKey: "bundled-demo-report",
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
    tab?: string | string[];
  }>;
}

export default async function HomePage({ searchParams }: HomePageProps) {
  const query = await searchParams;
  const reports = await listReports();
  const normalizedReports = reports.map(normalizeListEntry);
  const requestedReportId = getSingleValue(query.report);
  const requestedMonth = getSingleValue(query.month);
  const requestedPageId = getSingleValue(query.page);
  const requestedTabId = getSingleValue(query.tab);

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
  const canonicalTabId = resolveTab(canonicalPageId, requestedTabId);
  const shouldRedirectForTab =
    hasPageTabs(canonicalPageId) ? requestedTabId !== canonicalTabId : typeof requestedTabId !== "undefined";

  if (
    requestedReportId !== canonicalReportId ||
    requestedMonth !== canonicalMonth ||
    requestedPageId !== canonicalPageId ||
    shouldRedirectForTab
  ) {
    redirect(buildCanonicalUrl(canonicalReportId, canonicalMonth, canonicalPageId, canonicalTabId));
  }

  const [templateStyles, templateBody, initialExecSummary] = await Promise.all([
    loadTemplateStyles(),
    loadTemplateBodyMarkup(),
    getExecSummaryState(canonicalReportId, canonicalMonth),
  ]);

  return (
    <>
      <style dangerouslySetInnerHTML={{ __html: templateStyles }} />
      <ReportAppShell
        initialMonth={canonicalMonth}
        initialPageId={canonicalPageId}
        initialTabId={canonicalTabId}
        initialReport={activeReport}
        initialReports={normalizedReports}
        initialExecSummary={initialExecSummary}
        templateBody={templateBody}
      />
    </>
  );
}
