import { format, parseISO } from "date-fns";

import type { ExecSummaryState } from "@/lib/reports/exec-summary";
import type {
  AssetsLifecycleRow,
  BudgetCommercialRow,
  ChangeReleaseRow,
  DerivedNetworkMetricRow,
  DevDeliveryRow,
  NarrativeNoteRow,
  NormalizedReportSnapshot,
  OfficeNetworkAvailabilityRow,
  OldestTicketRow,
  PortfolioGanttMilestoneRow,
  PortfolioGanttWorkstreamRow,
  ProjectPortfolioRow,
  SecurityPatchingRow,
  ServiceAvailabilityRow,
  SupportOperationsRow,
  TopRiskRow,
} from "@/lib/workbook/types";

export type PrepMode = "editable" | "demo-readonly";
export type PrepCheckSeverity = "blocking" | "warning" | "info";
export type PrepSectionState = "changed" | "unchanged" | "missing" | "not-applicable";
export type PrepNarrativeState = "changed" | "unchanged" | "missing" | "not-applicable";

export interface ReportPrepCheck {
  id: string;
  severity: PrepCheckSeverity;
  source: string;
  title: string;
  reason: string;
  pageId: string;
  tabId: string | null;
  acknowledged: boolean;
  acknowledgeable: boolean;
}

export interface ReportPrepSummary {
  status: "ready" | "needs-attention";
  blockingCount: number;
  warningCount: number;
  infoCount: number;
  acknowledgedCount: number;
}

export interface ReportPrepMetricDelta {
  label: string;
  currentValue: string;
  previousValue: string;
  deltaLabel: string;
  changed: boolean;
}

export interface ReportPrepSectionRollup {
  id: string;
  label: string;
  pageId: string;
  tabId: string | null;
  state: PrepSectionState;
  summary: string;
  narrativeState: PrepNarrativeState;
  narrativeLabel: string;
  metrics: ReportPrepMetricDelta[];
}

export interface ReportPrepReviewItem {
  id: string;
  label: string;
  reason: string;
  pageId: string;
  tabId: string | null;
  priority: "high" | "medium" | "low";
}

export interface ReportPrepPreviousSummary {
  month: string;
  monthLabel: string;
  mode: ExecSummaryState["mode"];
  contentHtml: string;
  excerpt: string;
  updatedAt: string | null;
  available: boolean;
}

export interface ReportPrepView {
  mode: PrepMode;
  reportingMonth: string;
  reportingMonthLabel: string;
  previousMonth: string | null;
  previousMonthLabel: string | null;
  updatedAt: string | null;
  acknowledgedCheckIds: string[];
  readiness: {
    summary: ReportPrepSummary;
    checks: ReportPrepCheck[];
    reviewedChecks: ReportPrepCheck[];
  };
  rollover: {
    previousMonth: string | null;
    previousMonthLabel: string | null;
    sections: ReportPrepSectionRollup[];
    reviewQueue: ReportPrepReviewItem[];
    previousExecSummary: ReportPrepPreviousSummary | null;
  };
}

interface NumericMetricSpec {
  label: string;
  kind: "integer" | "percent" | "score" | "currency" | "days";
  current: number | null;
  previous: number | null;
}

interface SectionInput {
  id: string;
  label: string;
  pageId: string;
  tabId: string | null;
  metrics: NumericMetricSpec[];
  currentNarrative: string;
  previousNarrative: string;
  missingMessage: string;
  notApplicable?: boolean;
}

function formatMonthLabel(month: string): string {
  return format(parseISO(`${month}-01`), "MMMM yyyy");
}

function hasMeaningfulText(value: string | null | undefined): boolean {
  return Boolean(value && value.replace(/[\s.-]+/g, "").trim().length > 0);
}

function normalizeText(value: string | null | undefined): string {
  return (value ?? "").replace(/\s+/g, " ").trim().toLowerCase();
}

function formatNumber(value: number): string {
  return Number.isInteger(value) ? value.toLocaleString() : value.toLocaleString(undefined, { maximumFractionDigits: 1 });
}

function formatMetricValue(kind: NumericMetricSpec["kind"], value: number | null): string {
  if (value === null) {
    return "—";
  }

  switch (kind) {
    case "percent":
      return `${value.toFixed(1)}%`;
    case "score":
      return `${value.toFixed(1)}/5`;
    case "currency":
      return `£${formatNumber(value)}`;
    case "days":
      return `${value.toFixed(1)} days`;
    default:
      return formatNumber(value);
  }
}

function formatMetricDelta(kind: NumericMetricSpec["kind"], current: number | null, previous: number | null): string {
  if (current === null && previous === null) {
    return "No data";
  }

  if (current !== null && previous === null) {
    return "New this month";
  }

  if (current === null && previous !== null) {
    return "Missing this month";
  }

  const delta = (current ?? 0) - (previous ?? 0);
  if (Math.abs(delta) < 0.0001) {
    return "No change";
  }

  const sign = delta > 0 ? "+" : "";
  switch (kind) {
    case "percent":
      return `${sign}${delta.toFixed(1)} pts`;
    case "score":
      return `${sign}${delta.toFixed(1)}`;
    case "currency":
      return `${sign}£${formatNumber(delta)}`;
    case "days":
      return `${sign}${delta.toFixed(1)} days`;
    default:
      return `${sign}${formatNumber(delta)}`;
  }
}

function metricChanged(current: number | null, previous: number | null): boolean {
  if (current === null && previous === null) {
    return false;
  }

  if (current === null || previous === null) {
    return true;
  }

  return Math.abs(current - previous) >= 0.0001;
}

function getPreviousMonth(snapshot: NormalizedReportSnapshot, reportingMonth: string): string | null {
  const currentIndex = snapshot.availableMonths.indexOf(reportingMonth);
  if (currentIndex <= 0) {
    return null;
  }

  return snapshot.availableMonths[currentIndex - 1] ?? null;
}

function getMonthRow<T extends { reportingMonth: string }>(rows: T[], reportingMonth: string): T | null {
  return rows.find((row) => row.reportingMonth === reportingMonth) ?? null;
}

function getMonthRows<T extends { reportingMonth: string }>(rows: T[], reportingMonth: string): T[] {
  return rows.filter((row) => row.reportingMonth === reportingMonth);
}

function getSupportRow(snapshot: NormalizedReportSnapshot, reportingMonth: string): SupportOperationsRow | null {
  return getMonthRow(snapshot.supportOperations, reportingMonth);
}

function getServiceRows(snapshot: NormalizedReportSnapshot, reportingMonth: string): ServiceAvailabilityRow[] {
  return getMonthRows(snapshot.serviceAvailability, reportingMonth);
}

function getOfficeRows(snapshot: NormalizedReportSnapshot, reportingMonth: string): OfficeNetworkAvailabilityRow[] {
  return getMonthRows(snapshot.officeNetworkAvailability, reportingMonth);
}

function getSecurityRow(snapshot: NormalizedReportSnapshot, reportingMonth: string): SecurityPatchingRow | null {
  return getMonthRow(snapshot.securityPatching, reportingMonth);
}

function getAssetsRows(snapshot: NormalizedReportSnapshot, reportingMonth: string): AssetsLifecycleRow[] {
  return getMonthRows(snapshot.assetsLifecycle, reportingMonth);
}

function getChangeRow(snapshot: NormalizedReportSnapshot, reportingMonth: string): ChangeReleaseRow | null {
  return getMonthRow(snapshot.changeRelease, reportingMonth);
}

function getDevRow(snapshot: NormalizedReportSnapshot, reportingMonth: string): DevDeliveryRow | null {
  return getMonthRow(snapshot.devDelivery, reportingMonth);
}

function getProjectRows(snapshot: NormalizedReportSnapshot, reportingMonth: string): ProjectPortfolioRow[] {
  return getMonthRows(snapshot.projectPortfolio, reportingMonth);
}

function getRiskRows(snapshot: NormalizedReportSnapshot, reportingMonth: string): TopRiskRow[] {
  return getMonthRows(snapshot.topRisks, reportingMonth);
}

function getBudgetRows(snapshot: NormalizedReportSnapshot, reportingMonth: string): BudgetCommercialRow[] {
  return getMonthRows(snapshot.budgetCommercials, reportingMonth);
}

function getTicketRows(snapshot: NormalizedReportSnapshot, reportingMonth: string): OldestTicketRow[] {
  return getMonthRows(snapshot.oldestTickets, reportingMonth);
}

function getGanttWorkstreams(snapshot: NormalizedReportSnapshot, reportingMonth: string): PortfolioGanttWorkstreamRow[] {
  return getMonthRows(snapshot.portfolioGanttWorkstreams, reportingMonth).filter((row) => row.inScope);
}

function getGanttMilestones(snapshot: NormalizedReportSnapshot, reportingMonth: string): PortfolioGanttMilestoneRow[] {
  return getMonthRows(snapshot.portfolioGanttMilestones, reportingMonth);
}

function getDerivedNetworkRow(snapshot: NormalizedReportSnapshot, reportingMonth: string): DerivedNetworkMetricRow | null {
  return getMonthRow(snapshot.derivedNetworkMetrics, reportingMonth);
}

function getNarrativeRows(snapshot: NormalizedReportSnapshot, reportingMonth: string, section: string): NarrativeNoteRow[] {
  return snapshot.narrativeNotes.filter((row) => row.reportingMonth === reportingMonth && row.section === section);
}

function textSignature(values: string[]): string {
  return values.map(normalizeText).filter(Boolean).sort().join("|");
}

function notesSignature(rows: NarrativeNoteRow[]): string {
  return JSON.stringify(
    rows
      .map((row) => ({
        type: row.noteType,
        headline: normalizeText(row.headline),
        narrative: normalizeText(row.narrative),
        owner: normalizeText(row.owner),
      }))
      .sort((left, right) =>
        `${left.type}|${left.headline}|${left.owner}`.localeCompare(`${right.type}|${right.headline}|${right.owner}`),
      ),
  );
}

function riskSignature(rows: TopRiskRow[]): string {
  return JSON.stringify(
    rows
      .map((row) => ({
        riskIssue: normalizeText(row.riskIssue),
        type: normalizeText(row.type),
        owner: normalizeText(row.owner),
        impact: normalizeText(row.impact),
        likelihood: normalizeText(row.likelihood),
        ratingRag: normalizeText(row.ratingRag),
        targetDate: row.targetDate,
        decisionRequired: row.decisionRequired,
      }))
      .sort((left, right) => left.riskIssue.localeCompare(right.riskIssue)),
  );
}

function ticketSignature(rows: OldestTicketRow[]): string {
  return JSON.stringify(
    rows
      .map((row) => ({
        ticketId: normalizeText(row.ticketId),
        title: normalizeText(row.title),
        ageDays: row.ageDays,
        ownerQueue: normalizeText(row.ownerQueue),
        targetResolutionDate: row.targetResolutionDate,
      }))
      .sort((left, right) => left.ticketId.localeCompare(right.ticketId)),
  );
}

function ganttSignature(rows: PortfolioGanttWorkstreamRow[]): string {
  return JSON.stringify(
    rows
      .map((row) => ({
        workstreamName: normalizeText(row.workstreamName),
        sponsorOwner: normalizeText(row.sponsorOwner),
        statusRag: normalizeText(row.statusRag),
        startDate: row.startDate,
        endDate: row.endDate,
        progressDate: row.progressDate,
      }))
      .sort((left, right) => left.workstreamName.localeCompare(right.workstreamName)),
  );
}

function budgetTotals(rows: BudgetCommercialRow[]): { budget: number; actual: number; forecast: number; variance: number } {
  return rows.reduce(
    (totals, row) => ({
      budget: totals.budget + row.budgetAmount,
      actual: totals.actual + row.actualAmount,
      forecast: totals.forecast + row.forecastAmount,
      variance: totals.variance + row.variance,
    }),
    { budget: 0, actual: 0, forecast: 0, variance: 0 },
  );
}

function buildBudgetSignature(rows: BudgetCommercialRow[]): string {
  return JSON.stringify(budgetTotals(rows));
}

function buildMetricDeltas(metrics: NumericMetricSpec[]): ReportPrepMetricDelta[] {
  return metrics.map((metric) => ({
    label: metric.label,
    currentValue: formatMetricValue(metric.kind, metric.current),
    previousValue: formatMetricValue(metric.kind, metric.previous),
    deltaLabel: formatMetricDelta(metric.kind, metric.current, metric.previous),
    changed: metricChanged(metric.current, metric.previous),
  }));
}

function buildSectionRollup(input: SectionInput, previousMonthExists: boolean): ReportPrepSectionRollup {
  if (input.notApplicable) {
    return {
      id: input.id,
      label: input.label,
      pageId: input.pageId,
      tabId: input.tabId,
      state: "not-applicable",
      summary: "This section is not month-scoped in the workbook.",
      narrativeState: "not-applicable",
      narrativeLabel: "Not applicable",
      metrics: [],
    };
  }

  const metrics = buildMetricDeltas(input.metrics);
  const changedMetricLabels = metrics.filter((metric) => metric.changed).map((metric) => metric.label);
  const currentNarrativePresent = hasMeaningfulText(input.currentNarrative);
  const previousNarrativePresent = hasMeaningfulText(input.previousNarrative);
  const hasCurrentData = input.metrics.some((metric) => metric.current !== null) || currentNarrativePresent;
  const hasPreviousData = input.metrics.some((metric) => metric.previous !== null) || previousNarrativePresent;

  let narrativeState: PrepNarrativeState = "not-applicable";
  let narrativeLabel = "Not applicable";

  if (currentNarrativePresent || previousNarrativePresent) {
    if (!currentNarrativePresent) {
      narrativeState = "missing";
      narrativeLabel = "Missing this month";
    } else if (!previousMonthExists || !previousNarrativePresent) {
      narrativeState = "changed";
      narrativeLabel = previousMonthExists ? "New this month" : "Present";
    } else if (normalizeText(input.currentNarrative) === normalizeText(input.previousNarrative)) {
      narrativeState = "unchanged";
      narrativeLabel = "Unchanged";
    } else {
      narrativeState = "changed";
      narrativeLabel = "Updated";
    }
  }

  if (!hasCurrentData) {
    return {
      id: input.id,
      label: input.label,
      pageId: input.pageId,
      tabId: input.tabId,
      state: "missing",
      summary: input.missingMessage,
      narrativeState,
      narrativeLabel,
      metrics,
    };
  }

  if (!previousMonthExists || !hasPreviousData) {
    return {
      id: input.id,
      label: input.label,
      pageId: input.pageId,
      tabId: input.tabId,
      state: "changed",
      summary: previousMonthExists ? "New month data is available but there is no prior comparison baseline." : "No earlier month is available in this workbook.",
      narrativeState,
      narrativeLabel,
      metrics,
    };
  }

  if (changedMetricLabels.length > 0) {
    return {
      id: input.id,
      label: input.label,
      pageId: input.pageId,
      tabId: input.tabId,
      state: "changed",
      summary: `${changedMetricLabels.slice(0, 2).join(" · ")} changed versus the previous month.`,
      narrativeState,
      narrativeLabel,
      metrics,
    };
  }

  if (narrativeState === "changed" || narrativeState === "missing") {
    return {
      id: input.id,
      label: input.label,
      pageId: input.pageId,
      tabId: input.tabId,
      state: "changed",
      summary: narrativeState === "missing" ? "Narrative commentary is missing for this month." : "Narrative commentary changed versus the previous month.",
      narrativeState,
      narrativeLabel,
      metrics,
    };
  }

  return {
    id: input.id,
    label: input.label,
    pageId: input.pageId,
    tabId: input.tabId,
    state: "unchanged",
    summary: "No material KPI or narrative changes were detected.",
    narrativeState,
    narrativeLabel,
    metrics,
  };
}

function createCheck(
  id: string,
  severity: PrepCheckSeverity,
  source: string,
  title: string,
  reason: string,
  pageId: string,
  tabId: string | null,
  acknowledgedIds: Set<string>,
): ReportPrepCheck {
  return {
    id,
    severity,
    source,
    title,
    reason,
    pageId,
    tabId,
    acknowledged: severity === "warning" ? acknowledgedIds.has(id) : false,
    acknowledgeable: severity === "warning",
  };
}

function buildReadinessChecks(
  snapshot: NormalizedReportSnapshot,
  reportingMonth: string,
  previousMonth: string | null,
  execSummary: ExecSummaryState,
  acknowledgedIds: Set<string>,
): ReportPrepCheck[] {
  const checks: ReportPrepCheck[] = [];

  const serviceRows = getServiceRows(snapshot, reportingMonth);
  const officeRows = getOfficeRows(snapshot, reportingMonth);
  const supportRow = getSupportRow(snapshot, reportingMonth);
  const securityRow = getSecurityRow(snapshot, reportingMonth);
  const assetRows = getAssetsRows(snapshot, reportingMonth);
  const changeRow = getChangeRow(snapshot, reportingMonth);
  const devRow = getDevRow(snapshot, reportingMonth);
  const projectRows = getProjectRows(snapshot, reportingMonth);
  const budgetRows = getBudgetRows(snapshot, reportingMonth);
  const riskRows = getRiskRows(snapshot, reportingMonth);
  const ganttRows = getGanttWorkstreams(snapshot, reportingMonth);
  const execNotes = getNarrativeRows(snapshot, reportingMonth, "Executive scorecard");

  if (execSummary.mode === "empty") {
    checks.push(
      createCheck(
        "summary-missing",
        "blocking",
        "Exec Summary",
        "Exec summary is missing",
        `Add a leadership narrative for ${formatMonthLabel(reportingMonth)} before export.`,
        "p-summary",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (execSummary.mode === "carried-forward") {
    checks.push(
      createCheck(
        "summary-carried-forward",
        "warning",
        "Exec Summary",
        "Exec summary is still carried forward",
        "Review and save the inherited summary so this month has its own owned narrative.",
        "p-summary",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (serviceRows.length === 0) {
    checks.push(
      createCheck(
        "data-service-missing",
        "blocking",
        "Service Availability",
        "No service availability data for the active month",
        `The workbook has no service availability rows for ${formatMonthLabel(reportingMonth)}.`,
        "p-avail",
        "overview",
        acknowledgedIds,
      ),
    );
  }

  if (officeRows.length === 0) {
    checks.push(
      createCheck(
        "data-network-missing",
        "blocking",
        "Network & Offices",
        "No office network data for the active month",
        `The workbook has no office availability rows for ${formatMonthLabel(reportingMonth)}.`,
        "p-network",
        "map",
        acknowledgedIds,
      ),
    );
  }

  if (!supportRow) {
    checks.push(
      createCheck(
        "data-support-missing",
        "blocking",
        "Support Operations",
        "No support operations row for the active month",
        `The workbook has no support operations row for ${formatMonthLabel(reportingMonth)}.`,
        "p-support",
        "overview",
        acknowledgedIds,
      ),
    );
  }

  if (!securityRow) {
    checks.push(
      createCheck(
        "data-security-missing",
        "blocking",
        "Security & Patching",
        "No security row for the active month",
        `The workbook has no security/patching row for ${formatMonthLabel(reportingMonth)}.`,
        "p-security",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (assetRows.length === 0) {
    checks.push(
      createCheck(
        "data-assets-missing",
        "blocking",
        "Assets & Lifecycle",
        "No asset lifecycle rows for the active month",
        `The workbook has no asset lifecycle rows for ${formatMonthLabel(reportingMonth)}.`,
        "p-assets",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (!changeRow) {
    checks.push(
      createCheck(
        "data-change-missing",
        "blocking",
        "Change & Release",
        "No change and release row for the active month",
        `The workbook has no change/release row for ${formatMonthLabel(reportingMonth)}.`,
        "p-change",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (!devRow) {
    checks.push(
      createCheck(
        "data-dev-missing",
        "blocking",
        "Development & Delivery",
        "No development delivery row for the active month",
        `The workbook has no development delivery row for ${formatMonthLabel(reportingMonth)}.`,
        "p-dev",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (projectRows.length === 0) {
    checks.push(
      createCheck(
        "data-projects-missing",
        "blocking",
        "Project Portfolio",
        "No project portfolio rows for the active month",
        `The workbook has no project portfolio rows for ${formatMonthLabel(reportingMonth)}.`,
        "p-projects",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (budgetRows.length === 0) {
    checks.push(
      createCheck(
        "data-budget-missing",
        "blocking",
        "Budget & Commercials",
        "No budget rows for the active month",
        `The workbook has no budget/commercial rows for ${formatMonthLabel(reportingMonth)}.`,
        "p-budget",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (riskRows.length === 0) {
    checks.push(
      createCheck(
        "data-risks-missing",
        "blocking",
        "Risks & Decisions",
        "No risks or decisions are recorded for the active month",
        `The workbook has no risk register rows for ${formatMonthLabel(reportingMonth)}.`,
        "p-risks",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (snapshot.metadata.templateVersion >= 4 && ganttRows.length === 0) {
    checks.push(
      createCheck(
        "data-gantt-missing",
        "blocking",
        "Portfolio Gantt",
        "No in-scope Gantt workstreams for the active month",
        `The workbook has no in-scope Gantt workstreams for ${formatMonthLabel(reportingMonth)}.`,
        "p-gantt",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (execNotes.length === 0) {
    checks.push(
      createCheck(
        "note-exec-highlights-missing",
        "warning",
        "Executive Scorecard",
        "No executive highlights were provided",
        "The workbook has no narrative note cards for the executive highlights slide.",
        "p-exec",
        "highlights",
        acknowledgedIds,
      ),
    );
  }

  if (securityRow && !hasMeaningfulText(securityRow.commentary)) {
    checks.push(
      createCheck(
        "note-security-missing",
        "warning",
        "Security & Patching",
        "Security commentary is empty",
        "Add workbook commentary so the security note feels authored rather than purely numeric.",
        "p-security",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (devRow && !hasMeaningfulText(devRow.commentary)) {
    checks.push(
      createCheck(
        "note-dev-missing",
        "warning",
        "Development & Delivery",
        "Development commentary is empty",
        "Add workbook commentary so the delivery note explains the trend behind the numbers.",
        "p-dev",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (projectRows.length > 0 && !projectRows.some((row) => hasMeaningfulText(row.commentary))) {
    checks.push(
      createCheck(
        "note-projects-missing",
        "warning",
        "Project Portfolio",
        "Project commentary is empty",
        "Add workbook commentary on at least one active project so the portfolio slide has clearer narrative context.",
        "p-projects",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (budgetRows.length > 0 && !budgetRows.some((row) => hasMeaningfulText(row.commentary))) {
    checks.push(
      createCheck(
        "note-budget-missing",
        "warning",
        "Budget & Commercials",
        "Budget commentary is empty",
        "Add workbook commentary so the budget slide explains the main drivers behind the totals.",
        "p-budget",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (riskRows.length > 0 && !riskRows.some((row) => hasMeaningfulText(row.commentary))) {
    checks.push(
      createCheck(
        "note-risks-missing",
        "warning",
        "Risks & Decisions",
        "Risk commentary is empty",
        "Add workbook commentary so the governance note explains what leadership should do next.",
        "p-risks",
        null,
        acknowledgedIds,
      ),
    );
  }

  if (previousMonth) {
    const previousRiskRows = getRiskRows(snapshot, previousMonth);
    const previousTicketRows = getTicketRows(snapshot, previousMonth);
    const previousGanttRows = getGanttWorkstreams(snapshot, previousMonth);
    const previousBudgetRows = getBudgetRows(snapshot, previousMonth);
    const currentTicketRows = getTicketRows(snapshot, reportingMonth);

    if (riskRows.length > 0 && previousRiskRows.length > 0 && riskSignature(riskRows) === riskSignature(previousRiskRows)) {
      checks.push(
        createCheck(
          "risks-unchanged",
          "warning",
          "Risks & Decisions",
          "Risk register is unchanged versus last month",
          "Confirm the register was reviewed and not simply carried over unchanged.",
          "p-risks",
          null,
          acknowledgedIds,
        ),
      );
    }

    if (currentTicketRows.length > 0 && previousTicketRows.length > 0 && ticketSignature(currentTicketRows) === ticketSignature(previousTicketRows)) {
      checks.push(
        createCheck(
          "oldest-tickets-unchanged",
          "warning",
          "Support Operations",
          "Oldest open ticket list is unchanged versus last month",
          "Confirm the service desk ageing list was refreshed and the current open-ticket picture is still accurate.",
          "p-support",
          "detail",
          acknowledgedIds,
        ),
      );
    }

    if (ganttRows.length > 0 && previousGanttRows.length > 0 && ganttSignature(ganttRows) === ganttSignature(previousGanttRows)) {
      checks.push(
        createCheck(
          "gantt-unchanged",
          "warning",
          "Portfolio Gantt",
          "Gantt workstreams and dates are unchanged versus last month",
          "Confirm the rolling portfolio view was updated and key dates still reflect the current plan.",
          "p-gantt",
          null,
          acknowledgedIds,
        ),
      );
    }

    if (budgetRows.length > 0 && previousBudgetRows.length > 0 && buildBudgetSignature(budgetRows) === buildBudgetSignature(previousBudgetRows)) {
      checks.push(
        createCheck(
          "budget-totals-unchanged",
          "warning",
          "Budget & Commercials",
          "Budget totals are unchanged versus last month",
          "Confirm the current month totals were refreshed rather than copied forward unchanged.",
          "p-budget",
          null,
          acknowledgedIds,
        ),
      );
    }
  } else {
    checks.push(
      createCheck(
        "rollover-no-previous-month",
        "info",
        "Rollover",
        "No earlier month is available in this workbook",
        "Rollover comparisons will appear once the workbook contains at least two reporting months.",
        "p-summary",
        null,
        acknowledgedIds,
      ),
    );
  }

  return checks;
}

function buildReviewQueue(
  sections: ReportPrepSectionRollup[],
  checks: ReportPrepCheck[],
  previousSummary: ReportPrepPreviousSummary | null,
): ReportPrepReviewItem[] {
  const items: ReportPrepReviewItem[] = [];
  const seen = new Set<string>();

  checks
    .filter((check) => check.severity !== "info" && !check.acknowledged)
    .forEach((check) => {
      const key = `${check.pageId}:${check.tabId ?? "default"}`;
      if (seen.has(key)) {
        return;
      }
      seen.add(key);
      items.push({
        id: `check:${check.id}`,
        label: check.source,
        reason: check.title,
        pageId: check.pageId,
        tabId: check.tabId,
        priority: check.severity === "blocking" ? "high" : "medium",
      });
    });

  sections.forEach((section) => {
    const key = `${section.pageId}:${section.tabId ?? "default"}`;
    if (seen.has(key) || section.state === "unchanged" || section.state === "not-applicable") {
      return;
    }

    let reason = section.summary;
    let priority: ReportPrepReviewItem["priority"] = section.state === "missing" ? "high" : "medium";

    if (section.narrativeState === "missing") {
      reason = "Narrative commentary is missing for this month.";
      priority = "medium";
    }

    seen.add(key);
    items.push({
      id: `section:${section.id}`,
      label: section.label,
      reason,
      pageId: section.pageId,
      tabId: section.tabId,
      priority,
    });
  });

  if (previousSummary && previousSummary.available && previousSummary.mode !== "empty" && !seen.has("p-summary:default")) {
    items.unshift({
      id: "summary-copy-helper",
      label: "Exec Summary",
      reason: `A previous-month summary exists for ${previousSummary.monthLabel} and can be used as a draft starting point.`,
      pageId: "p-summary",
      tabId: null,
      priority: "low",
    });
  }

  return items.slice(0, 8);
}

function buildSectionRollups(
  snapshot: NormalizedReportSnapshot,
  reportingMonth: string,
  previousMonth: string | null,
  currentSummary: ExecSummaryState,
  previousSummary: ExecSummaryState | null,
): ReportPrepSectionRollup[] {
  const supportCurrent = getSupportRow(snapshot, reportingMonth);
  const supportPrevious = previousMonth ? getSupportRow(snapshot, previousMonth) : null;
  const securityCurrent = getSecurityRow(snapshot, reportingMonth);
  const securityPrevious = previousMonth ? getSecurityRow(snapshot, previousMonth) : null;
  const changeCurrent = getChangeRow(snapshot, reportingMonth);
  const changePrevious = previousMonth ? getChangeRow(snapshot, previousMonth) : null;
  const devCurrent = getDevRow(snapshot, reportingMonth);
  const devPrevious = previousMonth ? getDevRow(snapshot, previousMonth) : null;
  const projectsCurrent = getProjectRows(snapshot, reportingMonth);
  const projectsPrevious = previousMonth ? getProjectRows(snapshot, previousMonth) : [];
  const budgetCurrent = getBudgetRows(snapshot, reportingMonth);
  const budgetPrevious = previousMonth ? getBudgetRows(snapshot, previousMonth) : [];
  const risksCurrent = getRiskRows(snapshot, reportingMonth);
  const risksPrevious = previousMonth ? getRiskRows(snapshot, previousMonth) : [];
  const assetsCurrent = getAssetsRows(snapshot, reportingMonth);
  const assetsPrevious = previousMonth ? getAssetsRows(snapshot, previousMonth) : [];
  const serviceCurrent = getServiceRows(snapshot, reportingMonth);
  const servicePrevious = previousMonth ? getServiceRows(snapshot, previousMonth) : [];
  const networkCurrent = getDerivedNetworkRow(snapshot, reportingMonth);
  const networkPrevious = previousMonth ? getDerivedNetworkRow(snapshot, previousMonth) : null;
  const ganttCurrent = getGanttWorkstreams(snapshot, reportingMonth);
  const ganttPrevious = previousMonth ? getGanttWorkstreams(snapshot, previousMonth) : [];
  const ganttMilestonesCurrent = getGanttMilestones(snapshot, reportingMonth);
  const ganttMilestonesPrevious = previousMonth ? getGanttMilestones(snapshot, previousMonth) : [];
  const execNotesCurrent = getNarrativeRows(snapshot, reportingMonth, "Executive scorecard");
  const execNotesPrevious = previousMonth ? getNarrativeRows(snapshot, previousMonth, "Executive scorecard") : [];

  const findServiceMetric = (rows: ServiceAvailabilityRow[], serviceName: string, metric: keyof ServiceAvailabilityRow) =>
    rows.find((row) => row.serviceName === serviceName)?.[metric] ?? null;

  const totalOutage = (rows: ServiceAvailabilityRow[]) => rows.reduce((sum, row) => sum + row.outageMinutes, 0);
  const totalActiveDevices = (rows: AssetsLifecycleRow[]) => rows.reduce((sum, row) => sum + row.activeDevices, 0);
  const totalRefreshSpend = (rows: AssetsLifecycleRow[]) => rows.reduce((sum, row) => sum + row.refreshSpend, 0);
  const totalAssetIncidents = (rows: AssetsLifecycleRow[]) => rows.reduce((sum, row) => sum + row.incidentsLinkedToAgedKit, 0);
  const findAsset = (rows: AssetsLifecycleRow[], type: string) => rows.find((row) => row.assetType === type) ?? null;
  const averageConfidence = (rows: ProjectPortfolioRow[]) =>
    rows.length ? rows.reduce((sum, row) => sum + row.deliveryConfidencePct, 0) / rows.length : null;
  const decisionsNeeded = (rows: ProjectPortfolioRow[]) => rows.filter((row) => row.decisionNeeded).length;
  const amberRisks = (rows: TopRiskRow[]) => rows.filter((row) => row.ratingRag === "Amber").length;
  const onTrack = (rows: PortfolioGanttWorkstreamRow[]) => rows.filter((row) => row.statusRag === "Green").length;
  const atRisk = (rows: PortfolioGanttWorkstreamRow[]) => rows.filter((row) => row.statusRag === "Amber").length;

  const sections: ReportPrepSectionRollup[] = [
    buildSectionRollup(
      {
        id: "exec-summary",
        label: "Exec Summary",
        pageId: "p-summary",
        tabId: null,
        metrics: [],
        currentNarrative: currentSummary.contentHtml,
        previousNarrative: previousSummary?.contentHtml ?? "",
        missingMessage: "No exec summary has been saved for this month.",
      },
      Boolean(previousMonth),
    ),
    buildSectionRollup(
      {
        id: "executive-scorecard",
        label: "Executive Scorecard",
        pageId: "p-exec",
        tabId: "overview",
        metrics: [
          {
            label: "Support SLA",
            kind: "percent",
            current: supportCurrent?.resolutionSlaPct ?? null,
            previous: supportPrevious?.resolutionSlaPct ?? null,
          },
          {
            label: "User CSAT",
            kind: "score",
            current: supportCurrent?.ticketCsatScore ?? null,
            previous: supportPrevious?.ticketCsatScore ?? null,
          },
          {
            label: "Critical Vulns",
            kind: "integer",
            current: securityCurrent?.criticalVulns ?? null,
            previous: securityPrevious?.criticalVulns ?? null,
          },
          {
            label: "Change Success",
            kind: "percent",
            current: changeCurrent?.changeSuccessRatePct ?? null,
            previous: changePrevious?.changeSuccessRatePct ?? null,
          },
          {
            label: "Dev Backlog",
            kind: "integer",
            current: devCurrent?.devBacklogEnd ?? null,
            previous: devPrevious?.devBacklogEnd ?? null,
          },
        ],
        currentNarrative: notesSignature(execNotesCurrent),
        previousNarrative: notesSignature(execNotesPrevious),
        missingMessage: `No executive scorecard data is available for ${formatMonthLabel(reportingMonth)}.`,
      },
      Boolean(previousMonth),
    ),
    buildSectionRollup(
      {
        id: "service-availability",
        label: "Service Availability",
        pageId: "p-avail",
        tabId: "detail",
        metrics: [
          {
            label: "Network Availability",
            kind: "percent",
            current: (findServiceMetric(serviceCurrent, "Network", "availabilityPct") as number | null) ?? null,
            previous: (findServiceMetric(servicePrevious, "Network", "availabilityPct") as number | null) ?? null,
          },
          {
            label: "Private Cloud",
            kind: "percent",
            current: (findServiceMetric(serviceCurrent, "Private Cloud", "availabilityPct") as number | null) ?? null,
            previous: (findServiceMetric(servicePrevious, "Private Cloud", "availabilityPct") as number | null) ?? null,
          },
          {
            label: "Outage Minutes",
            kind: "integer",
            current: serviceCurrent.length ? totalOutage(serviceCurrent) : null,
            previous: servicePrevious.length ? totalOutage(servicePrevious) : null,
          },
        ],
        currentNarrative: "",
        previousNarrative: "",
        missingMessage: `No service availability rows are available for ${formatMonthLabel(reportingMonth)}.`,
      },
      Boolean(previousMonth),
    ),
    buildSectionRollup(
      {
        id: "network-offices",
        label: "Network & Offices",
        pageId: "p-network",
        tabId: "map",
        metrics: [
          {
            label: "Average Availability",
            kind: "percent",
            current: networkCurrent?.availabilityPct ?? null,
            previous: networkPrevious?.availabilityPct ?? null,
          },
          {
            label: "Offices Below 99%",
            kind: "integer",
            current: networkCurrent?.below99Offices ?? null,
            previous: networkPrevious?.below99Offices ?? null,
          },
          {
            label: "Worst Office Availability",
            kind: "percent",
            current: networkCurrent?.worstAvailabilityPct ?? null,
            previous: networkPrevious?.worstAvailabilityPct ?? null,
          },
        ],
        currentNarrative: "",
        previousNarrative: "",
        missingMessage: `No office availability rows are available for ${formatMonthLabel(reportingMonth)}.`,
      },
      Boolean(previousMonth),
    ),
    buildSectionRollup(
      {
        id: "support-operations",
        label: "Support Operations",
        pageId: "p-support",
        tabId: "overview",
        metrics: [
          {
            label: "Tickets Opened",
            kind: "integer",
            current: supportCurrent?.ticketsOpened ?? null,
            previous: supportPrevious?.ticketsOpened ?? null,
          },
          {
            label: "Tickets Closed",
            kind: "integer",
            current: supportCurrent?.ticketsClosed ?? null,
            previous: supportPrevious?.ticketsClosed ?? null,
          },
          {
            label: "Backlog End",
            kind: "integer",
            current: supportCurrent?.backlogEnd ?? null,
            previous: supportPrevious?.backlogEnd ?? null,
          },
          {
            label: "Resolution SLA",
            kind: "percent",
            current: supportCurrent?.resolutionSlaPct ?? null,
            previous: supportPrevious?.resolutionSlaPct ?? null,
          },
        ],
        currentNarrative: supportCurrent?.commentary ?? "",
        previousNarrative: supportPrevious?.commentary ?? "",
        missingMessage: `No support operations row is available for ${formatMonthLabel(reportingMonth)}.`,
      },
      Boolean(previousMonth),
    ),
    buildSectionRollup(
      {
        id: "security-patching",
        label: "Security & Patching",
        pageId: "p-security",
        tabId: null,
        metrics: [
          {
            label: "Critical Vulns",
            kind: "integer",
            current: securityCurrent?.criticalVulns ?? null,
            previous: securityPrevious?.criticalVulns ?? null,
          },
          {
            label: "Workstation Patch",
            kind: "percent",
            current: securityCurrent?.workstationPatchCompliancePct ?? null,
            previous: securityPrevious?.workstationPatchCompliancePct ?? null,
          },
          {
            label: "MFA Coverage",
            kind: "percent",
            current: securityCurrent?.mfaCoveragePct ?? null,
            previous: securityPrevious?.mfaCoveragePct ?? null,
          },
          {
            label: "Overdue Remediation",
            kind: "integer",
            current: securityCurrent?.overdueRemediationItems ?? null,
            previous: securityPrevious?.overdueRemediationItems ?? null,
          },
        ],
        currentNarrative: securityCurrent?.commentary ?? "",
        previousNarrative: securityPrevious?.commentary ?? "",
        missingMessage: `No security/patching row is available for ${formatMonthLabel(reportingMonth)}.`,
      },
      Boolean(previousMonth),
    ),
    buildSectionRollup(
      {
        id: "assets-lifecycle",
        label: "Assets & Lifecycle",
        pageId: "p-assets",
        tabId: null,
        metrics: [
          {
            label: "Active Devices",
            kind: "integer",
            current: assetsCurrent.length ? totalActiveDevices(assetsCurrent) : null,
            previous: assetsPrevious.length ? totalActiveDevices(assetsPrevious) : null,
          },
          {
            label: "Laptop In Lifecycle",
            kind: "percent",
            current: findAsset(assetsCurrent, "Laptop")?.withinLifecyclePct ?? null,
            previous: findAsset(assetsPrevious, "Laptop")?.withinLifecyclePct ?? null,
          },
          {
            label: "Refresh Spend",
            kind: "currency",
            current: assetsCurrent.length ? totalRefreshSpend(assetsCurrent) : null,
            previous: assetsPrevious.length ? totalRefreshSpend(assetsPrevious) : null,
          },
          {
            label: "Incidents Linked",
            kind: "integer",
            current: assetsCurrent.length ? totalAssetIncidents(assetsCurrent) : null,
            previous: assetsPrevious.length ? totalAssetIncidents(assetsPrevious) : null,
          },
        ],
        currentNarrative: textSignature(assetsCurrent.map((row) => row.commentary)),
        previousNarrative: textSignature(assetsPrevious.map((row) => row.commentary)),
        missingMessage: `No asset lifecycle rows are available for ${formatMonthLabel(reportingMonth)}.`,
      },
      Boolean(previousMonth),
    ),
    buildSectionRollup(
      {
        id: "change-release",
        label: "Change & Release",
        pageId: "p-change",
        tabId: null,
        metrics: [
          {
            label: "Success Rate",
            kind: "percent",
            current: changeCurrent?.changeSuccessRatePct ?? null,
            previous: changePrevious?.changeSuccessRatePct ?? null,
          },
          {
            label: "Releases Deployed",
            kind: "integer",
            current: changeCurrent?.releasesDeployed ?? null,
            previous: changePrevious?.releasesDeployed ?? null,
          },
          {
            label: "Failed Changes",
            kind: "integer",
            current: changeCurrent?.failedChanges ?? null,
            previous: changePrevious?.failedChanges ?? null,
          },
          {
            label: "Changes to Incidents",
            kind: "integer",
            current: changeCurrent?.changesCausingIncidents ?? null,
            previous: changePrevious?.changesCausingIncidents ?? null,
          },
        ],
        currentNarrative: changeCurrent?.commentary ?? "",
        previousNarrative: changePrevious?.commentary ?? "",
        missingMessage: `No change/release row is available for ${formatMonthLabel(reportingMonth)}.`,
      },
      Boolean(previousMonth),
    ),
    buildSectionRollup(
      {
        id: "development-delivery",
        label: "Development & Delivery",
        pageId: "p-dev",
        tabId: null,
        metrics: [
          {
            label: "Backlog End",
            kind: "integer",
            current: devCurrent?.devBacklogEnd ?? null,
            previous: devPrevious?.devBacklogEnd ?? null,
          },
          {
            label: "Tasks Closed",
            kind: "integer",
            current: devCurrent?.devTasksClosed ?? null,
            previous: devPrevious?.devTasksClosed ?? null,
          },
          {
            label: "Blocked Items",
            kind: "integer",
            current: devCurrent?.blockedItems ?? null,
            previous: devPrevious?.blockedItems ?? null,
          },
          {
            label: "Dev CSAT",
            kind: "score",
            current: devCurrent?.devCsatScore ?? null,
            previous: devPrevious?.devCsatScore ?? null,
          },
        ],
        currentNarrative: devCurrent?.commentary ?? "",
        previousNarrative: devPrevious?.commentary ?? "",
        missingMessage: `No development delivery row is available for ${formatMonthLabel(reportingMonth)}.`,
      },
      Boolean(previousMonth),
    ),
    buildSectionRollup(
      {
        id: "project-portfolio",
        label: "Project Portfolio",
        pageId: "p-projects",
        tabId: null,
        metrics: [
          {
            label: "Active Projects",
            kind: "integer",
            current: projectsCurrent.length || null,
            previous: projectsPrevious.length || null,
          },
          {
            label: "Avg Confidence",
            kind: "percent",
            current: averageConfidence(projectsCurrent),
            previous: averageConfidence(projectsPrevious),
          },
          {
            label: "Decisions Needed",
            kind: "integer",
            current: projectsCurrent.length ? decisionsNeeded(projectsCurrent) : null,
            previous: projectsPrevious.length ? decisionsNeeded(projectsPrevious) : null,
          },
        ],
        currentNarrative: textSignature(projectsCurrent.map((row) => row.commentary)),
        previousNarrative: textSignature(projectsPrevious.map((row) => row.commentary)),
        missingMessage: `No project portfolio rows are available for ${formatMonthLabel(reportingMonth)}.`,
      },
      Boolean(previousMonth),
    ),
    buildSectionRollup(
      {
        id: "rolling-roadmap",
        label: "Rolling Roadmap",
        pageId: "p-roadmap",
        tabId: null,
        metrics: [],
        currentNarrative: "",
        previousNarrative: "",
        missingMessage: "",
        notApplicable: true,
      },
      Boolean(previousMonth),
    ),
    buildSectionRollup(
      {
        id: "portfolio-gantt",
        label: "Portfolio Gantt",
        pageId: "p-gantt",
        tabId: null,
        metrics: [
          {
            label: "Active Workstreams",
            kind: "integer",
            current: ganttCurrent.length || null,
            previous: ganttPrevious.length || null,
          },
          {
            label: "On Track",
            kind: "integer",
            current: ganttCurrent.length ? onTrack(ganttCurrent) : null,
            previous: ganttPrevious.length ? onTrack(ganttPrevious) : null,
          },
          {
            label: "At Risk",
            kind: "integer",
            current: ganttCurrent.length ? atRisk(ganttCurrent) : null,
            previous: ganttPrevious.length ? atRisk(ganttPrevious) : null,
          },
          {
            label: "Milestones Due",
            kind: "integer",
            current: ganttMilestonesCurrent.length || null,
            previous: ganttMilestonesPrevious.length || null,
          },
        ],
        currentNarrative: textSignature(ganttCurrent.map((row) => row.detailCommentary)),
        previousNarrative: textSignature(ganttPrevious.map((row) => row.detailCommentary)),
        missingMessage: `No in-scope Gantt workstreams are available for ${formatMonthLabel(reportingMonth)}.`,
      },
      Boolean(previousMonth),
    ),
    buildSectionRollup(
      {
        id: "budget-commercials",
        label: "Budget & Commercials",
        pageId: "p-budget",
        tabId: null,
        metrics: [
          {
            label: "Total Budget",
            kind: "currency",
            current: budgetCurrent.length ? budgetTotals(budgetCurrent).budget : null,
            previous: budgetPrevious.length ? budgetTotals(budgetPrevious).budget : null,
          },
          {
            label: "Total Actual",
            kind: "currency",
            current: budgetCurrent.length ? budgetTotals(budgetCurrent).actual : null,
            previous: budgetPrevious.length ? budgetTotals(budgetPrevious).actual : null,
          },
          {
            label: "Forecast",
            kind: "currency",
            current: budgetCurrent.length ? budgetTotals(budgetCurrent).forecast : null,
            previous: budgetPrevious.length ? budgetTotals(budgetPrevious).forecast : null,
          },
          {
            label: "Variance",
            kind: "currency",
            current: budgetCurrent.length ? budgetTotals(budgetCurrent).variance : null,
            previous: budgetPrevious.length ? budgetTotals(budgetPrevious).variance : null,
          },
        ],
        currentNarrative: textSignature(budgetCurrent.map((row) => row.commentary)),
        previousNarrative: textSignature(budgetPrevious.map((row) => row.commentary)),
        missingMessage: `No budget/commercial rows are available for ${formatMonthLabel(reportingMonth)}.`,
      },
      Boolean(previousMonth),
    ),
    buildSectionRollup(
      {
        id: "risks-decisions",
        label: "Risks & Decisions",
        pageId: "p-risks",
        tabId: null,
        metrics: [
          {
            label: "Total Risks",
            kind: "integer",
            current: risksCurrent.length || null,
            previous: risksPrevious.length || null,
          },
          {
            label: "Decisions Needed",
            kind: "integer",
            current: risksCurrent.length ? risksCurrent.filter((row) => row.decisionRequired).length : null,
            previous: risksPrevious.length ? risksPrevious.filter((row) => row.decisionRequired).length : null,
          },
          {
            label: "Amber Risks",
            kind: "integer",
            current: risksCurrent.length ? amberRisks(risksCurrent) : null,
            previous: risksPrevious.length ? amberRisks(risksPrevious) : null,
          },
        ],
        currentNarrative: textSignature(risksCurrent.map((row) => row.commentary)),
        previousNarrative: textSignature(risksPrevious.map((row) => row.commentary)),
        missingMessage: `No risk register rows are available for ${formatMonthLabel(reportingMonth)}.`,
      },
      Boolean(previousMonth),
    ),
  ];

  return sections;
}

export function filterAcknowledgeableCheckIds(checks: ReportPrepCheck[], requestedIds: string[]): string[] {
  const allowedIds = new Set(checks.filter((check) => check.acknowledgeable).map((check) => check.id));
  return Array.from(new Set(requestedIds.filter((id) => allowedIds.has(id)))).sort();
}

export function buildReportPrepView(input: {
  snapshot: NormalizedReportSnapshot;
  reportingMonth: string;
  currentSummary: ExecSummaryState;
  previousMonthSummary: ExecSummaryState | null;
  mode: PrepMode;
  acknowledgedCheckIds?: string[];
  updatedAt?: string | null;
}): ReportPrepView {
  const previousMonth = getPreviousMonth(input.snapshot, input.reportingMonth);
  const previousMonthLabel = previousMonth ? formatMonthLabel(previousMonth) : null;
  const acknowledgedIdSet = new Set(input.acknowledgedCheckIds ?? []);
  const readinessChecks = buildReadinessChecks(
    input.snapshot,
    input.reportingMonth,
    previousMonth,
    input.currentSummary,
    acknowledgedIdSet,
  );
  const normalizedAcknowledgedIds = filterAcknowledgeableCheckIds(readinessChecks, input.acknowledgedCheckIds ?? []);
  const normalizedAcknowledgedSet = new Set(normalizedAcknowledgedIds);
  const checks = readinessChecks.map((check) =>
    check.acknowledgeable ? { ...check, acknowledged: normalizedAcknowledgedSet.has(check.id) } : check,
  );
  const openChecks = checks.filter((check) => !(check.acknowledgeable && check.acknowledged));
  const reviewedChecks = checks.filter((check) => check.acknowledgeable && check.acknowledged);
  const sections = buildSectionRollups(
    input.snapshot,
    input.reportingMonth,
    previousMonth,
    input.currentSummary,
    input.previousMonthSummary,
  );
  const previousExecSummary =
    previousMonth && input.previousMonthSummary
      ? {
          month: previousMonth,
          monthLabel: previousMonthLabel ?? previousMonth,
          mode: input.previousMonthSummary.mode,
          contentHtml: input.previousMonthSummary.contentHtml,
          excerpt: input.previousMonthSummary.excerpt,
          updatedAt: input.previousMonthSummary.updatedAt,
          available: hasMeaningfulText(input.previousMonthSummary.contentHtml),
        }
      : null;

  return {
    mode: input.mode,
    reportingMonth: input.reportingMonth,
    reportingMonthLabel: formatMonthLabel(input.reportingMonth),
    previousMonth,
    previousMonthLabel,
    updatedAt: input.updatedAt ?? null,
    acknowledgedCheckIds: normalizedAcknowledgedIds,
    readiness: {
      summary: {
        status: openChecks.some((check) => check.severity === "blocking" || check.severity === "warning") ? "needs-attention" : "ready",
        blockingCount: openChecks.filter((check) => check.severity === "blocking").length,
        warningCount: openChecks.filter((check) => check.severity === "warning").length,
        infoCount: openChecks.filter((check) => check.severity === "info").length,
        acknowledgedCount: reviewedChecks.length,
      },
      checks: openChecks,
      reviewedChecks,
    },
    rollover: {
      previousMonth,
      previousMonthLabel,
      sections,
      reviewQueue: previousMonth ? buildReviewQueue(sections, checks, previousExecSummary) : [],
      previousExecSummary,
    },
  };
}
