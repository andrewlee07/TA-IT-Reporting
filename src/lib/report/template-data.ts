import { format, parseISO } from "date-fns";

import { REPORT_PAGE_TABS } from "@/lib/report/blocks";
import type { ExecSummaryMode, ExecSummaryState } from "@/lib/reports/exec-summary";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

export interface TemplateData {
  meta: {
    availableMonths: string[];
    activeMonth: string;
    activeMonthLabel: string;
    monthLabels: Record<string, string>;
    monthRangeLabel: string;
    roadmapHorizonLabel: string;
    reportCutOffDates: Record<string, string>;
    templateKey: string;
    templateVersion: number;
    sourceFilename: string;
    pageTabs: Record<string, Array<{ id: string; label: string }>>;
  };
  execSummary: {
    mode: ExecSummaryMode;
    contentHtml: string;
    excerpt: string;
    updatedAt: string | null;
    sourceReportId: string | null;
  };
  support: Array<Record<string, number | string>>;
  service: Array<Record<string, number | string>>;
  security: Array<Record<string, number | string>>;
  assets: Array<Record<string, number | string>>;
  change: Array<Record<string, number | string>>;
  dev: Array<Record<string, number | string>>;
  projects: Array<Record<string, number | string>>;
  roadmap: Array<Record<string, number | string>>;
  ganttWorkstreams: Array<Record<string, number | string | boolean | null>>;
  ganttMilestones: Array<Record<string, number | string>>;
  chartSettings: Array<Record<string, number | string>>;
  budget: Array<Record<string, number | string>>;
  budgetMonthlyTotals: Array<Record<string, number | string>>;
  risks: Array<Record<string, number | string>>;
  tickets: Array<Record<string, number | string>>;
  narrative: Array<Record<string, number | string>>;
  officeLocations: Array<Record<string, number | string>>;
  officeNetwork: Array<Record<string, number | string>>;
  derivedNetwork: Array<Record<string, number | string>>;
}

export function formatMonthLabel(month: string): string {
  return format(parseISO(`${month}-01`), "MMMM yyyy");
}

export function formatMonthShort(month: string): string {
  return format(parseISO(`${month}-01`), "MMM");
}

function formatPct(value: number, decimals = 1): string {
  return `${value.toFixed(decimals)}%`;
}

function formatPctSmart(value: number): string {
  if (Number.isInteger(value)) {
    return `${value.toFixed(1)}%`;
  }

  if (Math.abs(value * 10 - Math.round(value * 10)) < 0.0001) {
    return `${value.toFixed(1)}%`;
  }

  return `${value.toFixed(2)}%`;
}

function formatScore(value: number): string {
  return `${value.toFixed(1)}/5`;
}

function yesNo(value: boolean): string {
  return value ? "Yes" : "No";
}

function sortRowsByMonth<T extends { Month: string }>(rows: T[]): T[] {
  return rows.slice().sort((left, right) => left.Month.localeCompare(right.Month));
}

function ensureMonthRow<T extends { Month: string }>(rows: T[], activeMonth: string, createDefaultRow: (month: string) => T): T[] {
  if (rows.some((row) => row.Month === activeMonth)) {
    return rows;
  }

  return sortRowsByMonth([...rows, createDefaultRow(activeMonth)]);
}

type AssetTemplateRow = {
  Month: string;
  AssetType: string;
  ActiveDevices: number;
  AvgAgeMths: number;
  PctWithin: string;
  PctOutside: string;
  StockOnHand: number;
  RefreshSpend: number;
  IncidentsLinked: number;
};

function ensureAssetRows(rows: AssetTemplateRow[], activeMonth: string): AssetTemplateRow[] {
  const assetTypes = ["Laptop", "Mobile", "Monitor"];
  const nextRows = [...rows];

  for (const assetType of assetTypes) {
    const hasRow = nextRows.some((row) => row.Month === activeMonth && row.AssetType === assetType);
    if (!hasRow) {
      nextRows.push({
        Month: activeMonth,
        AssetType: assetType,
        ActiveDevices: 0,
        AvgAgeMths: 0,
        PctWithin: "0.0%",
        PctOutside: "0.0%",
        StockOnHand: 0,
        RefreshSpend: 0,
        IncidentsLinked: 0,
      });
    }
  }

  return sortRowsByMonth(nextRows);
}

function createDefaultSupportRow(month: string) {
  return {
    Month: month,
    Opened: 0,
    Closed: 0,
    Backlog: 0,
    AvgAgeOpen: 0,
    AvgResolution: 0,
    FirstResponseSLA: "95.0%",
    ResolutionSLA: "95.0%",
    ReopenRate: "0.0%",
    MajorIncidents: 0,
    CSAT: "0.0/5",
    CSATRate: "0.0%",
    TopCategory: "Awaiting input",
    Commentary: "",
  };
}

function createDefaultSecurityRow(month: string) {
  return {
    Month: month,
    WkstationPatch: "0.0%",
    ServerPatch: "0.0%",
    CriticalPatch: "0.0%",
    DevicesOutside: 0,
    CritVulns: 0,
    HighVulns: 0,
    MedVulns: 0,
    LowVulns: 0,
    SecIncidents: 0,
    MFACoverage: "0.0%",
    EndpointCoverage: "0.0%",
    OverdueRemediation: 0,
    Commentary: "",
  };
}

function createDefaultChangeRow(month: string) {
  return {
    Month: month,
    TotalChanges: 0,
    StandardChanges: 0,
    NormalChanges: 0,
    EmergencyChanges: 0,
    SuccessfulChanges: 0,
    FailedChanges: 0,
    RolledBack: 0,
    SuccessRate: "89.0%",
    ChangesIncidents: 0,
    ReleasesDeployed: 0,
    Commentary: "",
  };
}

function createDefaultDevRow(month: string) {
  return {
    Month: month,
    Opened: 0,
    Closed: 0,
    BacklogEnd: 0,
    AvgAge: 0,
    OldestOpen: 0,
    Blocked: 0,
    Defects: 0,
    Enhancements: 0,
    TechDebt: 0,
    BAU: 0,
    CSAT: "0.0/5",
    Commentary: "",
  };
}

function createDefaultDerivedNetworkRow(month: string) {
  return {
    Month: month,
    Availability: "0.0%",
    OutageMins: 0,
    MajorIncidents: 0,
    PerfectOffices: 0,
    Below99_9Offices: 0,
    Below99Offices: 0,
    WorstOffice: "",
    WorstAvailability: "0.0%",
  };
}

function buildBudgetMonthlyTotals(snapshot: NormalizedReportSnapshot) {
  return snapshot.availableMonths.map((month) => {
    const rows = snapshot.budgetCommercials.filter((row) => row.reportingMonth === month);

    return {
      Month: month,
      Budget: rows.reduce((sum, row) => sum + row.budgetAmount, 0),
      Actual: rows.reduce((sum, row) => sum + row.actualAmount, 0),
      Forecast: rows.reduce((sum, row) => sum + row.forecastAmount, 0),
      Variance: rows.reduce((sum, row) => sum + row.variance, 0),
    };
  });
}

function buildRoadmapHorizon(snapshot: NormalizedReportSnapshot): string {
  const quarters = Array.from(new Set(snapshot.rollingRoadmap.map((row) => row.roadmapQuarter)));

  if (quarters.length === 0) {
    return "";
  }

  if (quarters.length === 1) {
    return quarters[0];
  }

  return `${quarters[0]} – ${quarters[quarters.length - 1]}`;
}

export function buildTemplateData(snapshot: NormalizedReportSnapshot, month: string, execSummary?: ExecSummaryState): TemplateData {
  const monthLabels = Object.fromEntries(snapshot.availableMonths.map((entry) => [entry, formatMonthShort(entry)]));
  const portfolioGanttWorkstreams = snapshot.portfolioGanttWorkstreams ?? [];
  const portfolioGanttMilestones = snapshot.portfolioGanttMilestones ?? [];
  const summaryState: ExecSummaryState = execSummary ?? {
    mode: "loading",
    contentHtml: "",
    excerpt: "",
    updatedAt: null,
    sourceReportId: null,
  };

  return {
    meta: {
      availableMonths: snapshot.availableMonths,
      activeMonth: month,
      activeMonthLabel: formatMonthLabel(month),
      monthLabels,
      monthRangeLabel: `${formatMonthShort(snapshot.availableMonths[0])} – ${formatMonthLabel(month)}`,
      roadmapHorizonLabel: buildRoadmapHorizon(snapshot),
      reportCutOffDates: Object.fromEntries(snapshot.periods.map((period) => [period.reportingMonth, period.reportCutOffDate ?? period.monthEndDate])),
      templateKey: snapshot.metadata.templateKey,
      templateVersion: snapshot.metadata.templateVersion,
      sourceFilename: snapshot.metadata.sourceFilename,
      pageTabs: Object.fromEntries(
        Object.entries(REPORT_PAGE_TABS).map(([pageId, tabs]) => [pageId, tabs.map((tab) => ({ id: tab.id, label: tab.label }))]),
      ),
    },
    execSummary: {
      mode: summaryState.mode,
      contentHtml: summaryState.contentHtml,
      excerpt: summaryState.excerpt,
      updatedAt: summaryState.updatedAt,
      sourceReportId: summaryState.sourceReportId,
    },
    support: ensureMonthRow(
      snapshot.supportOperations.map((row) => ({
        Month: row.reportingMonth,
        Opened: row.ticketsOpened,
        Closed: row.ticketsClosed,
        Backlog: row.backlogEnd,
        AvgAgeOpen: row.averageAgeOpenDays,
        AvgResolution: row.averageResolutionDays,
        FirstResponseSLA: formatPct(row.firstResponseSlaPct),
        ResolutionSLA: formatPct(row.resolutionSlaPct),
        ReopenRate: formatPct(row.reopenRatePct),
        MajorIncidents: row.majorIncidents,
        CSAT: formatScore(row.ticketCsatScore),
        CSATRate: formatPct(row.csatResponseRatePct),
        TopCategory: row.topCategory,
        Commentary: row.commentary,
      })),
      month,
      createDefaultSupportRow,
    ),
    service: snapshot.serviceAvailability.map((row) => ({
      Month: row.reportingMonth,
      Service: row.serviceName,
      Type: row.serviceType,
      Availability: formatPctSmart(row.availabilityPct),
      Target: formatPctSmart(row.targetPct),
      OutageMins: row.outageMinutes,
      MajorIncidents: row.majorIncidents,
      Commentary: row.commentary,
    })),
    security: ensureMonthRow(
      snapshot.securityPatching.map((row) => ({
        Month: row.reportingMonth,
        WkstationPatch: formatPct(row.workstationPatchCompliancePct),
        ServerPatch: formatPct(row.serverPatchCompliancePct),
        CriticalPatch: formatPct(row.criticalPatchCompliancePct),
        DevicesOutside: row.devicesOutsidePolicy,
        CritVulns: row.criticalVulns,
        HighVulns: row.highVulns,
        MedVulns: row.mediumVulns,
        LowVulns: row.lowVulns,
        SecIncidents: row.securityIncidents,
        MFACoverage: formatPct(row.mfaCoveragePct),
        EndpointCoverage: formatPct(row.endpointCoveragePct),
        OverdueRemediation: row.overdueRemediationItems,
        Commentary: row.commentary,
      })),
      month,
      createDefaultSecurityRow,
    ),
    assets: ensureAssetRows(
      snapshot.assetsLifecycle.map((row) => ({
        Month: row.reportingMonth,
        AssetType: row.assetType,
        ActiveDevices: row.activeDevices,
        AvgAgeMths: row.averageAgeMonths,
        PctWithin: formatPct(row.withinLifecyclePct),
        PctOutside: formatPct(row.outOfLifecyclePct),
        StockOnHand: row.stockOnHand,
        RefreshSpend: row.refreshSpend,
        IncidentsLinked: row.incidentsLinkedToAgedKit,
      })),
      month,
    ),
    change: ensureMonthRow(
      snapshot.changeRelease.map((row) => ({
        Month: row.reportingMonth,
        TotalChanges: row.totalChanges,
        StandardChanges: row.standardChanges,
        NormalChanges: row.normalChanges,
        EmergencyChanges: row.emergencyChanges,
        SuccessfulChanges: row.successfulChanges,
        FailedChanges: row.failedChanges,
        RolledBack: row.rolledBackChanges,
        SuccessRate: formatPct(row.changeSuccessRatePct),
        ChangesIncidents: row.changesCausingIncidents,
        ReleasesDeployed: row.releasesDeployed,
        Commentary: row.commentary,
      })),
      month,
      createDefaultChangeRow,
    ),
    dev: ensureMonthRow(
      snapshot.devDelivery.map((row) => ({
        Month: row.reportingMonth,
        Opened: row.devTasksOpened,
        Closed: row.devTasksClosed,
        BacklogEnd: row.devBacklogEnd,
        AvgAge: row.averageDevTaskAgeDays,
        OldestOpen: row.oldestOpenDevTaskDays,
        Blocked: row.blockedItems,
        Defects: row.defectsDelivered,
        Enhancements: row.enhancementsDelivered,
        TechDebt: row.techDebtDelivered,
        BAU: row.bauDelivered,
        CSAT: formatScore(row.devCsatScore),
        Commentary: row.commentary,
      })),
      month,
      createDefaultDevRow,
    ),
    projects: snapshot.projectPortfolio.map((row) => ({
      Month: row.reportingMonth,
      ProjectName: row.projectName,
      Sponsor: row.projectSponsor,
      StatusRAG: row.statusRag,
      Confidence: `${Math.round(row.deliveryConfidencePct)}%`,
      BudgetStatus: row.budgetStatus,
      MilestoneNext: row.milestoneNext,
      MilestoneDate: row.milestoneDate,
      ProjectCSAT: formatScore(row.projectCsatScore),
      Benefits: row.benefitsValueDelivered,
      Blockers: row.blockersDependencies,
      DecisionNeeded: yesNo(row.decisionNeeded),
      Commentary: row.commentary,
    })),
    roadmap: snapshot.rollingRoadmap.map((row) => ({
      Quarter: row.roadmapQuarter,
      Lane: row.lane,
      Initiative: row.initiative,
      StatusRAG: row.statusRag,
      Outcome: row.outcomeGoal,
      Owner: row.owner,
      Dependency: row.dependency,
      DecisionRequired: yesNo(row.decisionRequired),
      Notes: row.notes,
    })),
    ganttWorkstreams: portfolioGanttWorkstreams.map((row) => ({
      Month: row.reportingMonth,
      WorkstreamName: row.workstreamName,
      SponsorOwner: row.sponsorOwner,
      Domain: row.domain,
      StatusRAG: row.statusRag,
      StartDate: row.startDate,
      EndDate: row.endDate,
      ProgressDate: row.progressDate,
      Detail: row.detailCommentary,
      DisplayOrder: row.displayOrder,
      InScope: row.inScope,
    })),
    ganttMilestones: portfolioGanttMilestones.map((row) => ({
      Month: row.reportingMonth,
      WorkstreamName: row.workstreamName,
      MilestoneLabel: row.milestoneLabel,
      MilestoneDate: row.milestoneDate,
      DisplayOrder: row.displayOrder,
    })),
    chartSettings: (snapshot.chartSettings ?? []).map((row) => ({
      Month: row.reportingMonth,
      Page: row.page,
      ChartKey: row.chartKey,
      OverlayEnabled: row.overlayEnabled ? "Yes" : "No",
      OverlayMetric: row.overlayMetric,
      RollingWindow: row.rollingWindow,
      HealthyMin: row.healthyMin,
      AmberMin: row.amberMin,
      Commentary: row.commentary,
    })),
    budget: snapshot.budgetCommercials.map((row) => ({
      Month: row.reportingMonth,
      BudgetLine: row.budgetLine,
      Budget: row.budgetAmount,
      Actual: row.actualAmount,
      Forecast: row.forecastAmount,
      Variance: row.variance,
      Vendor: row.vendorContract,
      RenewalDue: row.renewalDueDate,
      RenewalValue: row.renewalValue,
      Owner: row.owner,
      Commentary: row.commentary,
    })),
    budgetMonthlyTotals: buildBudgetMonthlyTotals(snapshot),
    risks: snapshot.topRisks.map((row) => ({
      Month: row.reportingMonth,
      RiskIssue: row.riskIssue,
      Type: row.type,
      Owner: row.owner,
      Impact: row.impact,
      Likelihood: row.likelihood,
      RAG: row.ratingRag,
      Mitigation: row.currentControlMitigation,
      TargetDate: row.targetDate,
      DecisionRequired: yesNo(row.decisionRequired),
      Commentary: row.commentary,
    })),
    tickets: snapshot.oldestTickets.map((row) => ({
      Month: row.reportingMonth,
      TicketID: row.ticketId,
      Title: row.title,
      Category: row.category,
      OwnerQueue: row.ownerQueue,
      Status: row.currentStatus,
      AgeDays: row.ageDays,
      BusinessCritical: yesNo(row.businessCritical),
      Blocker: row.blockerReason,
      TargetDate: row.targetResolutionDate,
      NextAction: row.nextAction,
    })),
    narrative: snapshot.narrativeNotes.map((row) => ({
      Month: row.reportingMonth,
      Section: row.section,
      NoteType: row.noteType,
      Headline: row.headline,
      Narrative: row.narrative,
      Owner: row.owner,
    })),
    officeLocations: snapshot.officeLocations.map((row) => ({
      OfficeName: row.officeName,
      Region: row.region,
      DisplayOrder: row.displayOrder,
      MapX: row.mapX,
      MapY: row.mapY,
    })),
    officeNetwork: snapshot.officeLocations.flatMap((office) =>
      snapshot.officeNetworkAvailability
        .filter((row) => row.officeName === office.officeName)
        .map((row) => ({
          Month: row.reportingMonth,
          OfficeName: office.officeName,
          Region: office.region,
          DisplayOrder: office.displayOrder,
          MapX: office.mapX,
          MapY: office.mapY,
          Availability: formatPctSmart(row.availabilityPct),
          OutageMins: row.outageMinutes,
          MajorIncidents: row.majorIncidents,
          Commentary: row.commentary,
        })),
    ),
    derivedNetwork: ensureMonthRow(
      snapshot.derivedNetworkMetrics.map((row) => ({
        Month: row.reportingMonth,
        Availability: formatPctSmart(row.availabilityPct),
        OutageMins: row.outageMinutes,
        MajorIncidents: row.majorIncidents,
        PerfectOffices: row.perfectOffices,
        Below99_9Offices: row.below99_9Offices,
        Below99Offices: row.below99Offices,
        WorstOffice: row.worstOffice ?? "",
        WorstAvailability: row.worstAvailabilityPct === null ? "0.0%" : formatPctSmart(row.worstAvailabilityPct),
      })),
      month,
      createDefaultDerivedNetworkRow,
    ),
  };
}
