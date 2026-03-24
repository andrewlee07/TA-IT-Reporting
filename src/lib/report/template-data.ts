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
    support: snapshot.supportOperations.map((row) => ({
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
    security: snapshot.securityPatching.map((row) => ({
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
    assets: snapshot.assetsLifecycle.map((row) => ({
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
    change: snapshot.changeRelease.map((row) => ({
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
    dev: snapshot.devDelivery.map((row) => ({
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
    derivedNetwork: snapshot.derivedNetworkMetrics.map((row) => ({
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
  };
}
