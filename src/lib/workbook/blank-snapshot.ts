import { format, lastDayOfMonth } from "date-fns";

import { WORKBOOK_TEMPLATE_KEY, WORKBOOK_TEMPLATE_VERSION } from "@/lib/workbook/contracts";
import type {
  AssetsLifecycleRow,
  BudgetCommercialRow,
  ChangeReleaseRow,
  ChartSettingRow,
  DevDeliveryRow,
  EntityRow,
  NarrativeNoteRow,
  NormalizedReportSnapshot,
  OfficeLocationRow,
  OfficeNetworkAvailabilityRow,
  OldestTicketRow,
  PeriodRow,
  PortfolioGanttMilestoneRow,
  PortfolioGanttWorkstreamRow,
  ProjectPortfolioRow,
  RollingRoadmapRow,
  SecurityPatchingRow,
  ServiceAvailabilityRow,
  SupportOperationsRow,
  TopRiskRow,
} from "@/lib/workbook/types";

function buildPeriodRow(reportingMonth: string): PeriodRow {
  const date = new Date(`${reportingMonth}-01T00:00:00Z`);
  const monthEnd = lastDayOfMonth(date);
  const quarter = Math.floor(monthEnd.getUTCMonth() / 3) + 1;

  return {
    reportingMonth,
    monthEndDate: format(monthEnd, "yyyy-MM-dd"),
    quarter: `Q${quarter}`,
    financialYear: `${monthEnd.getUTCFullYear()}/${String((monthEnd.getUTCFullYear() + 1) % 100).padStart(2, "0")}`,
    isCurrentPeriod: true,
    reportCutOffDate: format(monthEnd, "yyyy-MM-dd"),
  };
}

function buildSupportOperationsRow(reportingMonth: string): SupportOperationsRow {
  return {
    reportingMonth,
    ticketsOpened: 0,
    ticketsClosed: 0,
    backlogEnd: 0,
    averageAgeOpenDays: 0,
    averageResolutionDays: 0,
    firstResponseSlaPct: 95,
    resolutionSlaPct: 95,
    reopenRatePct: 0,
    majorIncidents: 0,
    ticketCsatScore: 0,
    csatResponseRatePct: 0,
    topCategory: "Awaiting input",
    commentary: "",
  };
}

function buildSecurityPatchingRow(reportingMonth: string): SecurityPatchingRow {
  return {
    reportingMonth,
    workstationPatchCompliancePct: 0,
    serverPatchCompliancePct: 0,
    criticalPatchCompliancePct: 0,
    devicesOutsidePolicy: 0,
    criticalVulns: 0,
    highVulns: 0,
    mediumVulns: 0,
    lowVulns: 0,
    securityIncidents: 0,
    mfaCoveragePct: 0,
    endpointCoveragePct: 0,
    overdueRemediationItems: 0,
    commentary: "",
  };
}

function buildAssetLifecycleRow(reportingMonth: string, assetType: string): AssetsLifecycleRow {
  return {
    reportingMonth,
    assetType,
    activeDevices: 0,
    averageAgeMonths: 0,
    withinLifecyclePct: 0,
    outOfLifecyclePct: 0,
    stockOnHand: 0,
    monthsStockCover: 0,
    awaitingDeployment: 0,
    awaitingDisposal: 0,
    refreshSpend: 0,
    incidentsLinkedToAgedKit: 0,
    commentary: "",
  };
}

function buildChangeReleaseRow(reportingMonth: string): ChangeReleaseRow {
  return {
    reportingMonth,
    totalChanges: 0,
    standardChanges: 0,
    normalChanges: 0,
    emergencyChanges: 0,
    successfulChanges: 0,
    failedChanges: 0,
    rolledBackChanges: 0,
    changeSuccessRatePct: 89,
    changesCausingIncidents: 0,
    releasesDeployed: 0,
    plannedMaintenanceCompleted: 0,
    commentary: "",
  };
}

function buildDevDeliveryRow(reportingMonth: string): DevDeliveryRow {
  return {
    reportingMonth,
    devTasksOpened: 0,
    devTasksClosed: 0,
    devBacklogEnd: 0,
    averageDevTaskAgeDays: 0,
    oldestOpenDevTaskDays: 0,
    blockedItems: 0,
    defectsDelivered: 0,
    enhancementsDelivered: 0,
    techDebtDelivered: 0,
    bauDelivered: 0,
    devCsatScore: 0,
    commentary: "",
  };
}

export function createBlankSnapshot(initialMonth: string, title: string): NormalizedReportSnapshot {
  return {
    metadata: {
      templateKey: WORKBOOK_TEMPLATE_KEY,
      templateVersion: WORKBOOK_TEMPLATE_VERSION,
      sourceFilename: `${title}.xlsx`,
      tableMap: {},
    },
    availableMonths: [initialMonth],
    currentMonth: initialMonth,
    periods: [buildPeriodRow(initialMonth)],
    entities: [] as EntityRow[],
    officeLocations: [] as OfficeLocationRow[],
    officeNetworkAvailability: [] as OfficeNetworkAvailabilityRow[],
    serviceAvailability: [] as ServiceAvailabilityRow[],
    supportOperations: [buildSupportOperationsRow(initialMonth)],
    oldestTickets: [] as OldestTicketRow[],
    securityPatching: [buildSecurityPatchingRow(initialMonth)],
    assetsLifecycle: [
      buildAssetLifecycleRow(initialMonth, "Laptop"),
      buildAssetLifecycleRow(initialMonth, "Mobile"),
      buildAssetLifecycleRow(initialMonth, "Monitor"),
    ],
    changeRelease: [buildChangeReleaseRow(initialMonth)],
    devDelivery: [buildDevDeliveryRow(initialMonth)],
    projectPortfolio: [] as ProjectPortfolioRow[],
    rollingRoadmap: [] as RollingRoadmapRow[],
    portfolioGanttWorkstreams: [] as PortfolioGanttWorkstreamRow[],
    portfolioGanttMilestones: [] as PortfolioGanttMilestoneRow[],
    chartSettings: [] as ChartSettingRow[],
    budgetCommercials: [] as BudgetCommercialRow[],
    topRisks: [] as TopRiskRow[],
    narrativeNotes: [] as NarrativeNoteRow[],
    derivedNetworkMetrics: [],
  };
}
