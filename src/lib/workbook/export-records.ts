import { getSheetContractsForVersion, type SheetContract } from "@/lib/workbook/contracts";
import { NETWORK_SERVICE_NAME } from "@/lib/workbook/derived-network";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

export interface WorkbookSheetExport {
  sheetName: string;
  headers: string[];
  tableName?: string;
  rows: Array<Array<string | number>>;
}

function yesNo(value: boolean): string {
  return value ? "Yes" : "No";
}

function stringValue(value: string | number | boolean | null | undefined): string | number {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "boolean") {
    return yesNo(value);
  }

  return value;
}

function buildRows(snapshot: NormalizedReportSnapshot, contract: SheetContract): Array<Array<string | number>> {
  switch (contract.sheetName) {
    case "Periods":
      return snapshot.periods.map((row) => [
        row.reportingMonth,
        row.monthEndDate,
        row.quarter,
        row.financialYear,
        yesNo(row.isCurrentPeriod),
        row.reportCutOffDate,
      ]);
    case "Entities":
      return snapshot.entities.map((row) => [
        row.entityType,
        row.entityName,
        row.grouping,
        yesNo(row.inScope),
        row.notes,
      ]);
    case "Office_Locations":
      return snapshot.officeLocations.map((row) => [
        row.officeName,
        row.region,
        yesNo(row.inScope),
        row.displayOrder,
        row.mapX,
        row.mapY,
      ]);
    case "INPUT_Office_Network_Avail":
      return snapshot.officeNetworkAvailability.map((row) => [
        row.reportingMonth,
        row.officeName,
        row.availabilityPct,
        row.outageMinutes,
        row.majorIncidents,
        row.commentary,
      ]);
    case "INPUT_Service_Availability":
      return snapshot.serviceAvailability
        .filter((row) => row.serviceName !== NETWORK_SERVICE_NAME)
        .map((row) => [
          row.reportingMonth,
          row.serviceName,
          row.serviceType,
          row.availabilityPct,
          row.targetPct,
          row.outageMinutes,
          row.majorIncidents,
          stringValue(row.backupSuccessPct),
          row.restoreTestStatus,
          row.commentary,
        ]);
    case "INPUT_Support_Operations":
      return snapshot.supportOperations.map((row) => [
        row.reportingMonth,
        row.ticketsOpened,
        row.ticketsClosed,
        row.backlogEnd,
        row.averageAgeOpenDays,
        row.averageResolutionDays,
        row.firstResponseSlaPct,
        row.resolutionSlaPct,
        row.reopenRatePct,
        row.majorIncidents,
        row.ticketCsatScore,
        row.csatResponseRatePct,
        row.topCategory,
        row.commentary,
      ]);
    case "INPUT_Top_Oldest_Tickets":
      return snapshot.oldestTickets.map((row) => [
        row.reportingMonth,
        row.ticketId,
        row.title,
        row.category,
        row.ownerQueue,
        row.currentStatus,
        row.ageDays,
        yesNo(row.businessCritical),
        row.blockerReason,
        row.targetResolutionDate,
        row.nextAction,
      ]);
    case "INPUT_Security_Patching":
      return snapshot.securityPatching.map((row) => [
        row.reportingMonth,
        row.workstationPatchCompliancePct,
        row.serverPatchCompliancePct,
        row.criticalPatchCompliancePct,
        row.devicesOutsidePolicy,
        row.criticalVulns,
        row.highVulns,
        row.mediumVulns,
        row.lowVulns,
        row.securityIncidents,
        row.mfaCoveragePct,
        row.endpointCoveragePct,
        row.overdueRemediationItems,
        row.commentary,
      ]);
    case "INPUT_Assets_Lifecycle":
      return snapshot.assetsLifecycle.map((row) => [
        row.reportingMonth,
        row.assetType,
        row.activeDevices,
        row.averageAgeMonths,
        row.withinLifecyclePct,
        row.outOfLifecyclePct,
        row.stockOnHand,
        row.monthsStockCover,
        row.awaitingDeployment,
        row.awaitingDisposal,
        row.refreshSpend,
        row.incidentsLinkedToAgedKit,
        row.commentary,
      ]);
    case "INPUT_Change_Release":
      return snapshot.changeRelease.map((row) => [
        row.reportingMonth,
        row.totalChanges,
        row.standardChanges,
        row.normalChanges,
        row.emergencyChanges,
        row.successfulChanges,
        row.failedChanges,
        row.rolledBackChanges,
        row.changeSuccessRatePct,
        row.changesCausingIncidents,
        row.releasesDeployed,
        row.plannedMaintenanceCompleted,
        row.commentary,
      ]);
    case "INPUT_Dev_Delivery":
      return snapshot.devDelivery.map((row) => [
        row.reportingMonth,
        row.devTasksOpened,
        row.devTasksClosed,
        row.devBacklogEnd,
        row.averageDevTaskAgeDays,
        row.oldestOpenDevTaskDays,
        row.blockedItems,
        row.defectsDelivered,
        row.enhancementsDelivered,
        row.techDebtDelivered,
        row.bauDelivered,
        row.devCsatScore,
        row.commentary,
      ]);
    case "INPUT_Project_Portfolio":
      return snapshot.projectPortfolio.map((row) => [
        row.reportingMonth,
        row.projectName,
        row.projectSponsor,
        row.statusRag,
        row.deliveryConfidencePct,
        row.budgetStatus,
        row.milestoneNext,
        row.milestoneDate,
        row.projectCsatScore,
        row.benefitsValueDelivered,
        row.blockersDependencies,
        yesNo(row.decisionNeeded),
        row.commentary,
      ]);
    case "INPUT_Rolling_Roadmap":
      return snapshot.rollingRoadmap.map((row) => [
        row.roadmapQuarter,
        row.lane,
        row.initiative,
        row.statusRag,
        row.outcomeGoal,
        row.owner,
        row.dependency,
        yesNo(row.decisionRequired),
        row.notes,
      ]);
    case "INPUT_Gantt_Workstreams":
      return snapshot.portfolioGanttWorkstreams.map((row) => [
        row.reportingMonth,
        row.workstreamName,
        row.sponsorOwner,
        row.domain,
        row.statusRag,
        row.startDate,
        row.endDate,
        stringValue(row.progressDate),
        row.detailCommentary,
        row.displayOrder,
        yesNo(row.inScope),
      ]);
    case "INPUT_Gantt_Milestones":
      return snapshot.portfolioGanttMilestones.map((row) => [
        row.reportingMonth,
        row.workstreamName,
        row.milestoneLabel,
        row.milestoneDate,
        row.displayOrder,
      ]);
    case "INPUT_Budget_Commercials":
      return snapshot.budgetCommercials.map((row) => [
        row.reportingMonth,
        row.budgetLine,
        row.budgetAmount,
        row.actualAmount,
        row.forecastAmount,
        row.variance,
        row.cloudLicensingSpend,
        row.assetRefreshSpend,
        row.savingsAchieved,
        row.avoidableCostRisk,
        row.vendorContract,
        row.renewalDueDate,
        row.renewalValue,
        row.owner,
        row.commentary,
      ]);
    case "INPUT_Top_Risks":
      return snapshot.topRisks.map((row) => [
        row.reportingMonth,
        row.riskIssue,
        row.type,
        row.owner,
        row.impact,
        row.likelihood,
        row.ratingRag,
        row.currentControlMitigation,
        row.targetDate,
        yesNo(row.decisionRequired),
        row.commentary,
      ]);
    case "INPUT_Narrative_Notes":
      return snapshot.narrativeNotes.map((row) => [
        row.reportingMonth,
        row.section,
        row.noteType,
        row.headline,
        row.narrative,
        row.owner,
      ]);
    case "INPUT_Chart_Settings":
      return snapshot.chartSettings.map((row) => [
        row.reportingMonth,
        row.page,
        row.chartKey,
        yesNo(row.overlayEnabled),
        row.overlayMetric,
        row.rollingWindow,
        row.healthyMin,
        row.amberMin,
        row.commentary,
      ]);
    default:
      return [];
  }
}

export function buildWorkbookSheetExports(snapshot: NormalizedReportSnapshot): WorkbookSheetExport[] {
  return getSheetContractsForVersion(snapshot.metadata.templateVersion).map((contract) => ({
    sheetName: contract.sheetName,
    headers: contract.headers,
    tableName: contract.tableName,
    rows: buildRows(snapshot, contract),
  }));
}
