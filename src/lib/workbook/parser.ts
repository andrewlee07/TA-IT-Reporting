import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import * as XLSX from "xlsx";
import { format, isValid, parse, parseISO } from "date-fns";

import {
  OFFICE_NETWORK_SHEET_NAME,
  PORTFOLIO_GANTT_DOMAINS,
  PORTFOLIO_GANTT_MILESTONES_SHEET_NAME,
  PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME,
  REQUIRED_SHEET_NAMES,
  SHEET_CONTRACTS,
  WORKBOOK_TEMPLATE_KEY,
  WORKBOOK_TEMPLATE_VERSION,
} from "@/lib/workbook/contracts";
import type {
  AssetsLifecycleRow,
  BudgetCommercialRow,
  ChangeReleaseRow,
  DerivedNetworkMetricRow,
  DevDeliveryRow,
  EntityRow,
  NarrativeNoteRow,
  NormalizedReportSnapshot,
  OfficeLocationRow,
  OfficeNetworkAvailabilityRow,
  OldestTicketRow,
  ParseWorkbookResult,
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
import { WorkbookValidationError } from "@/lib/workbook/types";

const README_SHEET = "README";
const NETWORK_SERVICE_NAME = "Network";
const NETWORK_TARGET_PCT = 99.9;
const DATE_PATTERNS = ["yyyy-MM-dd", "d/M/yyyy", "dd/MM/yyyy", "M/d/yyyy", "MM/dd/yyyy", "d MMM yyyy", "dd MMM yyyy"];

type SheetRecord = Record<string, string>;

interface MetadataRecord {
  templateKey: string;
  templateVersion: number;
}

interface ParseWorkbookBufferOptions {
  skipTableValidation?: boolean;
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  trimValues: false,
});

function resolveXlsxPath(target: string, basePrefix = "xl/"): string {
  if (target.startsWith("/")) {
    return target.slice(1);
  }

  if (target.startsWith("../")) {
    return `${basePrefix}${target.slice(3)}`;
  }

  return `${basePrefix}${target.replace(/^\.?\//, "")}`;
}

function asArray<T>(value: T | T[] | undefined): T[] {
  if (!value) {
    return [];
  }

  return Array.isArray(value) ? value : [value];
}

function normalizeCell(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}

function parseBoolean(value: string, fieldName: string): boolean {
  const normalized = value.trim().toLowerCase();

  if (normalized === "yes" || normalized === "true") {
    return true;
  }

  if (normalized === "no" || normalized === "false") {
    return false;
  }

  throw new Error(`${fieldName} must be Yes/No or True/False.`);
}

function parseNumber(value: string, fieldName: string): number {
  const normalized = value.replace(/,/g, "").trim();
  const parsed = Number(normalized);

  if (Number.isNaN(parsed)) {
    throw new Error(`${fieldName} must be numeric.`);
  }

  return parsed;
}

function parsePercentage(value: string, fieldName: string): number {
  const normalized = value.trim().replace(/%$/, "");
  const parsed = Number(normalized);

  if (Number.isNaN(parsed)) {
    throw new Error(`${fieldName} must be a percentage.`);
  }

  return parsed;
}

function parseScoreOutOfFive(value: string, fieldName: string): number {
  const match = value.trim().match(/^(\d+(?:\.\d+)?)\s*\/\s*5$/);

  if (!match) {
    throw new Error(`${fieldName} must use the format x/5.`);
  }

  return Number(match[1]);
}

function parseDateString(value: string, fieldName: string): string {
  const trimmed = value.trim();

  if (!trimmed) {
    throw new Error(`${fieldName} is required.`);
  }

  const attemptedDates = [parseISO(trimmed), ...DATE_PATTERNS.map((pattern) => parse(trimmed, pattern, new Date()))];
  const parsedDate = attemptedDates.find((candidate) => isValid(candidate)) ?? new Date(trimmed);

  if (!isValid(parsedDate)) {
    throw new Error(`${fieldName} must be a valid date.`);
  }

  return format(parsedDate, "yyyy-MM-dd");
}

function parseOptionalDateString(value: string, fieldName: string): string | null {
  const trimmed = value.trim();

  if (!trimmed) {
    return null;
  }

  return parseDateString(trimmed, fieldName);
}

function roundTo(value: number, decimals = 2): number {
  const factor = 10 ** decimals;
  return Math.round(value * factor) / factor;
}

function getSheetMatrix(workbook: XLSX.WorkBook, sheetName: string): string[][] {
  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    throw new Error(`Missing worksheet: ${sheetName}`);
  }

  return (XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: false,
  }) as unknown[][]).map((row) => row.map(normalizeCell));
}

function findWorkbookMetadata(readmeMatrix: string[][]): MetadataRecord {
  let templateKey = "";
  let templateVersion = 0;

  for (const row of readmeMatrix) {
    const [label, value] = row;

    if (label === "Template Key") {
      templateKey = normalizeCell(value);
    }

    if (label === "Template Version") {
      templateVersion = Number.parseInt(normalizeCell(value), 10);
    }
  }

  return {
    templateKey,
    templateVersion,
  };
}

function validateHeaderRow(sheetName: string, matrix: string[][], expectedHeaders: string[]): void {
  const headerRow = matrix[2] ?? [];
  const actualHeaders = headerRow.slice(0, expectedHeaders.length).map(normalizeCell);

  if (actualHeaders.length !== expectedHeaders.length) {
    throw new Error(`${sheetName} must contain the exact header row on row 3.`);
  }

  for (let index = 0; index < expectedHeaders.length; index += 1) {
    if (actualHeaders[index] !== expectedHeaders[index]) {
      throw new Error(`${sheetName} header mismatch at column ${index + 1}. Expected "${expectedHeaders[index]}", received "${actualHeaders[index]}".`);
    }
  }
}

function sheetMatrixToRecords(sheetName: string, matrix: string[][], headers: string[]): SheetRecord[] {
  validateHeaderRow(sheetName, matrix, headers);

  return matrix
    .slice(3)
    .filter((row) => row.some((value) => normalizeCell(value) !== ""))
    .map((row) => {
      const record: SheetRecord = {};

      headers.forEach((header, index) => {
        record[header] = normalizeCell(row[index]);
      });

      return record;
    });
}

async function extractTableMap(buffer: Buffer): Promise<Record<string, string[]>> {
  const zip = await JSZip.loadAsync(buffer);
  const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
  const workbookRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("text");

  if (!workbookXml || !workbookRelsXml) {
    return {};
  }

  const workbook = xmlParser.parse(workbookXml);
  const workbookRels = xmlParser.parse(workbookRelsXml);
  const rels = new Map<string, string>();

  for (const relationship of asArray(workbookRels.Relationships?.Relationship)) {
    rels.set(relationship.Id, relationship.Target);
  }

  const tableMap: Record<string, string[]> = {};

  for (const sheet of asArray(workbook.workbook?.sheets?.sheet)) {
    const relationshipId = sheet["r:id"];
    const target = rels.get(relationshipId);

    if (!target) {
      continue;
    }

    const sheetPath = resolveXlsxPath(target);
    const worksheetXml = await zip.file(sheetPath)?.async("text");

    if (!worksheetXml) {
      continue;
    }

    const worksheet = xmlParser.parse(worksheetXml);
    const tableParts = asArray(worksheet.worksheet?.tableParts?.tablePart);

    if (tableParts.length === 0) {
      tableMap[sheet.name] = [];
      continue;
    }

    const sheetRelsPath = sheetPath.replace("worksheets/", "worksheets/_rels/") + ".rels";
    const sheetRelsXml = await zip.file(sheetRelsPath)?.async("text");

    if (!sheetRelsXml) {
      tableMap[sheet.name] = [];
      continue;
    }

    const sheetRels = xmlParser.parse(sheetRelsXml);
    const sheetRelMap = new Map<string, string>();

    for (const relationship of asArray(sheetRels.Relationships?.Relationship)) {
      sheetRelMap.set(relationship.Id, relationship.Target);
    }

    const tableNames: string[] = [];

    for (const tablePart of tableParts) {
      const tableTarget = sheetRelMap.get(tablePart["r:id"]);

      if (!tableTarget) {
        continue;
      }

      const tablePath = resolveXlsxPath(tableTarget);
      const tableXml = await zip.file(tablePath)?.async("text");

      if (!tableXml) {
        continue;
      }

      const table = xmlParser.parse(tableXml);
      tableNames.push(table.table?.displayName ?? table.table?.name ?? "");
    }

    tableMap[sheet.name] = tableNames.filter(Boolean);
  }

  return tableMap;
}

function validateTables(tableMap: Record<string, string[]>, issues: string[]): void {
  for (const contract of SHEET_CONTRACTS) {
    if (!contract.tableName) {
      continue;
    }

    const tableNames = tableMap[contract.sheetName] ?? [];

    if (!tableNames.includes(contract.tableName)) {
      issues.push(`${contract.sheetName} must contain the table "${contract.tableName}".`);
    }
  }
}

function ensureRequiredSheets(workbook: XLSX.WorkBook, issues: string[]): void {
  const presentSheets = new Set(workbook.SheetNames);

  for (const requiredSheet of REQUIRED_SHEET_NAMES) {
    if (!presentSheets.has(requiredSheet)) {
      issues.push(`Missing required worksheet: ${requiredSheet}.`);
    }
  }

  if (!presentSheets.has(README_SHEET)) {
    issues.push("Missing required worksheet: README.");
  }
}

function parsePeriods(records: SheetRecord[]): PeriodRow[] {
  return records.map((record) => ({
    reportingMonth: record["Reporting Month"],
    monthEndDate: record["Month End Date"],
    quarter: record["Quarter"],
    financialYear: record["Financial Year"],
    isCurrentPeriod: parseBoolean(record["Is Current Period"], "Periods.Is Current Period"),
    reportCutOffDate: parseDateString(record["Report Cut-Off Date"], "Periods.Report Cut-Off Date"),
  }));
}

function parseEntities(records: SheetRecord[]): EntityRow[] {
  return records.map((record) => ({
    entityType: record["Entity Type"],
    entityName: record["Entity Name"],
    grouping: record["Grouping"],
    inScope: parseBoolean(record["In Scope"], "Entities.In Scope"),
    notes: record["Notes"],
  }));
}

function parseOfficeLocations(records: SheetRecord[]): OfficeLocationRow[] {
  return records.map((record) => ({
    officeName: record["Office Name"],
    region: record["Region"],
    inScope: parseBoolean(record["In Scope"], "Office_Locations.In Scope"),
    displayOrder: parseNumber(record["Display Order"], "Office_Locations.Display Order"),
    mapX: parseNumber(record["Map X"], "Office_Locations.Map X"),
    mapY: parseNumber(record["Map Y"], "Office_Locations.Map Y"),
  }));
}

function parseOfficeNetworkAvailability(records: SheetRecord[]): OfficeNetworkAvailabilityRow[] {
  return records.map((record) => ({
    reportingMonth: record["Reporting Month"],
    officeName: record["Office Name"],
    availabilityPct: parsePercentage(record["Availability %"], `${OFFICE_NETWORK_SHEET_NAME}.Availability %`),
    outageMinutes: parseNumber(record["Outage Minutes"], `${OFFICE_NETWORK_SHEET_NAME}.Outage Minutes`),
    majorIncidents: parseNumber(record["Major Incidents"], `${OFFICE_NETWORK_SHEET_NAME}.Major Incidents`),
    commentary: record["Commentary"],
  }));
}

function parseServiceAvailability(records: SheetRecord[]): ServiceAvailabilityRow[] {
  return records.map((record) => ({
    reportingMonth: record["Reporting Month"],
    serviceName: record["Service Name"],
    serviceType: record["Service Type"],
    availabilityPct: parsePercentage(record["Availability %"], "INPUT_Service_Availability.Availability %"),
    targetPct: parsePercentage(record["Target %"], "INPUT_Service_Availability.Target %"),
    outageMinutes: parseNumber(record["Outage Minutes"], "INPUT_Service_Availability.Outage Minutes"),
    majorIncidents: parseNumber(record["Major Incidents"], "INPUT_Service_Availability.Major Incidents"),
    backupSuccessPct: record["Backup Success %"] ? parsePercentage(record["Backup Success %"], "INPUT_Service_Availability.Backup Success %") : null,
    restoreTestStatus: record["Restore Test Status"],
    commentary: record["Commentary"],
  }));
}

function parseSupportOperations(records: SheetRecord[]): SupportOperationsRow[] {
  return records.map((record) => ({
    reportingMonth: record["Reporting Month"],
    ticketsOpened: parseNumber(record["Tickets Opened"], "INPUT_Support_Operations.Tickets Opened"),
    ticketsClosed: parseNumber(record["Tickets Closed"], "INPUT_Support_Operations.Tickets Closed"),
    backlogEnd: parseNumber(record["Backlog End"], "INPUT_Support_Operations.Backlog End"),
    averageAgeOpenDays: parseNumber(record["Average Age Open Days"], "INPUT_Support_Operations.Average Age Open Days"),
    averageResolutionDays: parseNumber(record["Average Resolution Days"], "INPUT_Support_Operations.Average Resolution Days"),
    firstResponseSlaPct: parsePercentage(record["First Response SLA %"], "INPUT_Support_Operations.First Response SLA %"),
    resolutionSlaPct: parsePercentage(record["Resolution SLA %"], "INPUT_Support_Operations.Resolution SLA %"),
    reopenRatePct: parsePercentage(record["Reopen Rate %"], "INPUT_Support_Operations.Reopen Rate %"),
    majorIncidents: parseNumber(record["Major Incidents"], "INPUT_Support_Operations.Major Incidents"),
    ticketCsatScore: parseScoreOutOfFive(record["Ticket CSAT"], "INPUT_Support_Operations.Ticket CSAT"),
    csatResponseRatePct: parsePercentage(record["CSAT Response Rate %"], "INPUT_Support_Operations.CSAT Response Rate %"),
    topCategory: record["Top Category"],
    commentary: record["Commentary"],
  }));
}

function parseOldestTickets(records: SheetRecord[]): OldestTicketRow[] {
  return records.map((record) => ({
    reportingMonth: record["Reporting Month"],
    ticketId: record["Ticket ID"],
    title: record["Title"],
    category: record["Category"],
    ownerQueue: record["Owner / Queue"],
    currentStatus: record["Current Status"],
    ageDays: parseNumber(record["Age Days"], "INPUT_Top_Oldest_Tickets.Age Days"),
    businessCritical: parseBoolean(record["Business Critical?"], "INPUT_Top_Oldest_Tickets.Business Critical?"),
    blockerReason: record["Blocker / Reason Still Open"],
    targetResolutionDate: record["Target Resolution Date"],
    nextAction: record["Next Action"],
  }));
}

function parseSecurityPatching(records: SheetRecord[]): SecurityPatchingRow[] {
  return records.map((record) => ({
    reportingMonth: record["Reporting Month"],
    workstationPatchCompliancePct: parsePercentage(record["Workstation Patch Compliance %"], "INPUT_Security_Patching.Workstation Patch Compliance %"),
    serverPatchCompliancePct: parsePercentage(record["Server Patch Compliance %"], "INPUT_Security_Patching.Server Patch Compliance %"),
    criticalPatchCompliancePct: parsePercentage(record["Critical Patch Compliance %"], "INPUT_Security_Patching.Critical Patch Compliance %"),
    devicesOutsidePolicy: parseNumber(record["Devices Outside Policy"], "INPUT_Security_Patching.Devices Outside Policy"),
    criticalVulns: parseNumber(record["Critical Vulns"], "INPUT_Security_Patching.Critical Vulns"),
    highVulns: parseNumber(record["High Vulns"], "INPUT_Security_Patching.High Vulns"),
    mediumVulns: parseNumber(record["Medium Vulns"], "INPUT_Security_Patching.Medium Vulns"),
    lowVulns: parseNumber(record["Low Vulns"], "INPUT_Security_Patching.Low Vulns"),
    securityIncidents: parseNumber(record["Security Incidents"], "INPUT_Security_Patching.Security Incidents"),
    mfaCoveragePct: parsePercentage(record["MFA Coverage %"], "INPUT_Security_Patching.MFA Coverage %"),
    endpointCoveragePct: parsePercentage(record["Endpoint Coverage %"], "INPUT_Security_Patching.Endpoint Coverage %"),
    overdueRemediationItems: parseNumber(record["Overdue Remediation Items"], "INPUT_Security_Patching.Overdue Remediation Items"),
    commentary: record["Commentary"],
  }));
}

function parseAssetsLifecycle(records: SheetRecord[]): AssetsLifecycleRow[] {
  return records.map((record) => ({
    reportingMonth: record["Reporting Month"],
    assetType: record["Asset Type"],
    activeDevices: parseNumber(record["Active Devices"], "INPUT_Assets_Lifecycle.Active Devices"),
    averageAgeMonths: parseNumber(record["Average Age Months"], "INPUT_Assets_Lifecycle.Average Age Months"),
    withinLifecyclePct: parsePercentage(record["% Within Lifecycle"], "INPUT_Assets_Lifecycle.% Within Lifecycle"),
    outOfLifecyclePct: parsePercentage(record["% Out of Lifecycle"], "INPUT_Assets_Lifecycle.% Out of Lifecycle"),
    stockOnHand: parseNumber(record["Stock On Hand"], "INPUT_Assets_Lifecycle.Stock On Hand"),
    monthsStockCover: parseNumber(record["Months Stock Cover"], "INPUT_Assets_Lifecycle.Months Stock Cover"),
    awaitingDeployment: parseNumber(record["Awaiting Deployment"], "INPUT_Assets_Lifecycle.Awaiting Deployment"),
    awaitingDisposal: parseNumber(record["Awaiting Disposal"], "INPUT_Assets_Lifecycle.Awaiting Disposal"),
    refreshSpend: parseNumber(record["Refresh Spend"], "INPUT_Assets_Lifecycle.Refresh Spend"),
    incidentsLinkedToAgedKit: parseNumber(record["Incidents Linked to Aged Kit"], "INPUT_Assets_Lifecycle.Incidents Linked to Aged Kit"),
    commentary: record["Commentary"],
  }));
}

function parseChangeRelease(records: SheetRecord[]): ChangeReleaseRow[] {
  return records.map((record) => ({
    reportingMonth: record["Reporting Month"],
    totalChanges: parseNumber(record["Total Changes"], "INPUT_Change_Release.Total Changes"),
    standardChanges: parseNumber(record["Standard Changes"], "INPUT_Change_Release.Standard Changes"),
    normalChanges: parseNumber(record["Normal Changes"], "INPUT_Change_Release.Normal Changes"),
    emergencyChanges: parseNumber(record["Emergency Changes"], "INPUT_Change_Release.Emergency Changes"),
    successfulChanges: parseNumber(record["Successful Changes"], "INPUT_Change_Release.Successful Changes"),
    failedChanges: parseNumber(record["Failed Changes"], "INPUT_Change_Release.Failed Changes"),
    rolledBackChanges: parseNumber(record["Rolled Back Changes"], "INPUT_Change_Release.Rolled Back Changes"),
    changeSuccessRatePct: parsePercentage(record["Change Success Rate %"], "INPUT_Change_Release.Change Success Rate %"),
    changesCausingIncidents: parseNumber(record["Changes Causing Incidents"], "INPUT_Change_Release.Changes Causing Incidents"),
    releasesDeployed: parseNumber(record["Releases Deployed"], "INPUT_Change_Release.Releases Deployed"),
    plannedMaintenanceCompleted: parseNumber(record["Planned Maintenance Completed"], "INPUT_Change_Release.Planned Maintenance Completed"),
    commentary: record["Commentary"],
  }));
}

function parseDevDelivery(records: SheetRecord[]): DevDeliveryRow[] {
  return records.map((record) => ({
    reportingMonth: record["Reporting Month"],
    devTasksOpened: parseNumber(record["Dev Tasks Opened"], "INPUT_Dev_Delivery.Dev Tasks Opened"),
    devTasksClosed: parseNumber(record["Dev Tasks Closed"], "INPUT_Dev_Delivery.Dev Tasks Closed"),
    devBacklogEnd: parseNumber(record["Dev Backlog End"], "INPUT_Dev_Delivery.Dev Backlog End"),
    averageDevTaskAgeDays: parseNumber(record["Average Dev Task Age Days"], "INPUT_Dev_Delivery.Average Dev Task Age Days"),
    oldestOpenDevTaskDays: parseNumber(record["Oldest Open Dev Task Days"], "INPUT_Dev_Delivery.Oldest Open Dev Task Days"),
    blockedItems: parseNumber(record["Blocked Items"], "INPUT_Dev_Delivery.Blocked Items"),
    defectsDelivered: parseNumber(record["Defects Delivered"], "INPUT_Dev_Delivery.Defects Delivered"),
    enhancementsDelivered: parseNumber(record["Enhancements Delivered"], "INPUT_Dev_Delivery.Enhancements Delivered"),
    techDebtDelivered: parseNumber(record["Tech Debt Delivered"], "INPUT_Dev_Delivery.Tech Debt Delivered"),
    bauDelivered: parseNumber(record["BAU Delivered"], "INPUT_Dev_Delivery.BAU Delivered"),
    devCsatScore: parseScoreOutOfFive(record["Dev CSAT / Sponsor Score"], "INPUT_Dev_Delivery.Dev CSAT / Sponsor Score"),
    commentary: record["Commentary"],
  }));
}

function parseProjectPortfolio(records: SheetRecord[]): ProjectPortfolioRow[] {
  return records.map((record) => ({
    reportingMonth: record["Reporting Month"],
    projectName: record["Project Name"],
    projectSponsor: record["Project Sponsor"],
    statusRag: record["Status RAG"],
    deliveryConfidencePct: parsePercentage(record["Delivery Confidence %"], "INPUT_Project_Portfolio.Delivery Confidence %"),
    budgetStatus: record["Budget Status"],
    milestoneNext: record["Milestone Next"],
    milestoneDate: record["Milestone Date"],
    projectCsatScore: parseScoreOutOfFive(record["Project CSAT"], "INPUT_Project_Portfolio.Project CSAT"),
    benefitsValueDelivered: record["Benefits / Value Delivered"],
    blockersDependencies: record["Blockers / Dependencies"],
    decisionNeeded: parseBoolean(record["Decision Needed?"], "INPUT_Project_Portfolio.Decision Needed?"),
    commentary: record["Commentary"],
  }));
}

function parseRollingRoadmap(records: SheetRecord[]): RollingRoadmapRow[] {
  return records.map((record) => ({
    roadmapQuarter: record["Roadmap Quarter"],
    lane: record["Lane"],
    initiative: record["Initiative"],
    statusRag: record["Status RAG"],
    outcomeGoal: record["Outcome / Goal"],
    owner: record["Owner"],
    dependency: record["Dependency"],
    decisionRequired: parseBoolean(record["Decision Required"], "INPUT_Rolling_Roadmap.Decision Required"),
    notes: record["Notes"],
  }));
}

function parsePortfolioGanttWorkstreams(records: SheetRecord[]): PortfolioGanttWorkstreamRow[] {
  return records.map((record) => {
    const domain = record["Domain"];

    if (!PORTFOLIO_GANTT_DOMAINS.includes(domain as (typeof PORTFOLIO_GANTT_DOMAINS)[number])) {
      throw new Error(
        `${PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME}.Domain must be one of: ${PORTFOLIO_GANTT_DOMAINS.join(", ")}.`,
      );
    }

    const startDate = parseDateString(record["Start Date"], `${PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME}.Start Date`);
    const endDate = parseDateString(record["End Date"], `${PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME}.End Date`);

    if (startDate > endDate) {
      throw new Error(`${PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME}.Start Date must be on or before End Date.`);
    }

    return {
      reportingMonth: record["Reporting Month"],
      workstreamName: record["Workstream Name"],
      sponsorOwner: record["Sponsor / Owner"],
      domain,
      statusRag: record["Status RAG"],
      startDate,
      endDate,
      progressDate: parseOptionalDateString(record["Progress Date"], `${PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME}.Progress Date`),
      detailCommentary: record["Detail / Commentary"],
      displayOrder: parseNumber(record["Display Order"], `${PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME}.Display Order`),
      inScope: parseBoolean(record["In Scope"], `${PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME}.In Scope`),
    };
  });
}

function parsePortfolioGanttMilestones(records: SheetRecord[]): PortfolioGanttMilestoneRow[] {
  return records.map((record) => ({
    reportingMonth: record["Reporting Month"],
    workstreamName: record["Workstream Name"],
    milestoneLabel: record["Milestone Label"],
    milestoneDate: parseDateString(record["Milestone Date"], `${PORTFOLIO_GANTT_MILESTONES_SHEET_NAME}.Milestone Date`),
    displayOrder: parseNumber(record["Display Order"], `${PORTFOLIO_GANTT_MILESTONES_SHEET_NAME}.Display Order`),
  }));
}

function parseBudgetCommercials(records: SheetRecord[]): BudgetCommercialRow[] {
  return records.map((record) => ({
    reportingMonth: record["Reporting Month"],
    budgetLine: record["Budget Line"],
    budgetAmount: parseNumber(record["Budget Amount"], "INPUT_Budget_Commercials.Budget Amount"),
    actualAmount: parseNumber(record["Actual Amount"], "INPUT_Budget_Commercials.Actual Amount"),
    forecastAmount: parseNumber(record["Forecast Amount"], "INPUT_Budget_Commercials.Forecast Amount"),
    variance: parseNumber(record["Variance"], "INPUT_Budget_Commercials.Variance"),
    cloudLicensingSpend: parseNumber(record["Cloud / Licensing Spend"], "INPUT_Budget_Commercials.Cloud / Licensing Spend"),
    assetRefreshSpend: parseNumber(record["Asset Refresh Spend"], "INPUT_Budget_Commercials.Asset Refresh Spend"),
    savingsAchieved: parseNumber(record["Savings Achieved"], "INPUT_Budget_Commercials.Savings Achieved"),
    avoidableCostRisk: parseNumber(record["Avoidable Cost Risk"], "INPUT_Budget_Commercials.Avoidable Cost Risk"),
    vendorContract: record["Vendor / Contract"],
    renewalDueDate: record["Renewal Due Date"],
    renewalValue: parseNumber(record["Renewal Value"], "INPUT_Budget_Commercials.Renewal Value"),
    owner: record["Owner"],
    commentary: record["Commentary"],
  }));
}

function parseTopRisks(records: SheetRecord[]): TopRiskRow[] {
  return records.map((record) => ({
    reportingMonth: record["Reporting Month"],
    riskIssue: record["Risk / Issue"],
    type: record["Type"],
    owner: record["Owner"],
    impact: record["Impact"],
    likelihood: record["Likelihood"],
    ratingRag: record["Rating RAG"],
    currentControlMitigation: record["Current Control / Mitigation"],
    targetDate: record["Target Date"],
    decisionRequired: parseBoolean(record["Decision Required?"], "INPUT_Top_Risks.Decision Required?"),
    commentary: record["Commentary"],
  }));
}

function parseNarrativeNotes(records: SheetRecord[]): NarrativeNoteRow[] {
  return records.map((record) => ({
    reportingMonth: record["Reporting Month"],
    section: record["Section"],
    noteType: record["Note Type"],
    headline: record["Headline"],
    narrative: record["Narrative"],
    owner: record["Owner"],
  }));
}

function deriveNetworkMetrics(
  periods: PeriodRow[],
  officeLocations: OfficeLocationRow[],
  officeNetworkAvailability: OfficeNetworkAvailabilityRow[],
  issues: string[],
): DerivedNetworkMetricRow[] {
  const inScopeOffices = officeLocations.filter((office) => office.inScope);
  const officeNames = new Set(inScopeOffices.map((office) => office.officeName));

  return periods.map((period) => {
    const monthRows = officeNetworkAvailability.filter((row) => row.reportingMonth === period.reportingMonth && officeNames.has(row.officeName));
    const byOffice = new Map<string, OfficeNetworkAvailabilityRow[]>();

    for (const row of monthRows) {
      const rows = byOffice.get(row.officeName) ?? [];
      rows.push(row);
      byOffice.set(row.officeName, rows);
    }

    for (const office of inScopeOffices) {
      const rows = byOffice.get(office.officeName) ?? [];

      if (rows.length === 0) {
        issues.push(`Missing office network row for ${office.officeName} in ${period.reportingMonth}.`);
      }

      if (rows.length > 1) {
        issues.push(`Duplicate office network rows found for ${office.officeName} in ${period.reportingMonth}.`);
      }
    }

    const rows = inScopeOffices
      .map((office) => byOffice.get(office.officeName)?.[0])
      .filter((row): row is OfficeNetworkAvailabilityRow => Boolean(row));

    if (rows.length === 0) {
      return {
        reportingMonth: period.reportingMonth,
        availabilityPct: 0,
        outageMinutes: 0,
        majorIncidents: 0,
        perfectOffices: 0,
        below99_9Offices: 0,
        below99Offices: 0,
        worstOffice: null,
        worstAvailabilityPct: null,
      };
    }

    const availabilityPct = roundTo(rows.reduce((total, row) => total + row.availabilityPct, 0) / rows.length, 2);
    const outageMinutes = rows.reduce((total, row) => total + row.outageMinutes, 0);
    const majorIncidents = rows.reduce((total, row) => total + row.majorIncidents, 0);
    const perfectOffices = rows.filter((row) => row.availabilityPct === 100).length;
    const below99_9Offices = rows.filter((row) => row.availabilityPct < 99.9).length;
    const below99Offices = rows.filter((row) => row.availabilityPct < 99).length;
    const worst = rows.reduce((lowest, row) => (lowest.availabilityPct <= row.availabilityPct ? lowest : row), rows[0]);

    return {
      reportingMonth: period.reportingMonth,
      availabilityPct,
      outageMinutes,
      majorIncidents,
      perfectOffices,
      below99_9Offices,
      below99Offices,
      worstOffice: worst.officeName,
      worstAvailabilityPct: worst.availabilityPct,
    };
  });
}

function buildDerivedNetworkServiceRows(metrics: DerivedNetworkMetricRow[]): ServiceAvailabilityRow[] {
  return metrics.map((metric) => ({
    reportingMonth: metric.reportingMonth,
    serviceName: NETWORK_SERVICE_NAME,
    serviceType: NETWORK_SERVICE_NAME,
    availabilityPct: metric.availabilityPct,
    targetPct: NETWORK_TARGET_PCT,
    outageMinutes: metric.outageMinutes,
    majorIncidents: metric.majorIncidents,
    backupSuccessPct: null,
    restoreTestStatus: "",
    commentary: metric.worstOffice
      ? `${metric.perfectOffices} offices at 100%. Worst office: ${metric.worstOffice} (${metric.worstAvailabilityPct?.toFixed(2)}%).`
      : "No in-scope office network data available.",
  }));
}

function sortByReportingMonth<T extends { reportingMonth: string }>(rows: T[]): T[] {
  return [...rows].sort((left, right) => left.reportingMonth.localeCompare(right.reportingMonth));
}

function validatePortfolioGanttRows(
  periods: PeriodRow[],
  workstreams: PortfolioGanttWorkstreamRow[],
  milestones: PortfolioGanttMilestoneRow[],
  issues: string[],
): void {
  const validMonths = new Set(periods.map((period) => period.reportingMonth));
  const workstreamKeys = new Set(workstreams.map((row) => `${row.reportingMonth}::${row.workstreamName}`));

  for (const row of workstreams) {
    if (!validMonths.has(row.reportingMonth)) {
      issues.push(`Portfolio Gantt workstream "${row.workstreamName}" references unknown month ${row.reportingMonth}.`);
    }
  }

  for (const row of milestones) {
    if (!validMonths.has(row.reportingMonth)) {
      issues.push(`Portfolio Gantt milestone "${row.milestoneLabel}" references unknown month ${row.reportingMonth}.`);
    }

    if (!workstreamKeys.has(`${row.reportingMonth}::${row.workstreamName}`)) {
      issues.push(
        `Portfolio Gantt milestone "${row.milestoneLabel}" does not match a workstream named "${row.workstreamName}" in ${row.reportingMonth}.`,
      );
    }
  }
}

export async function parseWorkbookBuffer(
  buffer: Buffer,
  sourceFilename: string,
  options: ParseWorkbookBufferOptions = {},
): Promise<ParseWorkbookResult> {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const issues: string[] = [];

  ensureRequiredSheets(workbook, issues);

  if (issues.length > 0) {
    throw new WorkbookValidationError(issues);
  }

  const readmeMatrix = getSheetMatrix(workbook, README_SHEET);
  const metadata = findWorkbookMetadata(readmeMatrix);

  if (metadata.templateKey !== WORKBOOK_TEMPLATE_KEY) {
    issues.push(`README Template Key must equal ${WORKBOOK_TEMPLATE_KEY}.`);
  }

  if (metadata.templateVersion !== WORKBOOK_TEMPLATE_VERSION) {
    issues.push(`README Template Version must equal ${WORKBOOK_TEMPLATE_VERSION}.`);
  }

  const tableMap = await extractTableMap(buffer);

  if (!options.skipTableValidation) {
    validateTables(tableMap, issues);
  }

  const parsedSheets = Object.fromEntries(
    SHEET_CONTRACTS.map((contract) => [contract.sheetName, sheetMatrixToRecords(contract.sheetName, getSheetMatrix(workbook, contract.sheetName), contract.headers)]),
  );

  const periods = sortByReportingMonth(parsePeriods(parsedSheets.Periods));
  const entities = parseEntities(parsedSheets.Entities);
  const officeLocations = parseOfficeLocations(parsedSheets.Office_Locations).sort((left, right) => left.displayOrder - right.displayOrder);
  const officeNetworkAvailability = sortByReportingMonth(parseOfficeNetworkAvailability(parsedSheets[OFFICE_NETWORK_SHEET_NAME]));
  const serviceAvailabilityInput = parseServiceAvailability(parsedSheets.INPUT_Service_Availability);
  const portfolioGanttWorkstreams = sortByReportingMonth(parsePortfolioGanttWorkstreams(parsedSheets[PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME])).sort(
    (left, right) => left.reportingMonth.localeCompare(right.reportingMonth) || left.displayOrder - right.displayOrder || left.workstreamName.localeCompare(right.workstreamName),
  );
  const portfolioGanttMilestones = sortByReportingMonth(parsePortfolioGanttMilestones(parsedSheets[PORTFOLIO_GANTT_MILESTONES_SHEET_NAME])).sort(
    (left, right) =>
      left.reportingMonth.localeCompare(right.reportingMonth) ||
      left.displayOrder - right.displayOrder ||
      left.workstreamName.localeCompare(right.workstreamName) ||
      left.milestoneDate.localeCompare(right.milestoneDate),
  );

  if (serviceAvailabilityInput.some((row) => row.serviceName === NETWORK_SERVICE_NAME)) {
    issues.push(`INPUT_Service_Availability must not contain manual "${NETWORK_SERVICE_NAME}" rows in template v3.`);
  }

  const derivedNetworkMetrics = deriveNetworkMetrics(periods, officeLocations, officeNetworkAvailability, issues);
  const derivedNetworkServiceRows = buildDerivedNetworkServiceRows(derivedNetworkMetrics);
  validatePortfolioGanttRows(periods, portfolioGanttWorkstreams, portfolioGanttMilestones, issues);
  const serviceAvailability = sortByReportingMonth([
    ...serviceAvailabilityInput.filter((row) => row.serviceName !== NETWORK_SERVICE_NAME),
    ...derivedNetworkServiceRows,
  ]);

  const currentPeriod = periods.find((period) => period.isCurrentPeriod) ?? periods.at(-1);

  if (!currentPeriod) {
    issues.push("Periods must contain at least one reporting month.");
  }

  const availableMonths = periods.map((period) => period.reportingMonth);

  const snapshot: NormalizedReportSnapshot = {
    metadata: {
      templateKey: metadata.templateKey,
      templateVersion: metadata.templateVersion,
      sourceFilename,
      tableMap,
    },
    availableMonths,
    currentMonth: currentPeriod?.reportingMonth ?? "",
    periods,
    entities,
    officeLocations,
    officeNetworkAvailability,
    serviceAvailability,
    supportOperations: sortByReportingMonth(parseSupportOperations(parsedSheets.INPUT_Support_Operations)),
    oldestTickets: sortByReportingMonth(parseOldestTickets(parsedSheets.INPUT_Top_Oldest_Tickets)),
    securityPatching: sortByReportingMonth(parseSecurityPatching(parsedSheets.INPUT_Security_Patching)),
    assetsLifecycle: sortByReportingMonth(parseAssetsLifecycle(parsedSheets.INPUT_Assets_Lifecycle)),
    changeRelease: sortByReportingMonth(parseChangeRelease(parsedSheets.INPUT_Change_Release)),
    devDelivery: sortByReportingMonth(parseDevDelivery(parsedSheets.INPUT_Dev_Delivery)),
    projectPortfolio: sortByReportingMonth(parseProjectPortfolio(parsedSheets.INPUT_Project_Portfolio)),
    rollingRoadmap: parseRollingRoadmap(parsedSheets.INPUT_Rolling_Roadmap),
    portfolioGanttWorkstreams,
    portfolioGanttMilestones,
    budgetCommercials: sortByReportingMonth(parseBudgetCommercials(parsedSheets.INPUT_Budget_Commercials)),
    topRisks: sortByReportingMonth(parseTopRisks(parsedSheets.INPUT_Top_Risks)),
    narrativeNotes: sortByReportingMonth(parseNarrativeNotes(parsedSheets.INPUT_Narrative_Notes)),
    derivedNetworkMetrics,
  };

  if (issues.length > 0) {
    throw new WorkbookValidationError(issues);
  }

  return {
    snapshot,
    issues,
  };
}
