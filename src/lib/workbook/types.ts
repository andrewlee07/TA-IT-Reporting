export interface PeriodRow {
  reportingMonth: string;
  monthEndDate: string;
  quarter: string;
  financialYear: string;
  isCurrentPeriod: boolean;
  reportCutOffDate: string;
}

export interface EntityRow {
  entityType: string;
  entityName: string;
  grouping: string;
  inScope: boolean;
  notes: string;
}

export interface OfficeLocationRow {
  officeName: string;
  region: string;
  inScope: boolean;
  displayOrder: number;
  mapX: number;
  mapY: number;
}

export interface OfficeNetworkAvailabilityRow {
  reportingMonth: string;
  officeName: string;
  availabilityPct: number;
  outageMinutes: number;
  majorIncidents: number;
  commentary: string;
}

export interface ServiceAvailabilityRow {
  reportingMonth: string;
  serviceName: string;
  serviceType: string;
  availabilityPct: number;
  targetPct: number;
  outageMinutes: number;
  majorIncidents: number;
  backupSuccessPct: number | null;
  restoreTestStatus: string;
  commentary: string;
}

export interface SupportOperationsRow {
  reportingMonth: string;
  ticketsOpened: number;
  ticketsClosed: number;
  backlogEnd: number;
  averageAgeOpenDays: number;
  averageResolutionDays: number;
  firstResponseSlaPct: number;
  resolutionSlaPct: number;
  reopenRatePct: number;
  majorIncidents: number;
  ticketCsatScore: number;
  csatResponseRatePct: number;
  topCategory: string;
  commentary: string;
}

export interface OldestTicketRow {
  reportingMonth: string;
  ticketId: string;
  title: string;
  category: string;
  ownerQueue: string;
  currentStatus: string;
  ageDays: number;
  businessCritical: boolean;
  blockerReason: string;
  targetResolutionDate: string;
  nextAction: string;
}

export interface SecurityPatchingRow {
  reportingMonth: string;
  workstationPatchCompliancePct: number;
  serverPatchCompliancePct: number;
  criticalPatchCompliancePct: number;
  devicesOutsidePolicy: number;
  criticalVulns: number;
  highVulns: number;
  mediumVulns: number;
  lowVulns: number;
  securityIncidents: number;
  mfaCoveragePct: number;
  endpointCoveragePct: number;
  overdueRemediationItems: number;
  commentary: string;
}

export interface AssetsLifecycleRow {
  reportingMonth: string;
  assetType: string;
  activeDevices: number;
  averageAgeMonths: number;
  withinLifecyclePct: number;
  outOfLifecyclePct: number;
  stockOnHand: number;
  monthsStockCover: number;
  awaitingDeployment: number;
  awaitingDisposal: number;
  refreshSpend: number;
  incidentsLinkedToAgedKit: number;
  commentary: string;
}

export interface ChangeReleaseRow {
  reportingMonth: string;
  totalChanges: number;
  standardChanges: number;
  normalChanges: number;
  emergencyChanges: number;
  successfulChanges: number;
  failedChanges: number;
  rolledBackChanges: number;
  changeSuccessRatePct: number;
  changesCausingIncidents: number;
  releasesDeployed: number;
  plannedMaintenanceCompleted: number;
  commentary: string;
}

export interface DevDeliveryRow {
  reportingMonth: string;
  devTasksOpened: number;
  devTasksClosed: number;
  devBacklogEnd: number;
  averageDevTaskAgeDays: number;
  oldestOpenDevTaskDays: number;
  blockedItems: number;
  defectsDelivered: number;
  enhancementsDelivered: number;
  techDebtDelivered: number;
  bauDelivered: number;
  devCsatScore: number;
  commentary: string;
}

export interface ProjectPortfolioRow {
  reportingMonth: string;
  projectName: string;
  projectSponsor: string;
  statusRag: string;
  deliveryConfidencePct: number;
  budgetStatus: string;
  milestoneNext: string;
  milestoneDate: string;
  projectCsatScore: number;
  benefitsValueDelivered: string;
  blockersDependencies: string;
  decisionNeeded: boolean;
  commentary: string;
}

export interface RollingRoadmapRow {
  roadmapQuarter: string;
  lane: string;
  initiative: string;
  statusRag: string;
  outcomeGoal: string;
  owner: string;
  dependency: string;
  decisionRequired: boolean;
  notes: string;
}

export interface PortfolioGanttWorkstreamRow {
  reportingMonth: string;
  workstreamName: string;
  sponsorOwner: string;
  domain: string;
  statusRag: string;
  startDate: string;
  endDate: string;
  progressDate: string | null;
  detailCommentary: string;
  displayOrder: number;
  inScope: boolean;
}

export interface PortfolioGanttMilestoneRow {
  reportingMonth: string;
  workstreamName: string;
  milestoneLabel: string;
  milestoneDate: string;
  displayOrder: number;
}

export interface ChartSettingRow {
  reportingMonth: string;
  page: string;
  chartKey: string;
  overlayEnabled: boolean;
  overlayMetric: string;
  rollingWindow: number;
  healthyMin: number;
  amberMin: number;
  commentary: string;
}

export interface BudgetCommercialRow {
  reportingMonth: string;
  budgetLine: string;
  budgetAmount: number;
  actualAmount: number;
  forecastAmount: number;
  variance: number;
  cloudLicensingSpend: number;
  assetRefreshSpend: number;
  savingsAchieved: number;
  avoidableCostRisk: number;
  vendorContract: string;
  renewalDueDate: string;
  renewalValue: number;
  owner: string;
  commentary: string;
}

export interface TopRiskRow {
  reportingMonth: string;
  riskIssue: string;
  type: string;
  owner: string;
  impact: string;
  likelihood: string;
  ratingRag: string;
  currentControlMitigation: string;
  targetDate: string;
  decisionRequired: boolean;
  commentary: string;
}

export interface NarrativeNoteRow {
  reportingMonth: string;
  section: string;
  noteType: string;
  headline: string;
  narrative: string;
  owner: string;
}

export interface DerivedNetworkMetricRow {
  reportingMonth: string;
  availabilityPct: number;
  outageMinutes: number;
  majorIncidents: number;
  perfectOffices: number;
  below99_9Offices: number;
  below99Offices: number;
  worstOffice: string | null;
  worstAvailabilityPct: number | null;
}

export interface WorkbookMetadata {
  templateKey: string;
  templateVersion: number;
  sourceFilename: string;
  tableMap: Record<string, string[]>;
}

export interface NormalizedReportSnapshot {
  metadata: WorkbookMetadata;
  availableMonths: string[];
  currentMonth: string;
  periods: PeriodRow[];
  entities: EntityRow[];
  officeLocations: OfficeLocationRow[];
  officeNetworkAvailability: OfficeNetworkAvailabilityRow[];
  serviceAvailability: ServiceAvailabilityRow[];
  supportOperations: SupportOperationsRow[];
  oldestTickets: OldestTicketRow[];
  securityPatching: SecurityPatchingRow[];
  assetsLifecycle: AssetsLifecycleRow[];
  changeRelease: ChangeReleaseRow[];
  devDelivery: DevDeliveryRow[];
  projectPortfolio: ProjectPortfolioRow[];
  rollingRoadmap: RollingRoadmapRow[];
  portfolioGanttWorkstreams: PortfolioGanttWorkstreamRow[];
  portfolioGanttMilestones: PortfolioGanttMilestoneRow[];
  chartSettings: ChartSettingRow[];
  budgetCommercials: BudgetCommercialRow[];
  topRisks: TopRiskRow[];
  narrativeNotes: NarrativeNoteRow[];
  derivedNetworkMetrics: DerivedNetworkMetricRow[];
}

export interface ParseWorkbookResult {
  snapshot: NormalizedReportSnapshot;
  issues: string[];
}

export class WorkbookValidationError extends Error {
  constructor(public readonly issues: string[]) {
    super("Workbook validation failed.");
    this.name = "WorkbookValidationError";
  }
}
