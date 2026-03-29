import type { SectionId } from "@/lib/drafts/types";
import { NETWORK_SERVICE_NAME } from "@/lib/workbook/derived-network";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

export type CollectionKey =
  | "periods"
  | "entities"
  | "officeLocations"
  | "chartSettings"
  | "serviceAvailability"
  | "officeNetworkAvailability"
  | "derivedNetworkMetrics"
  | "supportOperations"
  | "oldestTickets"
  | "securityPatching"
  | "assetsLifecycle"
  | "changeRelease"
  | "devDelivery"
  | "projectPortfolio"
  | "rollingRoadmap"
  | "portfolioGanttWorkstreams"
  | "portfolioGanttMilestones"
  | "budgetCommercials"
  | "topRisks"
  | "narrativeNotes";

export type FieldType = "text" | "textarea" | "number" | "date" | "checkbox";

export interface FieldDefinition {
  key: string;
  label: string;
  type: FieldType;
  step?: string;
}

export interface CollectionConfig {
  key: CollectionKey;
  label: string;
  description: string;
  layout: "table" | "single" | "readonly";
  monthScoped: boolean;
  getRows: (snapshot: NormalizedReportSnapshot, selectedMonth: string) => Array<Record<string, unknown>>;
  setRows?: (snapshot: NormalizedReportSnapshot, selectedMonth: string, rows: Array<Record<string, unknown>>) => NormalizedReportSnapshot;
  createRow?: (selectedMonth: string) => Record<string, unknown>;
}

export interface EditorSectionConfig {
  id: SectionId;
  label: string;
  description: string;
  collections: CollectionConfig[];
}

function humanizeLabel(value: string): string {
  return value
    .replace(/([A-Z])/g, " $1")
    .replace(/^./, (first) => first.toUpperCase())
    .trim()
    .replace(/Pct/g, "%")
    .replace(/Csat/g, "CSAT")
    .replace(/Mfa/g, "MFA")
    .replace(/Bau/g, "BAU");
}

const TEXTAREA_FIELDS = new Set([
  "commentary",
  "narrative",
  "notes",
  "headline",
  "detailCommentary",
  "blockersDependencies",
  "currentControlMitigation",
  "benefitsValueDelivered",
  "nextAction",
  "blockerReason",
]);

const CHECKBOX_FIELDS = new Set([
  "isCurrentPeriod",
  "inScope",
  "overlayEnabled",
  "businessCritical",
  "decisionNeeded",
  "decisionRequired",
]);

function inferFieldType(key: string, value: unknown): FieldType {
  if (CHECKBOX_FIELDS.has(key) || typeof value === "boolean") {
    return "checkbox";
  }

  if (key.toLowerCase().includes("date")) {
    return "date";
  }

  if (TEXTAREA_FIELDS.has(key)) {
    return "textarea";
  }

  if (typeof value === "number") {
    return "number";
  }

  return "text";
}

export function buildFieldDefinitions(row: Record<string, unknown>): FieldDefinition[] {
  return Object.entries(row).map(([key, value]) => ({
    key,
    label: humanizeLabel(key),
    type: inferFieldType(key, value),
    step: typeof value === "number" && key.toLowerCase().includes("pct") ? "0.01" : "1",
  }));
}

function setCollectionRows<K extends CollectionKey>(
  snapshot: NormalizedReportSnapshot,
  key: K,
  rows: Array<Record<string, unknown>>,
): NormalizedReportSnapshot {
  return {
    ...snapshot,
    [key]: rows,
  } as NormalizedReportSnapshot;
}

function toEditorRows<T extends object>(rows: T[]): Array<Record<string, unknown>> {
  return rows as unknown as Array<Record<string, unknown>>;
}

function fromEditorRows<T>(rows: Array<Record<string, unknown>>): T[] {
  return rows as unknown as T[];
}

function filterByMonth(rows: Array<Record<string, unknown>>, selectedMonth: string): Array<Record<string, unknown>> {
  return rows.filter((row) => row.reportingMonth === selectedMonth);
}

export const EDITOR_SECTIONS_CONFIG: EditorSectionConfig[] = [
  {
    id: "overview-setup",
    label: "Overview & Setup",
    description: "Core report metadata, reporting periods, org scope, office map, and chart settings.",
    collections: [
      {
        key: "periods",
        label: "Reporting Periods",
        description: "Manage the rolling months included in this report.",
        layout: "table",
        monthScoped: false,
        getRows: (snapshot) => toEditorRows(snapshot.periods),
        setRows: (snapshot, _month, rows) => {
          const periods = fromEditorRows<typeof snapshot.periods[number]>(rows).sort((left, right) =>
            left.reportingMonth.localeCompare(right.reportingMonth),
          );
          const availableMonths = Array.from(new Set(periods.map((row) => row.reportingMonth)));
          return {
            ...snapshot,
            periods,
            availableMonths,
            currentMonth: availableMonths.includes(snapshot.currentMonth)
              ? snapshot.currentMonth
              : availableMonths[availableMonths.length - 1] ?? snapshot.currentMonth,
          };
        },
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
          monthEndDate: `${selectedMonth}-28`,
          quarter: "Q1",
          financialYear: "",
          isCurrentPeriod: false,
          reportCutOffDate: `${selectedMonth}-28`,
        }),
      },
      {
        key: "entities",
        label: "Entities",
        description: "Define in-scope entities and groupings.",
        layout: "table",
        monthScoped: false,
        getRows: (snapshot) => toEditorRows(snapshot.entities),
        setRows: (snapshot, _month, rows) => setCollectionRows(snapshot, "entities", rows),
        createRow: () => ({ entityType: "", entityName: "", grouping: "", inScope: true, notes: "" }),
      },
      {
        key: "officeLocations",
        label: "Office Locations",
        description: "Manage in-scope office reference data for map and network rollups.",
        layout: "table",
        monthScoped: false,
        getRows: (snapshot) => toEditorRows(snapshot.officeLocations),
        setRows: (snapshot, _month, rows) => setCollectionRows(snapshot, "officeLocations", rows),
        createRow: () => ({ officeName: "", region: "", inScope: true, displayOrder: 0, mapX: 0, mapY: 0 }),
      },
      {
        key: "chartSettings",
        label: "Chart Settings",
        description: "Optional chart overlays and thresholds.",
        layout: "table",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.chartSettings), selectedMonth),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          chartSettings: [
            ...snapshot.chartSettings.filter((row) => row.reportingMonth !== selectedMonth),
            ...fromEditorRows<typeof snapshot.chartSettings[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
          page: "Support Operations",
          chartKey: "support_ticket_volumes",
          overlayEnabled: true,
          overlayMetric: "Close Balance %",
          rollingWindow: 3,
          healthyMin: 100,
          amberMin: 97,
          commentary: "",
        }),
      },
    ],
  },
  {
    id: "availability-network",
    label: "Availability & Network",
    description: "Service availability, office network performance, and derived rollups.",
    collections: [
      {
        key: "serviceAvailability",
        label: "Service Availability",
        description: "Month-scoped service rows excluding the derived Network service.",
        layout: "table",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) =>
          toEditorRows(snapshot.serviceAvailability).filter(
            (row) => row.reportingMonth === selectedMonth && row.serviceName !== NETWORK_SERVICE_NAME,
          ),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          serviceAvailability: [
            ...snapshot.serviceAvailability.filter(
              (row) => row.reportingMonth !== selectedMonth || row.serviceName === NETWORK_SERVICE_NAME,
            ),
            ...fromEditorRows<typeof snapshot.serviceAvailability[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
          serviceName: "",
          serviceType: "",
          availabilityPct: 0,
          targetPct: 99.9,
          outageMinutes: 0,
          majorIncidents: 0,
          backupSuccessPct: null,
          restoreTestStatus: "",
          commentary: "",
        }),
      },
      {
        key: "officeNetworkAvailability",
        label: "Office Network Availability",
        description: "Per-office inputs that drive the derived network metrics.",
        layout: "table",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.officeNetworkAvailability), selectedMonth),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          officeNetworkAvailability: [
            ...snapshot.officeNetworkAvailability.filter((row) => row.reportingMonth !== selectedMonth),
            ...fromEditorRows<typeof snapshot.officeNetworkAvailability[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
          officeName: "",
          availabilityPct: 0,
          outageMinutes: 0,
          majorIncidents: 0,
          commentary: "",
        }),
      },
      {
        key: "derivedNetworkMetrics",
        label: "Derived Network Metrics",
        description: "Read-only rollups recalculated from office network rows.",
        layout: "readonly",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.derivedNetworkMetrics), selectedMonth),
      },
    ],
  },
  {
    id: "support-operations",
    label: "Support Operations",
    description: "Operational support KPIs and oldest tickets.",
    collections: [
      {
        key: "supportOperations",
        label: "Support Overview",
        description: "Single monthly support row for the selected period.",
        layout: "single",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.supportOperations), selectedMonth),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          supportOperations: [
            ...snapshot.supportOperations.filter((row) => row.reportingMonth !== selectedMonth),
            ...fromEditorRows<typeof snapshot.supportOperations[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
          ticketsOpened: 0,
          ticketsClosed: 0,
          backlogEnd: 0,
          averageAgeOpenDays: 0,
          averageResolutionDays: 0,
          firstResponseSlaPct: 0,
          resolutionSlaPct: 0,
          reopenRatePct: 0,
          majorIncidents: 0,
          ticketCsatScore: 0,
          csatResponseRatePct: 0,
          topCategory: "",
          commentary: "",
        }),
      },
      {
        key: "oldestTickets",
        label: "Oldest Tickets",
        description: "Dense ticket-level operational backlog list.",
        layout: "table",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.oldestTickets), selectedMonth),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          oldestTickets: [
            ...snapshot.oldestTickets.filter((row) => row.reportingMonth !== selectedMonth),
            ...fromEditorRows<typeof snapshot.oldestTickets[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
          ticketId: "",
          title: "",
          category: "",
          ownerQueue: "",
          currentStatus: "",
          ageDays: 0,
          businessCritical: false,
          blockerReason: "",
          targetResolutionDate: "",
          nextAction: "",
        }),
      },
    ],
  },
  {
    id: "security-assets",
    label: "Security & Assets",
    description: "Security posture and asset lifecycle.",
    collections: [
      {
        key: "securityPatching",
        label: "Security Patching",
        description: "Single monthly security posture row.",
        layout: "single",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.securityPatching), selectedMonth),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          securityPatching: [
            ...snapshot.securityPatching.filter((row) => row.reportingMonth !== selectedMonth),
            ...fromEditorRows<typeof snapshot.securityPatching[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
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
        }),
      },
      {
        key: "assetsLifecycle",
        label: "Asset Lifecycle",
        description: "Asset class rows for lifecycle, stock, and refresh tracking.",
        layout: "table",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.assetsLifecycle), selectedMonth),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          assetsLifecycle: [
            ...snapshot.assetsLifecycle.filter((row) => row.reportingMonth !== selectedMonth),
            ...fromEditorRows<typeof snapshot.assetsLifecycle[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
          assetType: "",
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
        }),
      },
    ],
  },
  {
    id: "change-delivery",
    label: "Change & Delivery",
    description: "Change performance and development delivery flow.",
    collections: [
      {
        key: "changeRelease",
        label: "Change & Release",
        description: "Single monthly operational row for change management.",
        layout: "single",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.changeRelease), selectedMonth),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          changeRelease: [
            ...snapshot.changeRelease.filter((row) => row.reportingMonth !== selectedMonth),
            ...fromEditorRows<typeof snapshot.changeRelease[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
          totalChanges: 0,
          standardChanges: 0,
          normalChanges: 0,
          emergencyChanges: 0,
          successfulChanges: 0,
          failedChanges: 0,
          rolledBackChanges: 0,
          changeSuccessRatePct: 0,
          changesCausingIncidents: 0,
          releasesDeployed: 0,
          plannedMaintenanceCompleted: 0,
          commentary: "",
        }),
      },
      {
        key: "devDelivery",
        label: "Development Delivery",
        description: "Single monthly delivery row for product/dev operations.",
        layout: "single",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.devDelivery), selectedMonth),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          devDelivery: [
            ...snapshot.devDelivery.filter((row) => row.reportingMonth !== selectedMonth),
            ...fromEditorRows<typeof snapshot.devDelivery[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
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
        }),
      },
    ],
  },
  {
    id: "projects-roadmap",
    label: "Projects & Roadmap",
    description: "Delivery portfolio, roadmap, and gantt planning.",
    collections: [
      {
        key: "projectPortfolio",
        label: "Project Portfolio",
        description: "Tracked projects, delivery confidence, and dependencies.",
        layout: "table",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.projectPortfolio), selectedMonth),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          projectPortfolio: [
            ...snapshot.projectPortfolio.filter((row) => row.reportingMonth !== selectedMonth),
            ...fromEditorRows<typeof snapshot.projectPortfolio[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
          projectName: "",
          projectSponsor: "",
          statusRag: "",
          deliveryConfidencePct: 0,
          budgetStatus: "",
          milestoneNext: "",
          milestoneDate: "",
          projectCsatScore: 0,
          benefitsValueDelivered: "",
          blockersDependencies: "",
          decisionNeeded: false,
          commentary: "",
        }),
      },
      {
        key: "rollingRoadmap",
        label: "Rolling Roadmap",
        description: "Quarter-based roadmap items and owners.",
        layout: "table",
        monthScoped: false,
        getRows: (snapshot) => toEditorRows(snapshot.rollingRoadmap),
        setRows: (snapshot, _month, rows) => setCollectionRows(snapshot, "rollingRoadmap", rows),
        createRow: () => ({
          roadmapQuarter: "",
          lane: "",
          initiative: "",
          statusRag: "",
          outcomeGoal: "",
          owner: "",
          dependency: "",
          decisionRequired: false,
          notes: "",
        }),
      },
      {
        key: "portfolioGanttWorkstreams",
        label: "Gantt Workstreams",
        description: "12-week delivery workstreams for the selected month.",
        layout: "table",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.portfolioGanttWorkstreams), selectedMonth),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          portfolioGanttWorkstreams: [
            ...snapshot.portfolioGanttWorkstreams.filter((row) => row.reportingMonth !== selectedMonth),
            ...fromEditorRows<typeof snapshot.portfolioGanttWorkstreams[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
          workstreamName: "",
          sponsorOwner: "",
          domain: "",
          statusRag: "",
          startDate: "",
          endDate: "",
          progressDate: "",
          detailCommentary: "",
          displayOrder: 0,
          inScope: true,
        }),
      },
      {
        key: "portfolioGanttMilestones",
        label: "Gantt Milestones",
        description: "Milestones associated with the selected month workstreams.",
        layout: "table",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.portfolioGanttMilestones), selectedMonth),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          portfolioGanttMilestones: [
            ...snapshot.portfolioGanttMilestones.filter((row) => row.reportingMonth !== selectedMonth),
            ...fromEditorRows<typeof snapshot.portfolioGanttMilestones[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
          workstreamName: "",
          milestoneLabel: "",
          milestoneDate: "",
          displayOrder: 0,
        }),
      },
    ],
  },
  {
    id: "finance-risks",
    label: "Finance & Risks",
    description: "Financial tracking, renewals, and top risks.",
    collections: [
      {
        key: "budgetCommercials",
        label: "Budget & Commercials",
        description: "Budget lines, renewals, and value tracking.",
        layout: "table",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.budgetCommercials), selectedMonth),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          budgetCommercials: [
            ...snapshot.budgetCommercials.filter((row) => row.reportingMonth !== selectedMonth),
            ...fromEditorRows<typeof snapshot.budgetCommercials[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
          budgetLine: "",
          budgetAmount: 0,
          actualAmount: 0,
          forecastAmount: 0,
          variance: 0,
          cloudLicensingSpend: 0,
          assetRefreshSpend: 0,
          savingsAchieved: 0,
          avoidableCostRisk: 0,
          vendorContract: "",
          renewalDueDate: "",
          renewalValue: 0,
          owner: "",
          commentary: "",
        }),
      },
      {
        key: "topRisks",
        label: "Top Risks",
        description: "Governance risks and leadership decisions.",
        layout: "table",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.topRisks), selectedMonth),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          topRisks: [
            ...snapshot.topRisks.filter((row) => row.reportingMonth !== selectedMonth),
            ...fromEditorRows<typeof snapshot.topRisks[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
          riskIssue: "",
          type: "",
          owner: "",
          impact: "",
          likelihood: "",
          ratingRag: "",
          currentControlMitigation: "",
          targetDate: "",
          decisionRequired: false,
          commentary: "",
        }),
      },
    ],
  },
  {
    id: "notes-narrative",
    label: "Notes & Narrative",
    description: "Workbook-backed narrative inputs for highlights and commentary surfaces.",
    collections: [
      {
        key: "narrativeNotes",
        label: "Narrative Notes",
        description: "Section-level narrative cards and authored notes.",
        layout: "table",
        monthScoped: true,
        getRows: (snapshot, selectedMonth) => filterByMonth(toEditorRows(snapshot.narrativeNotes), selectedMonth),
        setRows: (snapshot, selectedMonth, rows) => ({
          ...snapshot,
          narrativeNotes: [
            ...snapshot.narrativeNotes.filter((row) => row.reportingMonth !== selectedMonth),
            ...fromEditorRows<typeof snapshot.narrativeNotes[number]>(rows),
          ],
        }),
        createRow: (selectedMonth) => ({
          reportingMonth: selectedMonth,
          section: "",
          noteType: "",
          headline: "",
          narrative: "",
          owner: "",
        }),
      },
    ],
  },
];

export function getEditorSectionConfig(sectionId: SectionId): EditorSectionConfig {
  return EDITOR_SECTIONS_CONFIG.find((section) => section.id === sectionId) ?? EDITOR_SECTIONS_CONFIG[0];
}
