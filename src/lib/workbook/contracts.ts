export const WORKBOOK_TEMPLATE_KEY = "IT_EXEC_TEMPLATE_V4";
export const WORKBOOK_TEMPLATE_VERSION = 4;
export const LEGACY_WORKBOOK_TEMPLATES = [
  { key: "IT_EXEC_TEMPLATE_V3", version: 3 },
] as const;
export const OFFICE_NETWORK_SHEET_NAME = "INPUT_Office_Network_Avail";
export const PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME = "INPUT_Gantt_Workstreams";
export const PORTFOLIO_GANTT_MILESTONES_SHEET_NAME = "INPUT_Gantt_Milestones";
export const CHART_SETTINGS_SHEET_NAME = "INPUT_Chart_Settings";
export const CHART_SETTING_PAGE_OPTIONS = ["Support Operations"] as const;
export const CHART_SETTING_KEYS = ["support_ticket_volumes"] as const;
export const CHART_OVERLAY_METRICS = ["Close Balance %"] as const;
export const PORTFOLIO_GANTT_DOMAINS = [
  "Infrastructure",
  "End-user computing",
  "Security & compliance",
  "Applications & data",
  "Product / development",
  "Business transformation",
] as const;

export interface SheetContract {
  sheetName: string;
  headers: string[];
  tableName?: string;
}

const BASE_SHEET_CONTRACTS: SheetContract[] = [
  {
    sheetName: "Periods",
    headers: ["Reporting Month", "Month End Date", "Quarter", "Financial Year", "Is Current Period", "Report Cut-Off Date"],
  },
  {
    sheetName: "Entities",
    headers: ["Entity Type", "Entity Name", "Grouping", "In Scope", "Notes"],
  },
  {
    sheetName: "Office_Locations",
    headers: ["Office Name", "Region", "In Scope", "Display Order", "Map X", "Map Y"],
    tableName: "TOfficeLocations",
  },
  {
    sheetName: OFFICE_NETWORK_SHEET_NAME,
    headers: ["Reporting Month", "Office Name", "Availability %", "Outage Minutes", "Major Incidents", "Commentary"],
    tableName: "TOfficeNetworkAvailability",
  },
  {
    sheetName: "INPUT_Service_Availability",
    headers: [
      "Reporting Month",
      "Service Name",
      "Service Type",
      "Availability %",
      "Target %",
      "Outage Minutes",
      "Major Incidents",
      "Backup Success %",
      "Restore Test Status",
      "Commentary",
    ],
    tableName: "TServiceAvailability",
  },
  {
    sheetName: "INPUT_Support_Operations",
    headers: [
      "Reporting Month",
      "Tickets Opened",
      "Tickets Closed",
      "Backlog End",
      "Average Age Open Days",
      "Average Resolution Days",
      "First Response SLA %",
      "Resolution SLA %",
      "Reopen Rate %",
      "Major Incidents",
      "Ticket CSAT",
      "CSAT Response Rate %",
      "Top Category",
      "Commentary",
    ],
    tableName: "TSupportOperations",
  },
  {
    sheetName: "INPUT_Top_Oldest_Tickets",
    headers: [
      "Reporting Month",
      "Ticket ID",
      "Title",
      "Category",
      "Owner / Queue",
      "Current Status",
      "Age Days",
      "Business Critical?",
      "Blocker / Reason Still Open",
      "Target Resolution Date",
      "Next Action",
    ],
    tableName: "TTopOldestTickets",
  },
  {
    sheetName: "INPUT_Security_Patching",
    headers: [
      "Reporting Month",
      "Workstation Patch Compliance %",
      "Server Patch Compliance %",
      "Critical Patch Compliance %",
      "Devices Outside Policy",
      "Critical Vulns",
      "High Vulns",
      "Medium Vulns",
      "Low Vulns",
      "Security Incidents",
      "MFA Coverage %",
      "Endpoint Coverage %",
      "Overdue Remediation Items",
      "Commentary",
    ],
    tableName: "TSecurityPatching",
  },
  {
    sheetName: "INPUT_Assets_Lifecycle",
    headers: [
      "Reporting Month",
      "Asset Type",
      "Active Devices",
      "Average Age Months",
      "% Within Lifecycle",
      "% Out of Lifecycle",
      "Stock On Hand",
      "Months Stock Cover",
      "Awaiting Deployment",
      "Awaiting Disposal",
      "Refresh Spend",
      "Incidents Linked to Aged Kit",
      "Commentary",
    ],
    tableName: "TAssetsLifecycle",
  },
  {
    sheetName: "INPUT_Change_Release",
    headers: [
      "Reporting Month",
      "Total Changes",
      "Standard Changes",
      "Normal Changes",
      "Emergency Changes",
      "Successful Changes",
      "Failed Changes",
      "Rolled Back Changes",
      "Change Success Rate %",
      "Changes Causing Incidents",
      "Releases Deployed",
      "Planned Maintenance Completed",
      "Commentary",
    ],
    tableName: "TChangeRelease",
  },
  {
    sheetName: "INPUT_Dev_Delivery",
    headers: [
      "Reporting Month",
      "Dev Tasks Opened",
      "Dev Tasks Closed",
      "Dev Backlog End",
      "Average Dev Task Age Days",
      "Oldest Open Dev Task Days",
      "Blocked Items",
      "Defects Delivered",
      "Enhancements Delivered",
      "Tech Debt Delivered",
      "BAU Delivered",
      "Dev CSAT / Sponsor Score",
      "Commentary",
    ],
    tableName: "TDevDelivery",
  },
  {
    sheetName: "INPUT_Project_Portfolio",
    headers: [
      "Reporting Month",
      "Project Name",
      "Project Sponsor",
      "Status RAG",
      "Delivery Confidence %",
      "Budget Status",
      "Milestone Next",
      "Milestone Date",
      "Project CSAT",
      "Benefits / Value Delivered",
      "Blockers / Dependencies",
      "Decision Needed?",
      "Commentary",
    ],
    tableName: "TProjectPortfolio",
  },
  {
    sheetName: "INPUT_Rolling_Roadmap",
    headers: ["Roadmap Quarter", "Lane", "Initiative", "Status RAG", "Outcome / Goal", "Owner", "Dependency", "Decision Required", "Notes"],
    tableName: "TRollingRoadmap",
  },
  {
    sheetName: PORTFOLIO_GANTT_WORKSTREAMS_SHEET_NAME,
    headers: [
      "Reporting Month",
      "Workstream Name",
      "Sponsor / Owner",
      "Domain",
      "Status RAG",
      "Start Date",
      "End Date",
      "Progress Date",
      "Detail / Commentary",
      "Display Order",
      "In Scope",
    ],
    tableName: "TPortfolioGanttWorkstreams",
  },
  {
    sheetName: PORTFOLIO_GANTT_MILESTONES_SHEET_NAME,
    headers: ["Reporting Month", "Workstream Name", "Milestone Label", "Milestone Date", "Display Order"],
    tableName: "TPortfolioGanttMilestones",
  },
  {
    sheetName: "INPUT_Budget_Commercials",
    headers: [
      "Reporting Month",
      "Budget Line",
      "Budget Amount",
      "Actual Amount",
      "Forecast Amount",
      "Variance",
      "Cloud / Licensing Spend",
      "Asset Refresh Spend",
      "Savings Achieved",
      "Avoidable Cost Risk",
      "Vendor / Contract",
      "Renewal Due Date",
      "Renewal Value",
      "Owner",
      "Commentary",
    ],
    tableName: "TBudgetCommercials",
  },
  {
    sheetName: "INPUT_Top_Risks",
    headers: [
      "Reporting Month",
      "Risk / Issue",
      "Type",
      "Owner",
      "Impact",
      "Likelihood",
      "Rating RAG",
      "Current Control / Mitigation",
      "Target Date",
      "Decision Required?",
      "Commentary",
    ],
    tableName: "TTopRisks",
  },
  {
    sheetName: "INPUT_Narrative_Notes",
    headers: ["Reporting Month", "Section", "Note Type", "Headline", "Narrative", "Owner"],
    tableName: "TNarrativeNotes",
  },
];

const V4_ONLY_SHEET_CONTRACTS: SheetContract[] = [
  {
    sheetName: CHART_SETTINGS_SHEET_NAME,
    headers: [
      "Reporting Month",
      "Page",
      "Chart Key",
      "Overlay Enabled",
      "Overlay Metric",
      "Rolling Window",
      "Healthy Min",
      "Amber Min",
      "Commentary",
    ],
    tableName: "TChartSettings",
  },
];

export function getSheetContractsForVersion(version: number): SheetContract[] {
  if (version >= WORKBOOK_TEMPLATE_VERSION) {
    return [...BASE_SHEET_CONTRACTS, ...V4_ONLY_SHEET_CONTRACTS];
  }

  return [...BASE_SHEET_CONTRACTS];
}

export const SHEET_CONTRACTS = getSheetContractsForVersion(WORKBOOK_TEMPLATE_VERSION);
export const REQUIRED_SHEET_NAMES = SHEET_CONTRACTS.map((sheet) => sheet.sheetName);
