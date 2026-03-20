export interface ReportPageDefinition {
  id: string;
  label: string;
}

export interface ReportBlockDefinition {
  id: string;
  label: string;
}

export const REPORT_PAGES: ReportPageDefinition[] = [
  { id: "p-exec", label: "Executive Scorecard" },
  { id: "p-avail", label: "Service Availability" },
  { id: "p-network", label: "Network & Offices" },
  { id: "p-support", label: "Support Operations" },
  { id: "p-security", label: "Security & Patching" },
  { id: "p-assets", label: "Assets & Lifecycle" },
  { id: "p-change", label: "Change & Release" },
  { id: "p-dev", label: "Development & Delivery" },
  { id: "p-projects", label: "Project Portfolio" },
  { id: "p-roadmap", label: "Rolling Roadmap" },
  { id: "p-budget", label: "Budget & Commercials" },
  { id: "p-risks", label: "Risks & Decisions" },
];

export const REPORT_BLOCKS: Record<string, ReportBlockDefinition[]> = {
  "p-exec": [
    { id: "exec-kpi-support-sla", label: "Support SLA KPI" },
    { id: "exec-kpi-user-csat", label: "User CSAT KPI" },
    { id: "exec-kpi-critical-vulns", label: "Critical vulnerabilities KPI" },
    { id: "exec-kpi-change-success", label: "Change success KPI" },
    { id: "exec-kpi-dev-backlog", label: "Dev backlog KPI" },
  ],
  "p-avail": [
    { id: "avail-trend-block", label: "Availability trend chart" },
    { id: "avail-outage-block", label: "Outage minutes by service" },
    { id: "avail-note-block", label: "Availability analyst note" },
  ],
  "p-network": [
    { id: "net-kpi-avg-availability", label: "Average availability KPI" },
    { id: "net-kpi-total-offices", label: "Total offices KPI" },
    { id: "net-kpi-below-99", label: "Below 99 percent KPI" },
    { id: "net-kpi-below-99-9", label: "Below 99.9 percent KPI" },
    { id: "network-map-block", label: "Office network map" },
    { id: "office-list-block", label: "Office availability list" },
    { id: "net-trend-block", label: "Network trend chart" },
  ],
  "p-support": [
    { id: "support-hero", label: "Support hero panel" },
    { id: "support-kpi-opened", label: "Opened KPI" },
    { id: "support-kpi-closed", label: "Closed KPI" },
    { id: "support-kpi-backlog", label: "Backlog end KPI" },
    { id: "support-kpi-avg-resolution", label: "Average resolution KPI" },
    { id: "support-kpi-major-incidents", label: "Major incidents KPI" },
    { id: "support-vol-block", label: "Ticket volume chart" },
    { id: "support-cats-block", label: "Tickets by category" },
    { id: "support-tickets-block", label: "Oldest tickets table" },
  ],
  "p-security": [
    { id: "sec-kpi-critical-vulns", label: "Critical vulnerabilities KPI" },
    { id: "sec-kpi-workstation-patch", label: "Workstation patch KPI" },
    { id: "sec-kpi-mfa-coverage", label: "MFA coverage KPI" },
    { id: "sec-kpi-overdue-remediation", label: "Overdue remediation KPI" },
    { id: "sec-compliance-block", label: "Patch compliance bars" },
    { id: "sec-vuln-block", label: "Vulnerability trend chart" },
    { id: "sec-note-block", label: "Security note" },
  ],
  "p-assets": [
    { id: "asset-kpi-total-active-devices", label: "Total active devices KPI" },
    { id: "asset-kpi-laptops-in-lifecycle", label: "Laptops in lifecycle KPI" },
    { id: "asset-kpi-laptop-incidents", label: "Laptop incidents KPI" },
    { id: "asset-kpi-stock-cover", label: "Stock cover KPI" },
    { id: "asset-tile-laptop", label: "Laptop lifecycle tile" },
    { id: "asset-tile-mobile", label: "Mobile lifecycle tile" },
    { id: "asset-tile-monitor", label: "Monitor lifecycle tile" },
    { id: "asset-trend-block", label: "Lifecycle trend chart" },
    { id: "asset-spend-block", label: "Refresh spend chart" },
  ],
  "p-change": [
    { id: "change-hero", label: "Change success hero" },
    { id: "change-kpi-total-changes", label: "Total changes KPI" },
    { id: "change-kpi-releases-deployed", label: "Releases deployed KPI" },
    { id: "change-kpi-failed-changes", label: "Failed changes KPI" },
    { id: "change-kpi-incidents", label: "Changes to incidents KPI" },
    { id: "change-breakdown-block", label: "Change breakdown chart" },
  ],
  "p-dev": [
    { id: "dev-kpi-backlog-end", label: "Backlog end KPI" },
    { id: "dev-kpi-tasks-closed", label: "Tasks closed KPI" },
    { id: "dev-kpi-blocked-items", label: "Blocked items KPI" },
    { id: "dev-kpi-csat", label: "Development CSAT KPI" },
    { id: "dev-pipeline-block", label: "Backlog pipeline chart" },
    { id: "dev-mix-block", label: "Delivery mix chart" },
    { id: "dev-note-block", label: "Delivery note" },
  ],
  "p-projects": [
    { id: "projects-kpi-active-projects", label: "Active projects KPI" },
    { id: "projects-kpi-avg-confidence", label: "Average confidence KPI" },
    { id: "projects-kpi-decisions-needed", label: "Decisions needed KPI" },
    { id: "prj-note-block", label: "Portfolio note" },
  ],
  "p-roadmap": [{ id: "roadmap-quarter-2026-q2", label: "Roadmap quarter section" }],
  "p-budget": [
    { id: "budget-kpi-total-budget", label: "Total budget KPI" },
    { id: "budget-kpi-total-actual", label: "Total actual KPI" },
    { id: "budget-kpi-variance", label: "Variance KPI" },
    { id: "budget-kpi-forecast", label: "Forecast KPI" },
    { id: "bud-table-block", label: "Budget table" },
    { id: "bud-trend-block", label: "Budget trend chart" },
    { id: "bud-renewals-block", label: "Renewals list" },
  ],
  "p-risks": [
    { id: "risk-kpi-total-risks", label: "Total risks KPI" },
    { id: "risk-kpi-decisions-needed", label: "Decisions needed KPI" },
    { id: "risk-kpi-amber-risks", label: "Amber risks KPI" },
    { id: "risk-register-block", label: "Risk register" },
    { id: "risk-note-block", label: "Governance note" },
  ],
};

export function isValidPageId(pageId: string): boolean {
  return REPORT_PAGES.some((page) => page.id === pageId);
}

export function isValidBlockId(pageId: string, blockId: string): boolean {
  return (REPORT_BLOCKS[pageId] ?? []).some((block) => block.id === blockId);
}
