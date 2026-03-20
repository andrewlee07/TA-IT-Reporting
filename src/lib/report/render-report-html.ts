import path from "node:path";

import { buildTemplateData, formatMonthLabel } from "@/lib/report/template-data";
import type { ExecSummaryState } from "@/lib/reports/exec-summary";
import { loadTemplateSource } from "@/lib/report/template-source";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

interface RenderReportHtmlOptions {
  month?: string;
  initialPageId?: string;
  showAllPages?: boolean;
  hideChrome?: boolean;
  execSummary?: ExecSummaryState;
}
const runtimePath = path.resolve(process.cwd(), "src/lib/report/runtime.js");
const officeMapPath = path.resolve(process.cwd(), "src/lib/report/office-map.js");
const chartPath = path.resolve(process.cwd(), "node_modules/chart.js/dist/chart.umd.js");

let cachedRuntime: string | null = null;
let cachedOfficeMap: string | null = null;
let cachedChartLibrary: string | null = null;

async function loadRuntimeSource(): Promise<string> {
  if (!cachedRuntime) {
    cachedRuntime = await (await import("node:fs/promises")).readFile(runtimePath, "utf8");
  }

  return cachedRuntime;
}

async function loadChartLibrarySource(): Promise<string> {
  if (!cachedChartLibrary) {
    cachedChartLibrary = await (await import("node:fs/promises")).readFile(chartPath, "utf8");
  }

  return cachedChartLibrary;
}

async function loadOfficeMapSource(): Promise<string> {
  if (!cachedOfficeMap) {
    cachedOfficeMap = await (await import("node:fs/promises")).readFile(officeMapPath, "utf8");
  }

  return cachedOfficeMap;
}

export async function renderReportHtml(snapshot: NormalizedReportSnapshot, options: RenderReportHtmlOptions = {}): Promise<string> {
  const month = options.month && snapshot.availableMonths.includes(options.month) ? options.month : snapshot.currentMonth;
  const templateData = buildTemplateData(snapshot, month, options.execSummary);
  const template = await loadTemplateSource();
  const runtime = await loadRuntimeSource();
  const officeMap = await loadOfficeMapSource();
  const chartLibrary = await loadChartLibrarySource();

  const exportOverrides = options.hideChrome
    ? `<style>
body { background: #ffffff !important; }
.sidebar { display: none !important; }
.main { padding: 0 !important; width: 100% !important; }
.report-page { display: block !important; box-shadow: none !important; margin: 0 auto 24px !important; }
@media print {
  body { background: #ffffff !important; }
  .report-page { break-after: page; page-break-after: always; }
}
</style>`
    : "";

  const safeData = JSON.stringify(templateData).replace(/</g, "\\u003c");
  const inlineOfficeMap = officeMap.replace(/^export\s+/gm, "");
  const inlineRuntime = runtime.replace(/^import\s+\{[^}]+\}\s+from\s+"\.\/office-map";\s*$/m, "");
const runtimeBootstrap = `<script type="module">
const D = ${safeData};
const ACTIVE_MONTH = '${month}';
const INITIAL_PAGE_ID = '${options.initialPageId ?? "p-summary"}';
const SHOW_ALL_PAGES = ${options.showAllPages ? "true" : "false"};
${inlineOfficeMap}
${inlineRuntime}
initReportApp(document.querySelector('.shell'), {
  data: D,
  activeMonth: ACTIVE_MONTH,
  initialPageId: INITIAL_PAGE_ID,
  showAllPages: SHOW_ALL_PAGES,
  attachGlobals: true,
});
</script>`;

  return template
    .replace(`<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>`, `<script>${chartLibrary}</script>`)
    .replace("__REPORT_OVERRIDES__", exportOverrides)
    .replace(/<script>\s*const D = __REPORT_DATA__[\s\S]*?__REPORT_RUNTIME__/, runtimeBootstrap)
    .replaceAll("June 2026", formatMonthLabel(month))
    .replaceAll("Jan – Jun 2026", templateData.meta.monthRangeLabel)
    .replaceAll("Q2 2026 – Q2 2027", templateData.meta.roadmapHorizonLabel);
}
