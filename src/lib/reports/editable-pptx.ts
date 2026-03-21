import { format, parseISO } from "date-fns";
import type { Page } from "playwright";
import PptxGenJS from "pptxgenjs";

import { getReportSlides, type ReportSlideDefinition } from "@/lib/report/blocks";
import { buildTemplateData, type TemplateData } from "@/lib/report/template-data";
import type { ExecSummaryState } from "@/lib/reports/exec-summary";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

const CONTENT_X = 0.56;
const CONTENT_Y = 1.18;
const CONTENT_W = 12.22;

const COLORS = {
  blue: "005292",
  orange: "F57D00",
  teal: "219D98",
  amber: "F59E0B",
  red: "DC5B4A",
  green: "1F9F6E",
  ink: "122033",
  slate: "667085",
  muted: "98A2B3",
  line: "D8DEE9",
  panel: "F8FAFC",
  panelAlt: "F2F5F9",
  white: "FFFFFF",
};

interface EditablePptxInput {
  page: Page;
  snapshot: NormalizedReportSnapshot;
  month: string;
  reportTitle: string;
  execSummary?: ExecSummaryState;
}

interface KpiCard {
  label: string;
  value: string;
  note?: string;
  accent?: string;
}

type TemplateRow = Record<string, string | number | boolean | null>;

function toNodeBuffer(payload: string | ArrayBuffer | Blob | Uint8Array): Buffer {
  if (typeof payload === "string") {
    return Buffer.from(payload, "binary");
  }

  if (payload instanceof Uint8Array) {
    return Buffer.from(payload);
  }

  if (payload instanceof ArrayBuffer) {
    return Buffer.from(new Uint8Array(payload));
  }

  if (typeof Blob !== "undefined" && payload instanceof Blob) {
    throw new Error("Unexpected Blob output while generating editable PPTX in Node.js.");
  }

  throw new Error("Unsupported PPTX output payload.");
}

function fmtNumber(value: number, digits = 0): string {
  return new Intl.NumberFormat("en-GB", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  }).format(value);
}

function fmtCurrency(value: number): string {
  return new Intl.NumberFormat("en-GB", {
    style: "currency",
    currency: "GBP",
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(value);
}

function fmtMonthShort(month: string): string {
  return format(parseISO(`${month}-01`), "MMM");
}

function fmtDate(dateValue: string): string {
  return format(parseISO(dateValue), "d MMM yyyy");
}

function fmtDateShort(dateValue: string): string {
  return format(parseISO(dateValue), "d MMM");
}

function parseNumeric(value: string | number | boolean | null | undefined): number {
  if (typeof value === "number") {
    return value;
  }

  if (typeof value === "boolean" || value === null || value === undefined) {
    return 0;
  }

  const cleaned = value.replace(/[^0-9.-]+/g, "");
  const parsed = Number.parseFloat(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function byMonth<T extends TemplateRow>(rows: T[], month: string): T[] {
  return rows.filter((row) => row.Month === month);
}

function firstByMonth<T extends TemplateRow>(rows: T[], month: string): T | null {
  return byMonth(rows, month)[0] ?? null;
}

function getMetaLine(slide: ReportSlideDefinition, data: TemplateData): string {
  const activeMonthLabel = data.meta.activeMonthLabel;
  switch (slide.id) {
    case "p-summary":
      return `Leadership narrative · Reporting Period · ${activeMonthLabel}`;
    case "p-exec-overview":
      return `Top service, support and delivery signals · Reporting Period · ${activeMonthLabel}`;
    case "p-exec-highlights":
      return `Operational highlights and analyst commentary · Reporting Period · ${activeMonthLabel}`;
    case "p-avail-overview":
      return `Core platform availability overview · Reporting Period · ${activeMonthLabel}`;
    case "p-avail-detail":
      return `Availability and outage trend detail · Reporting Period · ${activeMonthLabel}`;
    case "p-network-map":
      return `Office-by-office availability and estate-wide trend review · Reporting Period · ${activeMonthLabel}`;
    case "p-network-detail":
      return `Office detail and trend analysis · Reporting Period · ${activeMonthLabel}`;
    case "p-support-overview":
      return `Service desk performance and SLA position · Reporting Period · ${activeMonthLabel}`;
    case "p-support-volumes":
      return `Ticket throughput, balance and pressure · Reporting Period · ${activeMonthLabel}`;
    case "p-support-detail":
      return `Category mix and ageing ticket detail · Reporting Period · ${activeMonthLabel}`;
    case "p-security":
      return `Patching, vulnerability and control posture · Reporting Period · ${activeMonthLabel}`;
    case "p-assets":
      return `Lifecycle health, refresh demand and spend · Reporting Period · ${activeMonthLabel}`;
    case "p-change":
      return `Change quality and release throughput · Reporting Period · ${activeMonthLabel}`;
    case "p-dev":
      return `Development backlog and delivery flow · Reporting Period · ${activeMonthLabel}`;
    case "p-projects":
      return `Active project status and sponsor visibility · Reporting Period · ${activeMonthLabel}`;
    case "p-roadmap":
      return `Quarter-by-quarter forward plan · Horizon · ${data.meta.roadmapHorizonLabel}`;
    case "p-gantt": {
      const cutOff = data.meta.reportCutOffDates[data.meta.activeMonth];
      const rollingStart = cutOff ? fmtDateShort(cutOff) : data.meta.activeMonthLabel;
      return `Rolling 12-week delivery view across active workstreams · Rolling from · ${rollingStart}`;
    }
    case "p-budget":
      return `Budget control, forecast and renewals · Reporting Period · ${activeMonthLabel}`;
    case "p-risks":
      return `Top current risks, issues and decisions · Reporting Period · ${activeMonthLabel}`;
    default:
      return `Reporting Period · ${activeMonthLabel}`;
  }
}

function addHeader(slide: PptxGenJS.Slide, slideDef: ReportSlideDefinition, data: TemplateData) {
  const titleRuns = slideDef.tabLabel
    ? [
        { text: slideDef.pageLabel, options: { color: COLORS.blue, bold: true } },
        { text: " - ", options: { color: COLORS.blue, bold: false } },
        { text: slideDef.tabLabel, options: { color: COLORS.orange, bold: true } },
      ]
    : [{ text: slideDef.pageLabel, options: { color: COLORS.blue, bold: true } }];

  slide.addText(titleRuns, {
    x: CONTENT_X,
    y: 0.28,
    w: CONTENT_W,
    h: 0.42,
    fontFace: "Arial",
    fontSize: 33,
    margin: 0,
    valign: "middle",
  });

  slide.addText(getMetaLine(slideDef, data), {
    x: CONTENT_X + 0.02,
    y: 0.78,
    w: CONTENT_W,
    h: 0.16,
    fontFace: "Arial",
    fontSize: 8.5,
    color: COLORS.orange,
    margin: 0,
  });

  slide.addShape("line", {
    x: CONTENT_X,
    y: 1.02,
    w: CONTENT_W,
    h: 0,
    line: { color: COLORS.line, width: 1 },
  });
}

function addFooter(slide: PptxGenJS.Slide, slideDef: ReportSlideDefinition, index: number, total: number) {
  slide.addShape("line", {
    x: CONTENT_X,
    y: 7.02,
    w: CONTENT_W,
    h: 0,
    line: { color: COLORS.line, width: 1 },
  });

  slide.addText("TeacherActive · IT Reporting", {
    x: CONTENT_X,
    y: 7.08,
    w: 3,
    h: 0.14,
    fontFace: "Arial",
    fontSize: 8,
    color: COLORS.slate,
    margin: 0,
  });

  slide.addText(`${slideDef.slideLabel.toUpperCase()} · PAGE ${index + 1} OF ${total}`, {
    x: 9.35,
    y: 7.08,
    w: 3.45,
    h: 0.14,
    fontFace: "Arial",
    fontSize: 8,
    bold: true,
    color: COLORS.muted,
    align: "right",
    margin: 0,
  });
}

function addSectionEyebrow(slide: PptxGenJS.Slide, label: string, x: number, y: number) {
  slide.addText(label.toUpperCase(), {
    x,
    y,
    w: 3.2,
    h: 0.14,
    fontFace: "Arial",
    fontSize: 8.5,
    bold: true,
    color: COLORS.orange,
    margin: 0,
  });
}

function addPanel(
  slide: PptxGenJS.Slide,
  x: number,
  y: number,
  w: number,
  h: number,
  opts: { title?: string; subtitle?: string; accent?: string; fill?: string } = {},
) {
  const panelShape: PptxGenJS.ShapeProps = {
    x,
    y,
    w,
    h,
    fill: { color: opts.fill ?? COLORS.white },
    line: { color: COLORS.line, width: 1 },
    rectRadius: 0.08,
  };
  slide.addShape("rect", panelShape);

  const accentShape: PptxGenJS.ShapeProps = {
    x,
    y,
    w,
    h: 0.045,
    fill: { color: opts.accent ?? COLORS.blue },
    line: { color: opts.accent ?? COLORS.blue, width: 0 },
  };
  slide.addShape("rect", accentShape);

  if (opts.title) {
    slide.addText(opts.title, {
      x: x + 0.18,
      y: y + 0.18,
      w: w - 0.36,
      h: 0.18,
      fontFace: "Arial",
      fontSize: 8.8,
      bold: true,
      color: COLORS.muted,
      margin: 0,
    });
  }

  if (opts.subtitle) {
    slide.addText(opts.subtitle, {
      x: x + 0.18,
      y: y + 0.38,
      w: w - 0.36,
      h: 0.16,
      fontFace: "Arial",
      fontSize: 7.5,
      color: COLORS.slate,
      margin: 0,
    });
  }
}

function addKpiRow(slide: PptxGenJS.Slide, cards: KpiCard[], x: number, y: number, w: number, h: number) {
  const gap = 0.14;
  const cardWidth = (w - gap * (cards.length - 1)) / cards.length;

  cards.forEach((card, index) => {
    const cardX = x + index * (cardWidth + gap);
    addPanel(slide, cardX, y, cardWidth, h, {
      title: card.label,
      accent: card.accent ?? COLORS.blue,
      fill: COLORS.white,
    });

    slide.addText(card.value, {
      x: cardX + 0.18,
      y: y + 0.48,
      w: cardWidth - 0.36,
      h: 0.45,
      fontFace: "Arial",
      fontSize: 23,
      bold: true,
      color: COLORS.ink,
      margin: 0,
    });

    if (card.note) {
      slide.addText(card.note, {
        x: cardX + 0.18,
        y: y + h - 0.36,
        w: cardWidth - 0.36,
        h: 0.2,
        fontFace: "Arial",
        fontSize: 8,
        color: COLORS.slate,
        margin: 0,
      });
    }
  });
}

function addHeroMetric(
  slide: PptxGenJS.Slide,
  x: number,
  y: number,
  w: number,
  h: number,
  opts: {
    eyebrow: string;
    value: string;
    label: string;
    note?: string;
    accent?: string;
  },
) {
  addPanel(slide, x, y, w, h, { accent: opts.accent ?? COLORS.orange, fill: COLORS.white });
  addSectionEyebrow(slide, opts.eyebrow, x + 0.18, y + 0.22);
  slide.addText(opts.value, {
    x: x + 0.18,
    y: y + 0.62,
    w: w * 0.46,
    h: 0.54,
    fontFace: "Arial",
    fontSize: 33,
    bold: true,
    color: COLORS.blue,
    margin: 0,
  });
  slide.addText(opts.label, {
    x: x + 0.18,
    y: y + 1.25,
    w: w - 0.36,
    h: 0.24,
    fontFace: "Arial",
    fontSize: 11,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  if (opts.note) {
    slide.addText(opts.note, {
      x: x + 0.18,
      y: y + 1.52,
      w: w - 0.36,
      h: 0.34,
      fontFace: "Arial",
      fontSize: 9,
      color: COLORS.slate,
      margin: 0,
    });
  }
}

function addInsightBox(
  slide: PptxGenJS.Slide,
  x: number,
  y: number,
  w: number,
  h: number,
  title: string,
  body: string,
  accent = COLORS.orange,
) {
  addPanel(slide, x, y, w, h, { title, accent, fill: COLORS.panel });
  slide.addText(body, {
    x: x + 0.18,
    y: y + 0.46,
    w: w - 0.36,
    h: h - 0.6,
    fontFace: "Arial",
    fontSize: 8.8,
    color: COLORS.ink,
    margin: 0,
    valign: "top",
  });
}

function addTable(
  slide: PptxGenJS.Slide,
  x: number,
  y: number,
  w: number,
  rows: string[][],
  headers: string[],
  fontSize = 8,
) {
  const headerCells: PptxGenJS.TableCell[] = headers.map((cell) => ({
    text: cell,
    options: {
      bold: true,
      fill: { color: COLORS.panelAlt },
      color: COLORS.ink,
    },
  }));
  const bodyRows: PptxGenJS.TableCell[][] = rows.map((row) =>
    row.map((cell) => ({
      text: cell,
      options: {
        color: COLORS.ink,
      },
    })),
  );
  const tableRows = [headerCells, ...bodyRows];
  const tableOptions: PptxGenJS.TableProps = {
    x,
    y,
    w,
    rowH: 0.24,
    fill: { color: COLORS.white },
    color: COLORS.ink,
    fontFace: "Arial",
    fontSize,
    border: { pt: 1, color: COLORS.line },
    margin: 0.05,
    valign: "middle",
    bold: false,
  };
  slide.addTable(tableRows, tableOptions);
}

function addBarList(
  slide: PptxGenJS.Slide,
  entries: Array<{ label: string; value: number; note?: string; color?: string }>,
  x: number,
  y: number,
  w: number,
  h: number,
) {
  const max = Math.max(1, ...entries.map((entry) => entry.value));
  const rowHeight = h / Math.max(entries.length, 1);

  entries.forEach((entry, index) => {
    const rowY = y + index * rowHeight;
    const trackX = x + 1.5;
    const trackW = w - 2.05;
    const fillW = Math.max(0.1, (trackW * entry.value) / max);

    slide.addText(entry.label, {
      x,
      y: rowY + 0.04,
      w: 1.45,
      h: 0.16,
      fontFace: "Arial",
      fontSize: 8,
      bold: true,
      color: COLORS.ink,
      margin: 0,
    });

    if (entry.note) {
      slide.addText(entry.note, {
        x,
        y: rowY + 0.18,
        w: 1.45,
        h: 0.12,
        fontFace: "Arial",
        fontSize: 6.8,
        color: COLORS.slate,
        margin: 0,
      });
    }

    const trackShape: PptxGenJS.ShapeProps = {
      x: trackX,
      y: rowY + 0.08,
      w: trackW,
      h: 0.1,
      fill: { color: COLORS.panelAlt },
      line: { color: COLORS.panelAlt, width: 0 },
      rectRadius: 0.05,
    };
    slide.addShape("rect", trackShape);

    const fillShape: PptxGenJS.ShapeProps = {
      x: trackX,
      y: rowY + 0.08,
      w: fillW,
      h: 0.1,
      fill: { color: entry.color ?? COLORS.blue },
      line: { color: entry.color ?? COLORS.blue, width: 0 },
      rectRadius: 0.05,
    };
    slide.addShape("rect", fillShape);

    slide.addText(fmtNumber(entry.value, entry.value % 1 === 0 ? 0 : 1), {
      x: x + w - 0.4,
      y: rowY + 0.02,
      w: 0.38,
      h: 0.16,
      fontFace: "Arial",
      fontSize: 8,
      bold: true,
      color: COLORS.ink,
      align: "right",
      margin: 0,
    });
  });
}

function addLineChart(
  slide: PptxGenJS.Slide,
  x: number,
  y: number,
  w: number,
  h: number,
  labels: string[],
  series: Array<{ name: string; values: number[] }>,
  colors: string[],
  yFormat = "#,##0",
) {
  const chartOptions: PptxGenJS.IChartOpts = {
    x,
    y,
    w,
    h,
    showLegend: true,
    legendPos: "t",
    chartColors: colors,
    lineSize: 2,
    lineDataSymbol: "circle",
    lineDataSymbolSize: 4,
    valAxisLabelFormatCode: yFormat,
    valAxisLabelFontSize: 7,
    catAxisLabelFontSize: 7,
    catAxisLineColor: COLORS.line,
    valAxisLineColor: COLORS.line,
    valGridLine: { color: COLORS.line },
    catGridLine: { style: "none" },
    chartArea: { border: { color: COLORS.white, pt: 0 }, fill: { color: COLORS.white } },
    plotArea: { border: { color: COLORS.white, pt: 0 }, fill: { color: COLORS.white } },
    fontFace: "Arial",
  };
  slide.addChart(
    "line",
    series.map((item) => ({
      name: item.name,
      labels,
      values: item.values,
    })),
    chartOptions,
  );
}

function addBarChart(
  slide: PptxGenJS.Slide,
  x: number,
  y: number,
  w: number,
  h: number,
  labels: string[],
  series: Array<{ name: string; values: number[] }>,
  colors: string[],
  opts: { stacked?: boolean; yFormat?: string } = {},
) {
  const chartOptions: PptxGenJS.IChartOpts = {
    x,
    y,
    w,
    h,
    showLegend: true,
    legendPos: "t",
    chartColors: colors,
    barGrouping: opts.stacked ? "stacked" : "clustered",
    barGapWidthPct: 60,
    valAxisLabelFormatCode: opts.yFormat ?? "#,##0",
    valAxisLabelFontSize: 7,
    catAxisLabelFontSize: 7,
    catAxisLineColor: COLORS.line,
    valAxisLineColor: COLORS.line,
    valGridLine: { color: COLORS.line },
    catGridLine: { style: "none" },
    chartArea: { border: { color: COLORS.white, pt: 0 }, fill: { color: COLORS.white } },
    plotArea: { border: { color: COLORS.white, pt: 0 }, fill: { color: COLORS.white } },
    fontFace: "Arial",
  };
  slide.addChart(
    "bar",
    series.map((item) => ({
      name: item.name,
      labels,
      values: item.values,
    })),
    chartOptions,
  );
}

function addDoughnutChart(
  slide: PptxGenJS.Slide,
  x: number,
  y: number,
  w: number,
  h: number,
  labels: string[],
  values: number[],
  colors: string[],
) {
  const chartOptions: PptxGenJS.IChartOpts = {
    x,
    y,
    w,
    h,
    showLegend: true,
    legendPos: "t",
    chartColors: colors,
    holeSize: 62,
    showPercent: true,
    showLabel: true,
    dataLabelPosition: "bestFit",
    dataLabelFontSize: 8,
    chartArea: { border: { color: COLORS.white, pt: 0 }, fill: { color: COLORS.white } },
    plotArea: { border: { color: COLORS.white, pt: 0 }, fill: { color: COLORS.white } },
    fontFace: "Arial",
  };
  slide.addChart(
    "doughnut",
    [
      {
        name: "Delivery mix",
        labels,
        values,
      },
    ],
    chartOptions,
  );
}

async function captureElementImage(page: Page, selector: string): Promise<string> {
  const locator = page.locator(selector).first();
  await locator.waitFor();
  const png = await locator.screenshot({ type: "png" });
  return `data:image/png;base64,${Buffer.from(png).toString("base64")}`;
}

function execSummaryLines(data: TemplateData): string[] {
  return data.execSummary.contentHtml
    .replace(/<\/(p|h2|h3|li|ul|ol)>/gi, "\n")
    .replace(/<li>/gi, "• ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
}

function narrativeForSection(data: TemplateData, terms: string[]): TemplateRow[] {
  const normalizedTerms = terms.map((term) => term.toLowerCase());
  return byMonth(data.narrative, data.meta.activeMonth).filter((row) =>
    normalizedTerms.some((term) => String(row.Section).toLowerCase().includes(term)),
  );
}

function buildExecutiveOverviewSlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const support = firstByMonth(data.support, data.meta.activeMonth);
  const security = firstByMonth(data.security, data.meta.activeMonth);
  const change = firstByMonth(data.change, data.meta.activeMonth);
  const dev = firstByMonth(data.dev, data.meta.activeMonth);

  if (!support || !security || !change || !dev) {
    return;
  }

  addSectionEyebrow(slide, "Executive scorecard", CONTENT_X, CONTENT_Y);
  slide.addText("Operational pulse", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 4.6,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addKpiRow(
    slide,
    [
      { label: "Support SLA", value: String(support.ResolutionSLA), note: "Resolution performance", accent: COLORS.blue },
      { label: "User CSAT", value: String(support.CSAT), note: "Customer sentiment", accent: COLORS.teal },
      { label: "Critical vulns", value: fmtNumber(parseNumeric(security.CritVulns)), note: "Open critical items", accent: COLORS.teal },
      { label: "Change success", value: String(change.SuccessRate), note: "Successful changes", accent: COLORS.orange },
      { label: "Dev backlog", value: fmtNumber(parseNumeric(dev.BacklogEnd)), note: "Open delivery backlog", accent: COLORS.muted },
    ],
    CONTENT_X,
    1.72,
    CONTENT_W,
    1.5,
  );

  const highlights = narrativeForSection(data, ["executive", "support", "security"]).slice(0, 3);
  addInsightBox(
    slide,
    CONTENT_X,
    3.5,
    5.95,
    2.62,
    "Current read-out",
    highlights.map((item) => `${item.Headline}: ${item.Narrative}`).join("\n\n"),
    COLORS.blue,
  );

  const summaryRows = [
    ["Service desk", `${support.Opened} opened / ${support.Closed} closed`, String(support.Commentary || "Steady month")],
    ["Security", `${security.CritVulns} critical · ${security.OverdueRemediation} overdue`, String(security.Commentary || "Controls improving")],
    ["Change", `${change.TotalChanges} changes · ${change.ChangesIncidents} incident-linked`, String(change.Commentary || "Release cadence healthy")],
    ["Delivery", `${dev.Closed} closed · ${dev.Blocked} blocked`, String(dev.Commentary || "Backlog under control")],
  ];
  addTable(slide, 6.72, 3.5, 6.06, summaryRows, ["Area", "Signal", "Commentary"], 8);
}

function buildExecutiveHighlightsSlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const services = byMonth(data.service, data.meta.activeMonth).slice(0, 6);
  const notes = narrativeForSection(data, ["executive", "support", "security", "assets"]).slice(0, 4);

  addSectionEyebrow(slide, "Operational highlights", CONTENT_X, CONTENT_Y);
  slide.addText("Where momentum is strongest", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 5.6,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  const tileW = (CONTENT_W - 0.28) / 3;
  const tileH = 1.35;
  services.forEach((service, index) => {
    const row = Math.floor(index / 3);
    const col = index % 3;
    const x = CONTENT_X + col * (tileW + 0.14);
    const y = 1.78 + row * (tileH + 0.14);
    addPanel(slide, x, y, tileW, tileH, {
      title: String(service.Service),
      subtitle: String(service.Type),
      accent: parseNumeric(service.Availability) >= parseNumeric(service.Target) ? COLORS.green : COLORS.orange,
    });
    slide.addText(String(service.Availability), {
      x: x + 0.18,
      y: y + 0.52,
      w: tileW - 0.36,
      h: 0.28,
      fontFace: "Arial",
      fontSize: 22,
      bold: true,
      color: COLORS.ink,
      margin: 0,
    });
    slide.addText(`Target ${service.Target} · ${fmtNumber(parseNumeric(service.OutageMins))} outage mins`, {
      x: x + 0.18,
      y: y + 0.92,
      w: tileW - 0.36,
      h: 0.18,
      fontFace: "Arial",
      fontSize: 8,
      color: COLORS.slate,
      margin: 0,
    });
  });

  notes.forEach((note, index) => {
    const x = CONTENT_X + (index % 2) * 6.18;
    const y = 4.78 + Math.floor(index / 2) * 1.0;
    addInsightBox(slide, x, y, 5.98, 0.82, String(note.Headline), String(note.Narrative), index % 2 === 0 ? COLORS.orange : COLORS.blue);
  });
}

function buildServiceOverviewSlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const services = byMonth(data.service, data.meta.activeMonth).slice(0, 6);
  const strongest = [...services].sort((a, b) => parseNumeric(b.Availability) - parseNumeric(a.Availability))[0];
  const weakest = [...services].sort((a, b) => parseNumeric(a.Availability) - parseNumeric(b.Availability))[0];

  addSectionEyebrow(slide, "Service availability", CONTENT_X, CONTENT_Y);
  slide.addText("Core service position", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 4.8,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  const tileW = (CONTENT_W - 0.28) / 3;
  services.forEach((service, index) => {
    const row = Math.floor(index / 3);
    const col = index % 3;
    const x = CONTENT_X + col * (tileW + 0.14);
    const y = 1.78 + row * 1.44;
    addPanel(slide, x, y, tileW, 1.28, {
      title: String(service.Service),
      subtitle: String(service.Type),
      accent: parseNumeric(service.Availability) >= parseNumeric(service.Target) ? COLORS.green : COLORS.orange,
    });
    slide.addText(String(service.Availability), {
      x: x + 0.18,
      y: y + 0.5,
      w: tileW - 0.36,
      h: 0.26,
      fontFace: "Arial",
      fontSize: 21,
      bold: true,
      color: COLORS.ink,
      margin: 0,
    });
    slide.addText(`${fmtNumber(parseNumeric(service.OutageMins))} outage mins · ${service.MajorIncidents} major incidents`, {
      x: x + 0.18,
      y: y + 0.9,
      w: tileW - 0.36,
      h: 0.16,
      fontFace: "Arial",
      fontSize: 7.8,
      color: COLORS.slate,
      margin: 0,
    });
  });

  addInsightBox(
    slide,
    CONTENT_X,
    4.84,
    CONTENT_W,
    0.86,
    "Analyst note",
    `${strongest?.Service ?? "Top service"} is the strongest service in-period at ${strongest?.Availability ?? "—"}, while ${weakest?.Service ?? "the weakest service"} sits at ${weakest?.Availability ?? "—"}.`,
    COLORS.blue,
  );
}

function buildServiceDetailSlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const monthLabels = data.meta.availableMonths.map((month) => fmtMonthShort(month));
  const services = Array.from(new Set(byMonth(data.service, data.meta.activeMonth).map((row) => String(row.Service)))).slice(0, 4);
  const serviceSeries = services.map((serviceName) => ({
    name: serviceName,
    values: data.meta.availableMonths.map((month) => {
      const row = data.service.find((entry) => entry.Month === month && entry.Service === serviceName);
      return parseNumeric(row?.Availability);
    }),
  }));
  const current = byMonth(data.service, data.meta.activeMonth).slice(0, 6);

  addSectionEyebrow(slide, "Trend detail", CONTENT_X, CONTENT_Y);
  slide.addText("Availability and outages", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 5.2,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addPanel(slide, CONTENT_X, 1.78, 7.55, 2.08, { title: "Availability trend", accent: COLORS.blue });
  addLineChart(slide, CONTENT_X + 0.18, 2.16, 7.19, 1.46, monthLabels, serviceSeries, [COLORS.blue, COLORS.orange, COLORS.teal, COLORS.amber], "0.0");

  addPanel(slide, 8.28, 1.78, 4.5, 2.08, { title: "Outage minutes · current month", accent: COLORS.orange });
  addBarChart(
    slide,
    8.44,
    2.16,
    4.18,
    1.46,
    current.map((row) => String(row.Service)),
    [{ name: "Outage mins", values: current.map((row) => parseNumeric(row.OutageMins)) }],
    [COLORS.orange],
  );

  const rows = current.map((row) => [
    String(row.Service),
    String(row.Availability),
    String(row.Target),
    fmtNumber(parseNumeric(row.OutageMins)),
  ]);
  addTable(slide, CONTENT_X, 4.2, CONTENT_W, rows, ["Service", "Availability", "Target", "Outage mins"], 8);
}

function buildNetworkMapSlide(slide: PptxGenJS.Slide, data: TemplateData, imageData: string) {
  const current = firstByMonth(data.derivedNetwork, data.meta.activeMonth);
  const officeCount = byMonth(data.officeNetwork, data.meta.activeMonth).length;
  if (!current) {
    return;
  }

  addSectionEyebrow(slide, "Network & offices", CONTENT_X, CONTENT_Y);
  slide.addText("Estate map view", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 4.6,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addKpiRow(
    slide,
    [
      { label: "Average availability", value: String(current.Availability), note: "Across the estate", accent: COLORS.blue },
      { label: "In-scope offices", value: fmtNumber(officeCount), note: "Live plotted estate", accent: COLORS.teal },
      { label: "Below 99.9%", value: fmtNumber(parseNumeric(current.Below99_9Offices)), note: "Monitoring attention", accent: COLORS.orange },
      { label: "Below 99%", value: fmtNumber(parseNumeric(current.Below99Offices)), note: "Intervention required", accent: COLORS.red },
    ],
    CONTENT_X,
    1.7,
    CONTENT_W,
    1.36,
  );

  addPanel(slide, CONTENT_X, 3.28, CONTENT_W, 3.3, { title: "Office network map", accent: COLORS.blue });
  slide.addImage({
    data: imageData,
    x: CONTENT_X + 0.1,
    y: 3.62,
    w: CONTENT_W - 0.2,
    h: 2.82,
  });
}

function buildNetworkDetailSlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const offices = byMonth(data.officeNetwork, data.meta.activeMonth)
    .sort((a, b) => parseNumeric(a.Availability) - parseNumeric(b.Availability))
    .slice(0, 12);
  const derived = data.meta.availableMonths.map((month) => firstByMonth(data.derivedNetwork, month));
  const note = firstByMonth(data.derivedNetwork, data.meta.activeMonth);

  addSectionEyebrow(slide, "Office detail", CONTENT_X, CONTENT_Y);
  slide.addText("Availability by office", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 5,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addPanel(slide, CONTENT_X, 1.78, 6.85, 4.1, { title: "Lowest-performing offices", accent: COLORS.blue });
  addTable(
    slide,
    CONTENT_X + 0.14,
    2.12,
    6.57,
    offices.map((row) => [String(row.OfficeName), String(row.Region), String(row.Availability), fmtNumber(parseNumeric(row.OutageMins))]),
    ["Office", "Region", "Availability", "Outage mins"],
    7.6,
  );

  addPanel(slide, 7.58, 1.78, 5.2, 2.36, { title: "Estate trend", accent: COLORS.teal });
  addLineChart(
    slide,
    7.74,
    2.12,
    4.88,
    1.66,
    data.meta.availableMonths.map((month) => fmtMonthShort(month)),
    [
      {
        name: "Availability",
        values: derived.map((row) => parseNumeric(row?.Availability)),
      },
    ],
    [COLORS.teal],
    "0.0",
  );

  addInsightBox(
    slide,
    7.58,
    4.34,
    5.2,
    1.54,
    "Current watchpoint",
    note?.WorstOffice
      ? `${note.WorstOffice} is the weakest office this month at ${note.WorstAvailability}. Focus remains on the sites currently below the 99.9% reliability band.`
      : "All tracked offices are performing within the expected range.",
    COLORS.orange,
  );
}

function buildSupportOverviewSlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const support = firstByMonth(data.support, data.meta.activeMonth);
  if (!support) {
    return;
  }

  addSectionEyebrow(slide, "Service desk", CONTENT_X, CONTENT_Y);
  slide.addText("Support operations and user experience", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 6.4,
    h: 0.3,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addHeroMetric(slide, CONTENT_X, 1.72, 6.15, 2.22, {
    eyebrow: "Service level",
    value: String(support.ResolutionSLA),
    label: "Resolution SLA compliance",
    note: `${data.meta.activeMonthLabel} · target 95.0%`,
    accent: COLORS.orange,
  });

  addKpiRow(
    slide,
    [
      { label: "Opened", value: fmtNumber(parseNumeric(support.Opened)), note: "Tickets this month", accent: COLORS.blue },
      { label: "Closed", value: fmtNumber(parseNumeric(support.Closed)), note: "Completed in-month", accent: COLORS.orange },
      { label: "Backlog end", value: fmtNumber(parseNumeric(support.Backlog)), note: "Open at month end", accent: COLORS.teal },
      { label: "Avg resolution", value: `${parseNumeric(support.AvgResolution).toFixed(1)} days`, note: "Average time to resolve", accent: COLORS.teal },
      { label: "Major incidents", value: fmtNumber(parseNumeric(support.MajorIncidents)), note: "Critical incidents", accent: COLORS.muted },
    ],
    CONTENT_X,
    4.3,
    CONTENT_W,
    1.52,
  );
}

function buildSupportVolumesSlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const monthLabels = data.meta.availableMonths.map((month) => fmtMonthShort(month));
  const supportRows = data.meta.availableMonths.map((month) => firstByMonth(data.support, month));
  const current = firstByMonth(data.support, data.meta.activeMonth);
  if (!current) {
    return;
  }

  addSectionEyebrow(slide, "Ticket volumes", CONTENT_X, CONTENT_Y);
  slide.addText("Opened versus closed", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 4.8,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addPanel(slide, CONTENT_X, 1.78, 8.2, 4.08, { title: "Monthly ticket volumes", accent: COLORS.blue });
  addBarChart(
    slide,
    CONTENT_X + 0.14,
    2.08,
    7.92,
    3.44,
    monthLabels,
    [
      { name: "Opened", values: supportRows.map((row) => parseNumeric(row?.Opened)) },
      { name: "Closed", values: supportRows.map((row) => parseNumeric(row?.Closed)) },
    ],
    [COLORS.blue, COLORS.orange],
  );

  const closeBalance = parseNumeric(current.Closed) === 0 ? 0 : (parseNumeric(current.Closed) / Math.max(parseNumeric(current.Opened), 1)) * 100;
  addInsightBox(
    slide,
    9.0,
    1.78,
    3.78,
    1.52,
    "Queue health",
    `${fmtNumber(parseNumeric(current.Closed) - parseNumeric(current.Opened))} net flow this month.\nClose balance is ${closeBalance.toFixed(1)}% with ${current.Backlog} tickets left in the month-end queue.`,
    closeBalance >= 100 ? COLORS.green : closeBalance >= 97 ? COLORS.orange : COLORS.red,
  );

  addInsightBox(
    slide,
    9.0,
    3.48,
    3.78,
    2.38,
    "Pressure note",
    String(current.Commentary || "Support performance held steady through the month."),
    COLORS.orange,
  );
}

function buildSupportDetailSlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const visibleSupport = data.support.filter((row) => data.meta.availableMonths.includes(String(row.Month)));
  const categoryCounts = visibleSupport.reduce<Record<string, number>>((acc, row) => {
    const key = String(row.TopCategory);
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
  const categories = Object.entries(categoryCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([label, value], index) => ({
      label,
      value,
      color: [COLORS.blue, COLORS.orange, COLORS.teal, COLORS.amber, COLORS.red][index % 5],
    }));
  const tickets = byMonth(data.tickets, data.meta.activeMonth)
    .sort((a, b) => parseNumeric(b.AgeDays) - parseNumeric(a.AgeDays))
    .slice(0, 6);

  addSectionEyebrow(slide, "Ticket detail", CONTENT_X, CONTENT_Y);
  slide.addText("Category mix and ageing queue", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 5.8,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addPanel(slide, CONTENT_X, 1.78, 4.25, 4.18, { title: "Tickets by category", accent: COLORS.blue });
  addBarList(slide, categories, CONTENT_X + 0.18, 2.16, 3.88, 3.42);

  addPanel(slide, 5.04, 1.78, 7.74, 4.18, { title: "Top oldest open tickets", accent: COLORS.orange });
  addTable(
    slide,
    5.18,
    2.12,
    7.46,
    tickets.map((ticket) => [
      String(ticket.TicketID),
      String(ticket.Title).slice(0, 28),
      `${ticket.Category} · ${ticket.OwnerQueue}`,
      fmtNumber(parseNumeric(ticket.AgeDays)),
      String(ticket.BusinessCritical),
    ]),
    ["Ticket ID", "Title", "Category · Queue", "Age", "Critical"],
    7.5,
  );
}

function buildSecuritySlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const current = firstByMonth(data.security, data.meta.activeMonth);
  const monthLabels = data.meta.availableMonths.map((month) => fmtMonthShort(month));
  const seriesRows = data.meta.availableMonths.map((month) => firstByMonth(data.security, month));

  if (!current) {
    return;
  }

  addSectionEyebrow(slide, "Security & patching", CONTENT_X, CONTENT_Y);
  slide.addText("Security posture", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 4.2,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addKpiRow(
    slide,
    [
      { label: "Critical vulns", value: fmtNumber(parseNumeric(current.CritVulns)), note: "Open critical items", accent: COLORS.teal },
      { label: "Workstation patch", value: String(current.WkstationPatch), note: "Compliance", accent: COLORS.blue },
      { label: "MFA coverage", value: String(current.MFACoverage), note: "Protected user base", accent: COLORS.teal },
      { label: "Overdue remediation", value: fmtNumber(parseNumeric(current.OverdueRemediation)), note: "Late actions", accent: COLORS.orange },
    ],
    CONTENT_X,
    1.72,
    CONTENT_W,
    1.34,
  );

  addPanel(slide, CONTENT_X, 3.28, 5.0, 2.55, { title: "Patch compliance", accent: COLORS.blue });
  addBarList(
    slide,
    [
      { label: "Workstations", value: parseNumeric(current.WkstationPatch), color: COLORS.blue },
      { label: "Servers", value: parseNumeric(current.ServerPatch), color: COLORS.orange },
      { label: "Critical", value: parseNumeric(current.CriticalPatch), color: COLORS.teal },
    ],
    CONTENT_X + 0.18,
    3.66,
    4.62,
    1.8,
  );

  addPanel(slide, 5.22, 3.28, 4.7, 2.55, { title: "Vulnerability trend", accent: COLORS.orange });
  addLineChart(
    slide,
    5.38,
    3.62,
    4.38,
    1.86,
    monthLabels,
    [
      { name: "Critical", values: seriesRows.map((row) => parseNumeric(row?.CritVulns)) },
      { name: "High", values: seriesRows.map((row) => parseNumeric(row?.HighVulns)) },
    ],
    [COLORS.red, COLORS.orange],
  );

  addInsightBox(slide, 10.08, 3.28, 2.7, 2.55, "Security note", String(current.Commentary || "Security posture improved through the month."), COLORS.teal);
}

function buildAssetsSlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const currentRows = byMonth(data.assets, data.meta.activeMonth);
  const monthLabels = data.meta.availableMonths.map((month) => fmtMonthShort(month));
  const types = Array.from(new Set(currentRows.map((row) => String(row.AssetType)))).slice(0, 3);

  const totalDevices = currentRows.reduce((sum, row) => sum + parseNumeric(row.ActiveDevices), 0);
  const laptops = currentRows.find((row) => String(row.AssetType).toLowerCase().includes("laptop")) ?? currentRows[0];
  const stockCover = currentRows.reduce((sum, row) => sum + parseNumeric(row.StockOnHand), 0);
  const incidents = currentRows.reduce((sum, row) => sum + parseNumeric(row.IncidentsLinked), 0);

  addSectionEyebrow(slide, "Assets & lifecycle", CONTENT_X, CONTENT_Y);
  slide.addText("Estate lifecycle health", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 5.3,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addKpiRow(
    slide,
    [
      { label: "Active devices", value: fmtNumber(totalDevices), note: "Tracked estate", accent: COLORS.blue },
      { label: "In lifecycle", value: String(laptops?.PctWithin ?? "0%"), note: "Laptop estate", accent: COLORS.teal },
      { label: "Aged-kit incidents", value: fmtNumber(incidents), note: "Linked incidents", accent: COLORS.orange },
      { label: "Stock on hand", value: fmtNumber(stockCover), note: "Available stock", accent: COLORS.muted },
    ],
    CONTENT_X,
    1.72,
    CONTENT_W,
    1.34,
  );

  const tileW = 3.84;
  types.forEach((type, index) => {
    const row = currentRows.find((item) => String(item.AssetType) === type);
    if (!row) {
      return;
    }
    const x = CONTENT_X + index * (tileW + 0.13);
    addPanel(slide, x, 3.28, tileW, 1.2, { title: type, accent: [COLORS.blue, COLORS.orange, COLORS.teal][index % 3] });
    slide.addText(String(row.PctWithin), {
      x: x + 0.18,
      y: 3.72,
      w: tileW - 0.36,
      h: 0.22,
      fontFace: "Arial",
      fontSize: 18,
      bold: true,
      color: COLORS.ink,
      margin: 0,
    });
    slide.addText(`${fmtNumber(parseNumeric(row.ActiveDevices))} active · ${fmtNumber(parseNumeric(row.AvgAgeMths))} avg months`, {
      x: x + 0.18,
      y: 4.02,
      w: tileW - 0.36,
      h: 0.16,
      fontFace: "Arial",
      fontSize: 7.8,
      color: COLORS.slate,
      margin: 0,
    });
  });

  addPanel(slide, CONTENT_X, 4.72, 6.0, 1.82, { title: "Lifecycle trend", accent: COLORS.blue });
  addLineChart(
    slide,
    CONTENT_X + 0.14,
    5.02,
    5.72,
    1.2,
    monthLabels,
    types.map((type) => ({
      name: type,
      values: data.meta.availableMonths.map((month) => {
        const row = data.assets.find((entry) => entry.Month === month && entry.AssetType === type);
        return parseNumeric(row?.PctWithin);
      }),
    })),
    [COLORS.blue, COLORS.orange, COLORS.teal],
    "0.0",
  );

  addPanel(slide, 6.7, 4.72, 6.08, 1.82, { title: "Refresh spend · current month", accent: COLORS.orange });
  addBarChart(
    slide,
    6.84,
    5.02,
    5.8,
    1.2,
    currentRows.map((row) => String(row.AssetType)),
    [{ name: "Refresh spend", values: currentRows.map((row) => parseNumeric(row.RefreshSpend)) }],
    [COLORS.orange],
    { yFormat: "#,##0" },
  );
}

function buildChangeSlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const current = firstByMonth(data.change, data.meta.activeMonth);
  const seriesRows = data.meta.availableMonths.map((month) => firstByMonth(data.change, month));

  if (!current) {
    return;
  }

  addSectionEyebrow(slide, "Change & release", CONTENT_X, CONTENT_Y);
  slide.addText("Delivery governance", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 4.8,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addHeroMetric(slide, CONTENT_X, 1.72, 4.08, 2.12, {
    eyebrow: "Change success",
    value: String(current.SuccessRate),
    label: "Successful changes",
    note: `${fmtNumber(parseNumeric(current.ReleasesDeployed))} releases deployed`,
    accent: COLORS.orange,
  });

  addKpiRow(
    slide,
    [
      { label: "Total changes", value: fmtNumber(parseNumeric(current.TotalChanges)), note: "Submitted this month", accent: COLORS.blue },
      { label: "Releases", value: fmtNumber(parseNumeric(current.ReleasesDeployed)), note: "Production releases", accent: COLORS.teal },
      { label: "Failed changes", value: fmtNumber(parseNumeric(current.FailedChanges)), note: "Unsuccessful changes", accent: COLORS.orange },
      { label: "Incident-linked", value: fmtNumber(parseNumeric(current.ChangesIncidents)), note: "Changes causing incidents", accent: COLORS.red },
    ],
    4.32,
    1.72,
    8.46,
    1.34,
  );

  addPanel(slide, CONTENT_X, 4.02, CONTENT_W, 2.52, { title: "Change mix by month", accent: COLORS.blue });
  addBarChart(
    slide,
    CONTENT_X + 0.14,
    4.32,
    CONTENT_W - 0.28,
    1.9,
    data.meta.availableMonths.map((month) => fmtMonthShort(month)),
    [
      { name: "Standard", values: seriesRows.map((row) => parseNumeric(row?.StandardChanges)) },
      { name: "Normal", values: seriesRows.map((row) => parseNumeric(row?.NormalChanges)) },
      { name: "Emergency", values: seriesRows.map((row) => parseNumeric(row?.EmergencyChanges)) },
    ],
    [COLORS.blue, COLORS.orange, COLORS.red],
    { stacked: true },
  );
}

function buildDevSlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const current = firstByMonth(data.dev, data.meta.activeMonth);
  const seriesRows = data.meta.availableMonths.map((month) => firstByMonth(data.dev, month));
  if (!current) {
    return;
  }

  addSectionEyebrow(slide, "Development & delivery", CONTENT_X, CONTENT_Y);
  slide.addText("Backlog and throughput", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 5.2,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addKpiRow(
    slide,
    [
      { label: "Backlog end", value: fmtNumber(parseNumeric(current.BacklogEnd)), note: "Open work at month-end", accent: COLORS.blue },
      { label: "Tasks closed", value: fmtNumber(parseNumeric(current.Closed)), note: "Completed this month", accent: COLORS.teal },
      { label: "Blocked items", value: fmtNumber(parseNumeric(current.Blocked)), note: "Awaiting unblock", accent: COLORS.orange },
      { label: "Dev CSAT", value: String(current.CSAT), note: "Stakeholder score", accent: COLORS.teal },
    ],
    CONTENT_X,
    1.72,
    CONTENT_W,
    1.34,
  );

  addPanel(slide, CONTENT_X, 3.3, 7.6, 2.64, { title: "Backlog pipeline", accent: COLORS.blue });
  addBarChart(
    slide,
    CONTENT_X + 0.14,
    3.62,
    7.32,
    1.98,
    data.meta.availableMonths.map((month) => fmtMonthShort(month)),
    [
      { name: "Opened", values: seriesRows.map((row) => parseNumeric(row?.Opened)) },
      { name: "Closed", values: seriesRows.map((row) => parseNumeric(row?.Closed)) },
    ],
    [COLORS.orange, COLORS.blue],
  );

  addPanel(slide, 8.36, 3.3, 4.42, 2.64, { title: "Delivery mix · current month", accent: COLORS.orange });
  addDoughnutChart(
    slide,
    8.52,
    3.62,
    4.1,
    1.98,
    ["Defects", "Enhancements", "Tech debt", "BAU"],
    [
      parseNumeric(current.Defects),
      parseNumeric(current.Enhancements),
      parseNumeric(current.TechDebt),
      parseNumeric(current.BAU),
    ],
    [COLORS.red, COLORS.blue, COLORS.orange, COLORS.teal],
  );

  addInsightBox(slide, CONTENT_X, 6.18, CONTENT_W, 0.6, "Delivery note", String(current.Commentary || "Delivery flow remained controlled through the month."), COLORS.blue);
}

function buildProjectsSlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const projects = byMonth(data.projects, data.meta.activeMonth);
  const activeProjects = projects.length;
  const avgConfidence =
    activeProjects > 0 ? projects.reduce((sum, row) => sum + parseNumeric(row.Confidence), 0) / activeProjects : 0;
  const decisions = projects.filter((row) => String(row.DecisionNeeded) === "Yes").length;
  const note = narrativeForSection(data, ["project", "delivery", "portfolio"]).slice(0, 1)[0];

  addSectionEyebrow(slide, "Project portfolio", CONTENT_X, CONTENT_Y);
  slide.addText("Active projects", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 4.6,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addKpiRow(
    slide,
    [
      { label: "Active projects", value: fmtNumber(activeProjects), note: "In the current portfolio", accent: COLORS.blue },
      { label: "Avg confidence", value: `${avgConfidence.toFixed(0)}%`, note: "Delivery confidence", accent: COLORS.teal },
      { label: "Decisions needed", value: fmtNumber(decisions), note: "Sponsor or board direction", accent: COLORS.orange },
    ],
    CONTENT_X,
    1.72,
    CONTENT_W,
    1.34,
  );

  addPanel(slide, CONTENT_X, 3.3, 8.3, 2.88, { title: "Project status", accent: COLORS.blue });
  addTable(
    slide,
    CONTENT_X + 0.14,
    3.62,
    8.02,
    projects.slice(0, 6).map((row) => [
      String(row.ProjectName).slice(0, 24),
      String(row.StatusRAG),
      String(row.Confidence),
      String(row.MilestoneNext),
      String(row.BudgetStatus),
    ]),
    ["Project", "RAG", "Confidence", "Next milestone", "Budget"],
    7.4,
  );

  addInsightBox(
    slide,
    8.52,
    3.3,
    4.26,
    2.88,
    "Portfolio note",
    note ? `${note.Headline}\n${note.Narrative}` : `${activeProjects} active projects are being tracked this month, with ${decisions} requiring sponsor or board direction.`,
    COLORS.orange,
  );
}

function buildRoadmapSlide(slide: PptxGenJS.Slide, data: TemplateData, imageData: string) {
  addSectionEyebrow(slide, "Strategy & planning", CONTENT_X, CONTENT_Y);
  slide.addText("Rolling IT roadmap", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 4.8,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  slide.addText("On track", {
    x: CONTENT_X,
    y: 1.76,
    w: 1.1,
    h: 0.12,
    fontFace: "Arial",
    fontSize: 7.8,
    color: COLORS.slate,
    margin: 0,
  });
  const onTrackShape: PptxGenJS.ShapeProps = {
    x: CONTENT_X + 0.95,
    y: 1.81,
    w: 0.24,
    h: 0.05,
    fill: { color: COLORS.green },
    line: { color: COLORS.green, width: 0 },
  };
  slide.addShape("rect", onTrackShape);
  slide.addText("At risk", {
    x: CONTENT_X + 1.5,
    y: 1.76,
    w: 1.1,
    h: 0.12,
    fontFace: "Arial",
    fontSize: 7.8,
    color: COLORS.slate,
    margin: 0,
  });
  const atRiskShape: PptxGenJS.ShapeProps = {
    x: CONTENT_X + 2.2,
    y: 1.81,
    w: 0.24,
    h: 0.05,
    fill: { color: COLORS.amber },
    line: { color: COLORS.amber, width: 0 },
  };
  slide.addShape("rect", atRiskShape);
  slide.addText("Decision required", {
    x: CONTENT_X + 2.7,
    y: 1.76,
    w: 1.5,
    h: 0.12,
    fontFace: "Arial",
    fontSize: 7.8,
    color: COLORS.slate,
    margin: 0,
  });
  const decisionShape: PptxGenJS.ShapeProps = {
    x: CONTENT_X + 3.9,
    y: 1.81,
    w: 0.24,
    h: 0.05,
    fill: { color: COLORS.blue },
    line: { color: COLORS.blue, width: 0 },
  };
  slide.addShape("rect", decisionShape);

  addPanel(slide, CONTENT_X, 2.02, CONTENT_W, 4.5, { title: "Roadmap matrix", accent: COLORS.blue });
  slide.addImage({
    data: imageData,
    x: CONTENT_X + 0.12,
    y: 2.28,
    w: CONTENT_W - 0.24,
    h: 4.04,
  });
}

function buildGanttSlide(slide: PptxGenJS.Slide, data: TemplateData, imageData: string) {
  const month = data.meta.activeMonth;
  const workstreams = byMonth(data.ganttWorkstreams, month).filter((row) => Boolean(row.InScope));
  const milestones = byMonth(data.ganttMilestones, month);
  const onTrack = workstreams.filter((row) => String(row.StatusRAG).toLowerCase() === "green").length;
  const atRisk = workstreams.filter((row) => String(row.StatusRAG).toLowerCase() === "amber").length;

  addSectionEyebrow(slide, "Delivery & projects", CONTENT_X, CONTENT_Y);
  slide.addText("12-week rolling portfolio view", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 6.2,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addPanel(slide, CONTENT_X, 1.78, CONTENT_W, 3.8, { title: "Portfolio gantt", accent: COLORS.blue });
  slide.addImage({
    data: imageData,
    x: CONTENT_X + 0.1,
    y: 2.06,
    w: CONTENT_W - 0.2,
    h: 3.3,
  });

  addKpiRow(
    slide,
    [
      { label: "Active workstreams", value: fmtNumber(workstreams.length), note: "In this 12-week window", accent: COLORS.blue },
      { label: "On track", value: fmtNumber(onTrack), note: "Green RAG", accent: COLORS.teal },
      { label: "At risk", value: fmtNumber(atRisk), note: "Amber RAG", accent: COLORS.orange },
      { label: "Milestones due", value: fmtNumber(milestones.length), note: "Within the rolling window", accent: COLORS.blue },
    ],
    CONTENT_X,
    5.76,
    CONTENT_W,
    1.02,
  );
}

function buildBudgetSlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const currentRows = byMonth(data.budget, data.meta.activeMonth);
  const totals = firstByMonth(data.budgetMonthlyTotals, data.meta.activeMonth);
  const monthSeries = data.meta.availableMonths.map((month) => firstByMonth(data.budgetMonthlyTotals, month));

  addSectionEyebrow(slide, "Budget & commercials", CONTENT_X, CONTENT_Y);
  slide.addText("Financial position", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 4.6,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addKpiRow(
    slide,
    [
      { label: "Total budget", value: fmtCurrency(parseNumeric(totals?.Budget)), note: "Current month budget", accent: COLORS.blue },
      { label: "Total actual", value: fmtCurrency(parseNumeric(totals?.Actual)), note: "Actual spend", accent: COLORS.teal },
      { label: "Variance", value: fmtCurrency(parseNumeric(totals?.Variance)), note: "Budget versus actual", accent: COLORS.orange },
      { label: "Forecast", value: fmtCurrency(parseNumeric(totals?.Forecast)), note: "Expected out-turn", accent: COLORS.blue },
    ],
    CONTENT_X,
    1.72,
    CONTENT_W,
    1.34,
  );

  addPanel(slide, CONTENT_X, 3.3, 6.3, 2.66, { title: "Budget lines", accent: COLORS.blue });
  addTable(
    slide,
    CONTENT_X + 0.14,
    3.62,
    6.02,
    currentRows.slice(0, 5).map((row) => [
      String(row.BudgetLine),
      fmtCurrency(parseNumeric(row.Budget)),
      fmtCurrency(parseNumeric(row.Actual)),
      fmtCurrency(parseNumeric(row.Variance)),
    ]),
    ["Budget line", "Budget", "Actual", "Variance"],
    7.5,
  );

  addPanel(slide, 7.0, 3.3, 5.78, 1.76, { title: "Monthly totals", accent: COLORS.orange });
  addLineChart(
    slide,
    7.14,
    3.62,
    5.5,
    1.18,
    data.meta.availableMonths.map((month) => fmtMonthShort(month)),
    [
      { name: "Budget", values: monthSeries.map((row) => parseNumeric(row?.Budget)) },
      { name: "Actual", values: monthSeries.map((row) => parseNumeric(row?.Actual)) },
      { name: "Forecast", values: monthSeries.map((row) => parseNumeric(row?.Forecast)) },
    ],
    [COLORS.blue, COLORS.orange, COLORS.teal],
    "#,##0",
  );

  addPanel(slide, 7.0, 5.24, 5.78, 1.42, { title: "Upcoming renewals", accent: COLORS.blue });
  addTable(
    slide,
    7.14,
    5.52,
    5.5,
    currentRows
      .filter((row) => String(row.RenewalDue))
      .slice(0, 3)
      .map((row) => [String(row.Vendor), fmtDate(String(row.RenewalDue)), fmtCurrency(parseNumeric(row.RenewalValue))]),
    ["Vendor", "Renewal due", "Value"],
    7.4,
  );
}

function buildRisksSlide(slide: PptxGenJS.Slide, data: TemplateData) {
  const risks = byMonth(data.risks, data.meta.activeMonth);
  const total = risks.length;
  const decisions = risks.filter((row) => String(row.DecisionRequired) === "Yes").length;
  const amber = risks.filter((row) => String(row.RAG).toLowerCase() === "amber").length;
  const note = narrativeForSection(data, ["risk", "governance", "decision"]).slice(0, 1)[0];

  addSectionEyebrow(slide, "Governance", CONTENT_X, CONTENT_Y);
  slide.addText("Risks and decisions", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 4.8,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addKpiRow(
    slide,
    [
      { label: "Total risks", value: fmtNumber(total), note: "Open items", accent: COLORS.blue },
      { label: "Decisions needed", value: fmtNumber(decisions), note: "Board or sponsor action", accent: COLORS.orange },
      { label: "Amber risks", value: fmtNumber(amber), note: "Watchlist items", accent: COLORS.red },
    ],
    CONTENT_X,
    1.72,
    CONTENT_W,
    1.34,
  );

  addPanel(slide, CONTENT_X, 3.3, 8.4, 3.22, { title: "Risk register", accent: COLORS.blue });
  addTable(
    slide,
    CONTENT_X + 0.14,
    3.6,
    8.12,
    risks.slice(0, 5).map((row) => [
      String(row.RiskIssue).slice(0, 42),
      String(row.Owner),
      String(row.RAG),
      String(row.TargetDate),
      String(row.DecisionRequired),
    ]),
    ["Risk / issue", "Owner", "RAG", "Target date", "Decision"],
    7.3,
  );

  addInsightBox(
    slide,
    9.08,
    3.3,
    3.7,
    3.22,
    "Governance note",
    note ? `${note.Headline}\n${note.Narrative}` : `${decisions} item(s) currently require board or executive direction.`,
    COLORS.orange,
  );
}

function buildExecSummarySlide(slide: PptxGenJS.Slide, data: TemplateData) {
  addSectionEyebrow(slide, "Executive summary", CONTENT_X, CONTENT_Y);
  slide.addText("Leadership narrative", {
    x: CONTENT_X,
    y: CONTENT_Y + 0.2,
    w: 5.2,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 24,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });

  addPanel(slide, CONTENT_X, 1.78, CONTENT_W, 4.9, { title: "Exec Summary", accent: COLORS.blue });
  const lines = execSummaryLines(data);
  const summaryTextOptions: PptxGenJS.TextPropsOptions = {
    x: CONTENT_X + 0.26,
    y: 2.14,
    w: CONTENT_W - 0.52,
    h: 4.26,
    fontFace: "Arial",
    fontSize: 11.2,
    color: COLORS.ink,
    margin: 0,
    valign: "top",
  };
  slide.addText(lines.join("\n\n"), summaryTextOptions);

  if (data.execSummary.updatedAt) {
    slide.addText(`Last updated · ${fmtDate(data.execSummary.updatedAt.slice(0, 10))}`, {
      x: CONTENT_X + 0.26,
      y: 6.42,
      w: 3.4,
      h: 0.14,
      fontFace: "Arial",
      fontSize: 8,
      color: COLORS.slate,
      margin: 0,
    });
  }
}

async function buildSlideContent(slide: PptxGenJS.Slide, slideDef: ReportSlideDefinition, data: TemplateData, page: Page) {
  switch (slideDef.id) {
    case "p-summary":
      buildExecSummarySlide(slide, data);
      break;
    case "p-exec-overview":
      buildExecutiveOverviewSlide(slide, data);
      break;
    case "p-exec-highlights":
      buildExecutiveHighlightsSlide(slide, data);
      break;
    case "p-avail-overview":
      buildServiceOverviewSlide(slide, data);
      break;
    case "p-avail-detail":
      buildServiceDetailSlide(slide, data);
      break;
    case "p-network-map":
      buildNetworkMapSlide(slide, data, await captureElementImage(page, "#network-map-block"));
      break;
    case "p-network-detail":
      buildNetworkDetailSlide(slide, data);
      break;
    case "p-support-overview":
      buildSupportOverviewSlide(slide, data);
      break;
    case "p-support-volumes":
      buildSupportVolumesSlide(slide, data);
      break;
    case "p-support-detail":
      buildSupportDetailSlide(slide, data);
      break;
    case "p-security":
      buildSecuritySlide(slide, data);
      break;
    case "p-assets":
      buildAssetsSlide(slide, data);
      break;
    case "p-change":
      buildChangeSlide(slide, data);
      break;
    case "p-dev":
      buildDevSlide(slide, data);
      break;
    case "p-projects":
      buildProjectsSlide(slide, data);
      break;
    case "p-roadmap":
      buildRoadmapSlide(slide, data, await captureElementImage(page, "#rdm-grid"));
      break;
    case "p-gantt":
      buildGanttSlide(slide, data, await captureElementImage(page, "#gantt-chart-block"));
      break;
    case "p-budget":
      buildBudgetSlide(slide, data);
      break;
    case "p-risks":
      buildRisksSlide(slide, data);
      break;
    default:
      addInsightBox(slide, CONTENT_X, 1.72, CONTENT_W, 1.5, "Editable export", `No editable renderer is registered for ${slideDef.slideLabel}.`, COLORS.orange);
  }
}

export async function renderEditablePptx(input: EditablePptxInput): Promise<Buffer> {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "OpenAI Codex";
  pptx.company = "TeacherActive";
  pptx.subject = "Editable IT Reporting slide deck";
  pptx.title = `${input.reportTitle} editable deck`;

  const data = buildTemplateData(input.snapshot, input.month, input.execSummary);
  const slides = getReportSlides();

  for (const [index, slideDef] of slides.entries()) {
    const slide = pptx.addSlide();
    slide.background = { color: COLORS.white };
    addHeader(slide, slideDef, data);
    await buildSlideContent(slide, slideDef, data, input.page);
    addFooter(slide, slideDef, index, slides.length);
  }

  const rawBuffer = await pptx.write({ outputType: "nodebuffer", compression: true });
  return toNodeBuffer(rawBuffer);
}
