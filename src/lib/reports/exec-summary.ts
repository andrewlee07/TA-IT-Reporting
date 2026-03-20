import { format, parseISO } from "date-fns";

export type ExecSummaryMode = "explicit" | "carried-forward" | "empty" | "demo-readonly" | "loading";

export interface ExecSummaryState {
  mode: ExecSummaryMode;
  contentHtml: string;
  excerpt: string;
  updatedAt: string | null;
  sourceReportId: string | null;
}

const ALLOWED_TAGS = new Set(["p", "br", "strong", "b", "em", "ul", "ol", "li", "a", "h2", "h3"]);
const SAFE_HREF_PROTOCOLS = ["http://", "https://", "mailto:"];

function formatMonthLabel(month: string): string {
  return format(parseISO(`${month}-01`), "MMMM yyyy");
}

function escapeAttribute(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function normalizeHref(value: string): string | null {
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }

  const normalized = trimmed.toLowerCase();
  if (!SAFE_HREF_PROTOCOLS.some((protocol) => normalized.startsWith(protocol))) {
    return null;
  }

  return trimmed;
}

function normalizeTagName(tagName: string): string | null {
  const normalized = tagName.toLowerCase();

  if (normalized === "div") {
    return "p";
  }
  if (normalized === "h1") {
    return "h2";
  }
  if (normalized === "h4" || normalized === "h5" || normalized === "h6") {
    return "h3";
  }

  return ALLOWED_TAGS.has(normalized) ? normalized : null;
}

export function stripHtmlToText(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, " ")
    .replace(/<\/(p|li|h2|h3|ul|ol)>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function buildExecSummaryExcerpt(html: string): string {
  const text = stripHtmlToText(html);
  return text.length > 240 ? `${text.slice(0, 237).trimEnd()}...` : text;
}

export function sanitizeExecSummaryHtml(rawHtml: string): string {
  const normalizedInput = rawHtml
    .replace(/\r\n/g, "\n")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/<(\/?)div\b[^>]*>/gi, (_match, closingSlash) => (closingSlash ? "</p>" : "<p>"))
    .replace(/<(\/?)h1\b[^>]*>/gi, (_match, closingSlash) => (closingSlash ? "</h2>" : "<h2>"))
    .replace(/<(\/?)(h4|h5|h6)\b[^>]*>/gi, (_match, closingSlash) => (closingSlash ? "</h3>" : "<h3>"));

  const sanitized = normalizedInput.replace(/<(\/?)([a-z0-9-]+)([^>]*)>/gi, (_match, closingSlash, rawTagName, rawAttributes) => {
    const tagName = normalizeTagName(rawTagName);
    if (!tagName) {
      return "";
    }

    if (closingSlash) {
      return `</${tagName}>`;
    }

    if (tagName === "br") {
      return "<br>";
    }

    if (tagName === "a") {
      const hrefMatch = rawAttributes.match(/\shref\s*=\s*("([^"]*)"|'([^']*)'|([^\s>]+))/i);
      const href = hrefMatch?.[2] ?? hrefMatch?.[3] ?? hrefMatch?.[4] ?? "";
      const safeHref = normalizeHref(href);

      if (!safeHref) {
        return "<a>";
      }

      return `<a href="${escapeAttribute(safeHref)}" target="_blank" rel="noopener noreferrer">`;
    }

    return `<${tagName}>`;
  });

  return sanitized
    .replace(/<(p|h2|h3|ul|ol|li)>\s*<\/\1>/gi, "")
    .replace(/(<br>\s*){3,}/gi, "<br><br>")
    .replace(/\s+<\/(p|li|h2|h3)>/gi, "</$1>")
    .trim();
}

export function deriveReportSeriesKey(filename: string): string {
  return filename
    .replace(/\.[^.]+$/, "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .replace(/-{2,}/g, "-");
}

export function createDemoExecSummary(month: string): ExecSummaryState {
  const monthLabel = formatMonthLabel(month);
  const contentHtml = `
    <h2>${monthLabel} executive narrative</h2>
    <p><strong>Technology operations remained stable through ${monthLabel}</strong>, with the strongest signal coming from service continuity, improved delivery discipline, and a cleaner risk posture across active workstreams.</p>
    <p>The current leadership focus is to <strong>turn strong operational performance into visible business confidence</strong>: keep the estate reliable, keep project sequencing tight, and remove avoidable friction before the next planning cycle.</p>
    <ul>
      <li><strong>Operational picture:</strong> core services and user support are holding steady, which creates room to focus on portfolio decisions rather than incident recovery.</li>
      <li><strong>Delivery picture:</strong> active change and development demand is moving in the right direction, but sequencing and sponsor decisions still determine pace.</li>
      <li><strong>Leadership ask:</strong> use this reporting pack to confirm priorities, unblock cross-team dependencies, and keep the next quarter’s work aligned to business value.</li>
    </ul>
    <p><strong>Overall:</strong> IT is in a good position, but continued clarity on prioritisation and decision ownership is what will keep this momentum visible at executive level.</p>
  `.trim();

  return {
    mode: "demo-readonly",
    contentHtml,
    excerpt: buildExecSummaryExcerpt(contentHtml),
    updatedAt: "2026-06-20T09:00:00.000Z",
    sourceReportId: null,
  };
}

