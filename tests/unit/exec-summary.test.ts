import { describe, expect, it } from "vitest";

import { buildExecSummaryExcerpt, deriveReportSeriesKey, sanitizeExecSummaryHtml } from "@/lib/reports/exec-summary";

describe("exec summary helpers", () => {
  it("sanitizes rich text to the allowed subset", () => {
    const sanitized = sanitizeExecSummaryHtml(`
      <h1>Title</h1>
      <div><strong>Alpha</strong> <script>alert(1)</script><a href="javascript:alert(1)">bad</a></div>
      <div><a href="https://teacheractive.com">safe</a></div>
    `);

    expect(sanitized).toContain("<h2>Title</h2>");
    expect(sanitized).toContain("<p><strong>Alpha</strong> <a>bad</a></p>");
    expect(sanitized).toContain('<a href="https://teacheractive.com" target="_blank" rel="noopener noreferrer">safe</a>');
    expect(sanitized).not.toContain("<script");
    expect(sanitized).not.toContain("javascript:");
  });

  it("builds a trimmed excerpt from html content", () => {
    const excerpt = buildExecSummaryExcerpt("<h2>Headline</h2><p>One two three four.</p>");
    expect(excerpt).toBe("Headline One two three four.");
  });

  it("derives a stable report series key from a filename", () => {
    expect(deriveReportSeriesKey("IT Exec Reporting.Ingestion Template.xlsx")).toBe("it-exec-reporting-ingestion-template");
  });
});
