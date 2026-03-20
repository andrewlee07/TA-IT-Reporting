"use client";

import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type ChangeEvent } from "react";
import { createPortal } from "react-dom";
import Chart from "chart.js/auto";

import { REPORT_PAGES, isValidPageId } from "@/lib/report/blocks";
import { buildTemplateData, formatMonthLabel } from "@/lib/report/template-data";
import { initReportApp } from "@/lib/report/runtime";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

interface ReportListEntry {
  id: string;
  title: string;
  originalFilename: string;
  templateKey: string;
  templateVersion: number;
  currentMonth: string;
  availableMonths: string[];
  createdAt: string;
  updatedAt: string;
}

export interface AppReportRecord extends ReportListEntry {
  snapshot: NormalizedReportSnapshot;
}

interface ReportAppShellProps {
  initialReport: AppReportRecord;
  initialReports: ReportListEntry[];
  initialMonth: string;
  initialPageId: string;
  templateBody: string;
}

interface PortalTargets {
  toggle: Element | null;
  period: Element | null;
  utilities: Element | null;
  reports: Element | null;
}

type ClientExportFormat = "png" | "jpeg";

interface ClientExportTarget {
  id: string;
  label: string;
  element: HTMLElement;
}

interface ReportApiPayload {
  report?: AppReportRecord;
  error?: string;
  issues?: string[];
}

function buildCanonicalUrl(reportId: string, month: string, pageId: string): string {
  const params = new URLSearchParams();
  params.set("report", reportId);
  params.set("month", month);
  params.set("page", pageId);
  return `/?${params.toString()}`;
}

function sanitizeFilename(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .replace(/-{2,}/g, "-");
}

function buildClientExportFilename(reportTitle: string, month: string, label: string, format: ClientExportFormat): string {
  return `${sanitizeFilename(reportTitle)}-${month}-${sanitizeFilename(label)}.${format === "jpeg" ? "jpg" : "png"}`;
}

function toReportListEntry(report: AppReportRecord): ReportListEntry {
  return {
    id: report.id,
    title: report.title,
    originalFilename: report.originalFilename,
    templateKey: report.templateKey,
    templateVersion: report.templateVersion,
    currentMonth: report.currentMonth,
    availableMonths: report.availableMonths,
    createdAt: report.createdAt,
    updatedAt: report.updatedAt,
  };
}

function ensureMonth(report: Pick<AppReportRecord, "availableMonths" | "currentMonth">, month: string | null | undefined): string {
  if (month && report.availableMonths.includes(month)) {
    return month;
  }

  return report.currentMonth;
}

function ensurePage(pageId: string | null | undefined): string {
  return pageId && isValidPageId(pageId) ? pageId : REPORT_PAGES[0].id;
}

async function fetchJson<T>(input: RequestInfo | URL, init?: RequestInit): Promise<T> {
  const response = await fetch(input, init);
  const payload = (await response.json()) as T & { error?: string };

  if (!response.ok) {
    throw new Error(payload.error ?? "Request failed.");
  }

  return payload;
}

export function ReportAppShell({
  initialReport,
  initialReports,
  initialMonth,
  initialPageId,
  templateBody,
}: ReportAppShellProps) {
  const mountRef = useRef<HTMLDivElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const controllerRef = useRef<ReturnType<typeof initReportApp> | null>(null);
  const reportCacheRef = useRef(new Map<string, AppReportRecord>([[initialReport.id, initialReport]]));
  const activeReportRef = useRef(initialReport);
  const selectedMonthRef = useRef(initialMonth);
  const selectedPageRef = useRef(initialPageId);
  const exportTargetsRef = useRef(new Map<string, ClientExportTarget>());

  const [reports, setReports] = useState<ReportListEntry[]>(initialReports);
  const [activeReport, setActiveReport] = useState<AppReportRecord>(initialReport);
  const [selectedMonth, setSelectedMonth] = useState(initialMonth);
  const [selectedPageId, setSelectedPageId] = useState(initialPageId);
  const [targets, setTargets] = useState<PortalTargets>({ toggle: null, period: null, utilities: null, reports: null });
  const [statusMessage, setStatusMessage] = useState<string | null>(null);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [uploadIssues, setUploadIssues] = useState<string[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [isSwitchingReport, setIsSwitchingReport] = useState(false);
  const [busyExport, setBusyExport] = useState<string | null>(null);
  const [busyClientExport, setBusyClientExport] = useState<string | null>(null);
  const [exportError, setExportError] = useState<string | null>(null);
  const [clientExportFormat, setClientExportFormat] = useState<ClientExportFormat>("png");
  const [exportMode, setExportMode] = useState(false);
  const [selectedExportIds, setSelectedExportIds] = useState<string[]>([]);
  const [activeExportTargets, setActiveExportTargets] = useState<ClientExportTarget[]>([]);
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);

  const templateData = useMemo(() => buildTemplateData(activeReport.snapshot, selectedMonth), [activeReport.snapshot, selectedMonth]);
  const reportOptions = useMemo(() => {
    const saved = reports.map((report) => ({
      id: report.id,
      label: report.title,
    }));

    return [
      {
        id: "demo",
        label: "Bundled Demo Report",
      },
      ...saved.filter((report) => report.id !== "demo"),
    ];
  }, [reports]);

  useEffect(() => {
    activeReportRef.current = activeReport;
  }, [activeReport]);

  useEffect(() => {
    selectedMonthRef.current = selectedMonth;
  }, [selectedMonth]);

  useEffect(() => {
    selectedPageRef.current = selectedPageId;
  }, [selectedPageId]);

  useEffect(() => {
    try {
      const storedValue = window.localStorage.getItem("ta-it-reporting-sidebar-collapsed");
      if (storedValue === "true") {
        setIsSidebarCollapsed(true);
      }
    } catch {
      // localStorage access is optional
    }
  }, []);

  useEffect(() => {
    try {
      window.localStorage.setItem("ta-it-reporting-sidebar-collapsed", String(isSidebarCollapsed));
    } catch {
      // localStorage access is optional
    }
  }, [isSidebarCollapsed]);

  const syncUrl = useCallback((reportId: string, month: string, pageId: string, historyMode: "push" | "replace" = "push") => {
    const url = buildCanonicalUrl(reportId, month, pageId);
    const method = historyMode === "replace" ? "replaceState" : "pushState";
    window.history[method]({}, "", url);
  }, []);

  const refreshReportList = useCallback(async (newReport?: AppReportRecord) => {
    try {
      const payload = await fetchJson<{ reports: ReportListEntry[] }>("/api/reports");
      const normalizedReports = payload.reports.map((report) => ({
        ...report,
        createdAt: String(report.createdAt),
        updatedAt: String(report.updatedAt),
      }));

      if (newReport) {
        setReports([toReportListEntry(newReport), ...normalizedReports.filter((report) => report.id !== newReport.id)]);
        return;
      }

      setReports(normalizedReports);
    } catch {
      if (newReport) {
        setReports((current) => [toReportListEntry(newReport), ...current.filter((report) => report.id !== newReport.id)]);
      }
    }
  }, []);

  const loadReport = useCallback(async (reportId: string) => {
    const cached = reportCacheRef.current.get(reportId);
    if (cached) {
      return cached;
    }

    const payload = await fetchJson<ReportApiPayload>(`/api/reports/${reportId}`);
    if (!payload.report) {
      throw new Error("Report not found.");
    }

    reportCacheRef.current.set(payload.report.id, payload.report);
    return payload.report;
  }, []);

  const activateReport = useCallback(
    async (
      reportId: string,
      options: {
        month?: string | null;
        pageId?: string | null;
        historyMode?: "push" | "replace" | "none";
      } = {},
    ) => {
      const currentReport = activeReportRef.current;

      if (reportId === currentReport.id && !options.month && !options.pageId) {
        return;
      }

      setIsSwitchingReport(true);
      setStatusMessage(null);
      setUploadError(null);
      setUploadIssues([]);

      try {
        const report = await loadReport(reportId);
        const nextMonth = ensureMonth(report, options.month);
        const nextPageId = ensurePage(options.pageId);

        setActiveReport(report);
        setSelectedMonth(nextMonth);
        setSelectedPageId(nextPageId);
        setStatusMessage(`Viewing ${report.title}`);

        if (options.historyMode !== "none") {
          syncUrl(report.id, nextMonth, nextPageId, options.historyMode ?? "push");
        }
      } catch (error) {
        setUploadError(error instanceof Error ? error.message : "Failed to load report.");
      } finally {
        setIsSwitchingReport(false);
      }
    },
    [loadReport, syncUrl],
  );

  const handlePopState = useCallback(async () => {
    const params = new URLSearchParams(window.location.search);
    const currentReport = activeReportRef.current;
    const nextReportId = params.get("report") ?? currentReport.id;
    const nextPageId = ensurePage(params.get("page"));

    if (nextReportId !== currentReport.id) {
      await activateReport(nextReportId, {
        month: params.get("month"),
        pageId: nextPageId,
        historyMode: "none",
      });
      return;
    }

    setSelectedMonth(ensureMonth(currentReport, params.get("month")));
    setSelectedPageId(nextPageId);
  }, [activateReport]);

  useEffect(() => {
    const listener = () => {
      void handlePopState();
    };

    window.addEventListener("popstate", listener);
    return () => window.removeEventListener("popstate", listener);
  }, [handlePopState]);

  useEffect(() => {
    exportTargetsRef.current.clear();
    setSelectedExportIds([]);
    setExportError(null);
    setActiveExportTargets([]);
  }, [activeReport.id, selectedMonth, selectedPageId]);

  const handlePageChange = useCallback(
    (pageId: string) => {
      if (selectedPageRef.current === pageId) {
        return;
      }

      selectedPageRef.current = pageId;
      setSelectedPageId(pageId);
      syncUrl(activeReportRef.current.id, selectedMonthRef.current, pageId);
    },
    [syncUrl],
  );

  useLayoutEffect(() => {
    const mountNode = mountRef.current;
    if (!mountNode) {
      return;
    }

    mountNode.innerHTML = templateBody;

    const shellRoot = mountNode.querySelector(".shell");
    if (!shellRoot) {
      return;
    }

    shellRoot.classList.add("app-embedded");

    setTargets({
      toggle: mountNode.querySelector("#sidebar-toggle-slot"),
      period: mountNode.querySelector("#sidebar-period-slot"),
      utilities: mountNode.querySelector("#sidebar-app-utilities-slot"),
      reports: mountNode.querySelector("#sidebar-report-list-slot"),
    });

    controllerRef.current = initReportApp(shellRoot, {
      ChartLib: Chart,
      data: templateData,
      activeMonth: selectedMonth,
      initialPageId: selectedPageRef.current,
      showAllPages: false,
      attachGlobals: true,
      onPageChange: handlePageChange,
    });

    return () => {
      controllerRef.current?.destroy();
      controllerRef.current = null;
      mountNode.innerHTML = "";
      setTargets({ toggle: null, period: null, utilities: null, reports: null });
    };
  }, [handlePageChange, selectedMonth, templateBody, templateData]);

  useEffect(() => {
    const shellRoot = mountRef.current?.querySelector(".shell");
    if (!(shellRoot instanceof HTMLElement)) {
      return;
    }

    shellRoot.classList.toggle("sidebar-collapsed", isSidebarCollapsed);
  }, [isSidebarCollapsed]);

  useEffect(() => {
    controllerRef.current?.showPage(selectedPageId);
  }, [selectedPageId]);

  useEffect(() => {
    const shellRoot = mountRef.current?.querySelector(".shell");
    if (!(shellRoot instanceof HTMLElement)) {
      return;
    }

    shellRoot.classList.toggle("export-mode", exportMode);
  }, [activeReport.id, exportMode, selectedMonth, selectedPageId]);

  const downloadBlob = useCallback((blob: Blob, filename: string) => {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    anchor.click();
    URL.revokeObjectURL(url);
  }, []);

  const renderElementToBlob = useCallback(async (element: HTMLElement, format: ClientExportFormat) => {
    const { default: html2canvas } = await import("html2canvas");
    const canvas = await html2canvas(element, {
      backgroundColor: "#ffffff",
      scale: 2,
      useCORS: true,
      allowTaint: false,
      logging: false,
      imageTimeout: 0,
      onclone: (_clonedDocument, clonedElement) => {
        clonedElement.querySelectorAll?.(".export-icon").forEach((icon) => icon.remove());
      },
    });

    const mimeType = format === "jpeg" ? "image/jpeg" : "image/png";
    const quality = format === "jpeg" ? 0.95 : 1;

    return new Promise<Blob>((resolve, reject) => {
      canvas.toBlob((blob) => {
        if (!blob) {
          reject(new Error("Unable to generate export image."));
          return;
        }

        resolve(blob);
      }, mimeType, quality);
    });
  }, []);

  const withExportChromeHidden = useCallback(async <T,>(work: () => Promise<T>): Promise<T> => {
    const shellRoot = mountRef.current?.querySelector(".shell");
    const hadExportMode = shellRoot instanceof HTMLElement ? shellRoot.classList.contains("export-mode") : false;

    if (shellRoot instanceof HTMLElement) {
      shellRoot.classList.remove("export-mode");
    }

    try {
      return await work();
    } finally {
      if (shellRoot instanceof HTMLElement && hadExportMode) {
        shellRoot.classList.add("export-mode");
      }
    }
  }, []);

  const exportSingleTarget = useCallback(
    async (targetId: string) => {
      if (busyClientExport !== null || busyExport !== null) {
        return;
      }

      const target = exportTargetsRef.current.get(targetId);
      if (!target) {
        setExportError("That report section is not available to export.");
        return;
      }

      setBusyClientExport(`single:${targetId}`);
      setExportError(null);

      try {
        const blob = await withExportChromeHidden(() => renderElementToBlob(target.element, clientExportFormat));
        downloadBlob(
          blob,
          buildClientExportFilename(activeReportRef.current.title, selectedMonthRef.current, target.label, clientExportFormat),
        );
        setStatusMessage(`Exported ${target.label} as ${clientExportFormat.toUpperCase()}.`);
      } catch (error) {
        setExportError(error instanceof Error ? error.message : "Section export failed.");
      } finally {
        setBusyClientExport(null);
      }
    },
    [busyClientExport, busyExport, clientExportFormat, downloadBlob, renderElementToBlob, withExportChromeHidden],
  );

  const exportSelectedTargets = useCallback(async () => {
    if (busyClientExport !== null || busyExport !== null) {
      return;
    }

    const targetsToExport = selectedExportIds
      .map((id) => exportTargetsRef.current.get(id))
      .filter((target): target is ClientExportTarget => Boolean(target))
      .sort((left, right) => {
        const position = left.element.compareDocumentPosition(right.element);
        return position & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1;
      });

    if (targetsToExport.length === 0) {
      setExportError("Select at least one report section to export.");
      return;
    }

    setBusyClientExport("selected");
    setExportError(null);

    try {
      const wrapper = document.createElement("div");
      wrapper.style.cssText = [
        "position:fixed",
        "left:-9999px",
        "top:0",
        "width:1200px",
        "padding:32px",
        "background:#ffffff",
        "display:flex",
        "flex-direction:column",
        "gap:24px",
        "font-family:Arial,sans-serif",
      ].join(";");

      const header = document.createElement("div");
      header.style.cssText = [
        "display:flex",
        "align-items:center",
        "justify-content:space-between",
        "padding-bottom:16px",
        "border-bottom:3px solid #005292",
        "margin-bottom:8px",
      ].join(";");
      header.innerHTML = `
        <div style="display:flex;align-items:center;gap:12px;">
          <div style="width:28px;height:28px;border-radius:4px;background:#F57D00;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:11px;">TA</div>
          <div>
            <div style="font-size:15px;font-weight:700;color:#005292;">TeacherActive · Information Technology</div>
            <div style="font-size:11px;color:#9CA3AF;margin-top:2px;">${activeReportRef.current.title} · ${formatMonthLabel(selectedMonthRef.current)}</div>
          </div>
        </div>
        <div style="font-size:10px;color:#9CA3AF;">INTERNAL · CONFIDENTIAL</div>
      `;
      wrapper.appendChild(header);

      targetsToExport.forEach((target) => {
        const clone = target.element.cloneNode(true);
        if (!(clone instanceof HTMLElement)) {
          return;
        }

        clone.querySelectorAll(".export-icon").forEach((icon) => icon.remove());
        clone.classList.remove("exportable", "selected");
        clone.style.width = "100%";
        clone.style.position = "relative";
        wrapper.appendChild(clone);
      });

      const footer = document.createElement("div");
      footer.style.cssText =
        "border-top:1px solid #E5E7EB;padding-top:12px;display:flex;justify-content:space-between;font-size:10px;color:#9CA3AF;font-family:Arial,sans-serif;";
      footer.innerHTML = `<span>Source: TABS · Internal systems · ${formatMonthLabel(selectedMonthRef.current)}</span><span>${targetsToExport.length} section${targetsToExport.length === 1 ? "" : "s"} exported</span>`;
      wrapper.appendChild(footer);

      document.body.appendChild(wrapper);

      try {
        const blob = await withExportChromeHidden(() => renderElementToBlob(wrapper, clientExportFormat));
        downloadBlob(
          blob,
          buildClientExportFilename(
            activeReportRef.current.title,
            selectedMonthRef.current,
            `${selectedPageRef.current}-selection`,
            clientExportFormat,
          ),
        );
        setStatusMessage(`Exported ${targetsToExport.length} section${targetsToExport.length === 1 ? "" : "s"} as ${clientExportFormat.toUpperCase()}.`);
      } finally {
        document.body.removeChild(wrapper);
      }
    } catch (error) {
      setExportError(error instanceof Error ? error.message : "Combined export failed.");
    } finally {
      setBusyClientExport(null);
    }
  }, [busyClientExport, busyExport, clientExportFormat, downloadBlob, renderElementToBlob, selectedExportIds, withExportChromeHidden]);

  useEffect(() => {
    const shellRoot = mountRef.current?.querySelector(".shell");
    if (!(shellRoot instanceof HTMLElement)) {
      return;
    }

    const activePage = shellRoot.querySelector(`#${selectedPageId}`);
    if (!(activePage instanceof HTMLElement)) {
      return;
    }

    const cleanupCallbacks: Array<() => void> = [];
    const exportTargetMap = exportTargetsRef.current;
    exportTargetMap.clear();

    const roots = Array.from(activePage.querySelectorAll<HTMLElement>("[data-export-id][data-export-label]"));
    const nextTargets: ClientExportTarget[] = [];

    roots.forEach((exportRoot) => {
      const blockId = exportRoot.dataset.exportId;
      const blockLabel = exportRoot.dataset.exportLabel;

      if (!blockId || !blockLabel) {
        return;
      }

      exportRoot.classList.add("exportable");
      exportRoot.dataset.exportTargetId = blockId;
      exportRoot.classList.toggle("selected", selectedExportIds.includes(blockId));

      const icon = document.createElement("button");
      icon.type = "button";
      icon.className = "export-icon";
      icon.title = `Export ${blockLabel}`;
      icon.setAttribute("data-export-target-id", blockId);
      icon.innerHTML =
        "<svg viewBox='0 0 24 24' fill='none' stroke='white' stroke-width='2.5' stroke-linecap='round' stroke-linejoin='round'><path d='M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4'/><polyline points='7 10 12 15 17 10'/><line x1='12' y1='15' x2='12' y2='3'/></svg>";

      const handleIconClick = (event: MouseEvent) => {
        event.preventDefault();
        event.stopPropagation();
        void exportSingleTarget(blockId);
      };

      const handleTargetClick = (event: MouseEvent) => {
        if (!exportMode) {
          return;
        }

        event.preventDefault();
        event.stopPropagation();
        setSelectedExportIds((current) =>
          current.includes(blockId) ? current.filter((id) => id !== blockId) : [...current, blockId],
        );
      };

      icon.addEventListener("click", handleIconClick);
      exportRoot.addEventListener("click", handleTargetClick);
      exportRoot.appendChild(icon);
      exportTargetMap.set(blockId, {
        id: blockId,
        label: blockLabel,
        element: exportRoot,
      });
      nextTargets.push({
        id: blockId,
        label: blockLabel,
        element: exportRoot,
      });

      cleanupCallbacks.push(() => {
        icon.removeEventListener("click", handleIconClick);
        exportRoot.removeEventListener("click", handleTargetClick);
        icon.remove();
        exportRoot.classList.remove("exportable", "selected");
        delete exportRoot.dataset.exportTargetId;
      });
    });

    setActiveExportTargets(nextTargets);

    return () => {
      exportTargetMap.clear();
      setActiveExportTargets([]);
      cleanupCallbacks.forEach((cleanup) => cleanup());
    };
  }, [activeReport.id, exportMode, exportSingleTarget, selectedExportIds, selectedMonth, selectedPageId]);

  const handleMonthChange = useCallback(
    (month: string) => {
      const nextMonth = ensureMonth(activeReportRef.current, month);
      setSelectedMonth(nextMonth);
      syncUrl(activeReportRef.current.id, nextMonth, selectedPageRef.current);
    },
    [syncUrl],
  );

  const handleReportSelect = useCallback(
    async (reportId: string) => {
      await activateReport(reportId, {
        month: selectedMonthRef.current,
        pageId: selectedPageRef.current,
        historyMode: "push",
      });
    },
    [activateReport],
  );

  const handleUpload = useCallback(
    async (file: File | null) => {
      if (!file) {
        return;
      }

      setIsUploading(true);
      setUploadError(null);
      setUploadIssues([]);
      setStatusMessage(null);

      try {
        const formData = new FormData();
        formData.append("workbook", file);

        const response = await fetch("/api/reports", {
          method: "POST",
          body: formData,
        });

        const payload = (await response.json()) as ReportApiPayload;

        if (!response.ok || !payload.report) {
          setUploadError(payload.error ?? "Upload failed.");
          setUploadIssues(payload.issues ?? []);
          return;
        }

        reportCacheRef.current.set(payload.report.id, payload.report);
        await refreshReportList(payload.report);

        const nextPageId = REPORT_PAGES[0].id;
        setActiveReport(payload.report);
        setSelectedMonth(payload.report.currentMonth);
        setSelectedPageId(nextPageId);
        setStatusMessage(`Uploaded ${payload.report.originalFilename}`);
        syncUrl(payload.report.id, payload.report.currentMonth, nextPageId);
      } catch (error) {
        setUploadError(error instanceof Error ? error.message : "Upload failed.");
      } finally {
        setIsUploading(false);
        if (fileInputRef.current) {
          fileInputRef.current.value = "";
        }
      }
    },
    [refreshReportList, syncUrl],
  );

  const handleFileSelection = useCallback(
    async (event: ChangeEvent<HTMLInputElement>) => {
      await handleUpload(event.target.files?.[0] ?? null);
    },
    [handleUpload],
  );

  const downloadExport = useCallback(async (exportType: "page-png" | "full-pdf") => {
    setBusyExport(exportType);
    setExportError(null);

    try {
      const payload: Record<string, string> = {
        exportType,
        month: selectedMonthRef.current,
      };

      if (exportType !== "full-pdf") {
        payload.pageId = selectedPageRef.current;
      }

      const response = await fetch(`/api/reports/${activeReportRef.current.id}/exports`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (!response.ok) {
        const body = (await response.json()) as { error?: string };
        throw new Error(body.error ?? "Export failed.");
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download =
        response.headers.get("content-disposition")?.match(/filename=\"?([^\";]+)\"?/i)?.[1] ??
        `${activeReportRef.current.title}-${exportType}`;
      anchor.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      setExportError(error instanceof Error ? error.message : "Export failed.");
    } finally {
      setBusyExport(null);
    }
  }, []);

  const toggleExportMode = useCallback(() => {
    setExportError(null);
    if (exportMode) {
      setSelectedExportIds([]);
    }
    setExportMode((current) => !current);
  }, [exportMode]);

  const clearSelectedExports = useCallback(() => {
    setSelectedExportIds([]);
    setExportError(null);
  }, []);

  const periodPortal =
    targets.period &&
    createPortal(
      <div className="sidebar-stack-tight">
        <label className="sidebar-field-label" htmlFor="report-month-select">
          Reporting Period
        </label>
        <select
          className="sidebar-select"
          id="report-month-select"
          onChange={(event) => handleMonthChange(event.target.value)}
          value={selectedMonth}
        >
          {activeReport.availableMonths.map((month) => (
            <option key={month} value={month}>
              {formatMonthLabel(month)}
            </option>
          ))}
        </select>
      </div>,
      targets.period,
    );

  const togglePortal =
    targets.toggle &&
    createPortal(
      <button
        aria-label={isSidebarCollapsed ? "Expand sidebar" : "Collapse sidebar"}
        className="sidebar-toggle-button"
        onClick={() => setIsSidebarCollapsed((current) => !current)}
        title={isSidebarCollapsed ? "Expand sidebar" : "Collapse sidebar"}
        type="button"
      >
        <svg fill="none" viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg">
          <path d="M10.5 3.5 6 8l4.5 4.5" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.8" />
        </svg>
      </button>,
      targets.toggle,
    );

  const utilitiesPortal =
    targets.utilities &&
    createPortal(
      <div className="sidebar-stack">
        <div className="sidebar-stack-tight">
          <label className="sidebar-field-label" htmlFor="active-report-select">
            Active Report
          </label>
          <select
            className="sidebar-select"
            disabled={isSwitchingReport || isUploading}
            id="active-report-select"
            onChange={(event) => void handleReportSelect(event.target.value)}
            value={activeReport.id}
          >
            {reportOptions.map((report) => (
              <option key={report.id} value={report.id}>
                {report.label}
              </option>
            ))}
          </select>
          <div className="sidebar-meta sidebar-meta-ellipsis" title={activeReport.title}>
            {activeReport.title}
          </div>
          <div className="sidebar-meta">Template {activeReport.templateKey} · v{activeReport.templateVersion}</div>
        </div>

        <div className="sidebar-stack-tight">
          <span className="sidebar-field-label">Workbook Upload</span>
          <input
            accept=".xlsx"
            className="sidebar-input-hidden"
            onChange={handleFileSelection}
            ref={fileInputRef}
            type="file"
          />
          <button
            className="sidebar-button primary"
            disabled={isUploading || isSwitchingReport}
            onClick={() => fileInputRef.current?.click()}
            type="button"
          >
            {isUploading ? "Uploading workbook..." : "Upload workbook"}
          </button>
          <a className="sidebar-link" href="/templates/IT_Exec_Reporting_Ingestion_Template_v3_dummy_data.xlsx">
            Download template
          </a>
        </div>

        <div className="sidebar-stack-tight">
          <span className="sidebar-field-label">Exports</span>
          <div className="sidebar-inline">
            <button
              className="sidebar-button secondary"
              disabled={busyExport !== null || busyClientExport !== null}
              onClick={() => void downloadExport("page-png")}
              type="button"
            >
              {busyExport === "page-png" ? "Rendering..." : "Page PNG"}
            </button>
            <button
              className="sidebar-button secondary"
              disabled={busyExport !== null || busyClientExport !== null}
              onClick={() => void downloadExport("full-pdf")}
              type="button"
            >
              {busyExport === "full-pdf" ? "Rendering..." : "Full PDF"}
            </button>
          </div>
          <button
            className={`sidebar-button ${exportMode ? "primary is-active" : "secondary"}`}
            disabled={busyExport !== null || busyClientExport !== null || activeExportTargets.length === 0}
            onClick={toggleExportMode}
            type="button"
          >
            {exportMode ? "Exit Select Mode" : "Select to Export"}
          </button>
          {exportMode ? (
            <>
              <div className="sidebar-inline">
                <button
                  className={`sidebar-button secondary ${clientExportFormat === "png" ? "is-active" : ""}`}
                  disabled={busyExport !== null || busyClientExport !== null}
                  onClick={() => setClientExportFormat("png")}
                  type="button"
                >
                  PNG
                </button>
                <button
                  className={`sidebar-button secondary ${clientExportFormat === "jpeg" ? "is-active" : ""}`}
                  disabled={busyExport !== null || busyClientExport !== null}
                  onClick={() => setClientExportFormat("jpeg")}
                  type="button"
                >
                  JPEG
                </button>
              </div>
              <div className="sidebar-meta">
                {selectedExportIds.length} item{selectedExportIds.length === 1 ? "" : "s"} selected
                <br />
                {activeExportTargets.length} exportable item{activeExportTargets.length === 1 ? "" : "s"} on this page
              </div>
              <div className="sidebar-inline">
                <button
                  className="sidebar-button secondary"
                  disabled={!exportMode || busyExport !== null || busyClientExport !== null || selectedExportIds.length === 0}
                  onClick={() => void exportSelectedTargets()}
                  type="button"
                >
                  {busyClientExport === "selected" ? "Rendering..." : "Export Selected"}
                </button>
                <button
                  className="sidebar-button secondary"
                  disabled={!exportMode || busyExport !== null || busyClientExport !== null || selectedExportIds.length === 0}
                  onClick={clearSelectedExports}
                  type="button"
                >
                  Clear
                </button>
              </div>
            </>
          ) : null}
        </div>

        {statusMessage ? <div className="sidebar-meta">{statusMessage}</div> : null}
        {uploadError ? <div className="sidebar-error">{uploadError}</div> : null}
        {uploadIssues.length > 0 ? (
          <div className="sidebar-error">
            {uploadIssues.map((issue) => (
              <div key={issue}>{issue}</div>
            ))}
          </div>
        ) : null}
        {exportError ? <div className="sidebar-error">{exportError}</div> : null}
      </div>,
      targets.utilities,
    );

  const reportsPortal =
    targets.reports &&
    createPortal(
      <div className="sidebar-report-list">
        <button
          className={`sidebar-report-item ${activeReport.id === "demo" ? "active" : ""}`}
          disabled={isSwitchingReport || isUploading}
          onClick={() => void handleReportSelect("demo")}
          type="button"
        >
          <div className="sidebar-report-title">Bundled Demo Report</div>
          <div className="sidebar-report-sub">Prototype snapshot · 2026-06</div>
        </button>

        {reports.length === 0 ? (
          <div className="sidebar-empty">No saved workbooks yet. Upload a workbook to create the first saved report.</div>
        ) : (
          reports.map((report) => (
            <button
              className={`sidebar-report-item ${activeReport.id === report.id ? "active" : ""}`}
              disabled={isSwitchingReport || isUploading}
              key={report.id}
              onClick={() => void handleReportSelect(report.id)}
              type="button"
            >
              <div className="sidebar-report-title">{report.title}</div>
              <div className="sidebar-report-sub">
                {formatMonthLabel(report.currentMonth)} · v{report.templateVersion}
              </div>
            </button>
          ))
        )}
      </div>,
      targets.reports,
    );

  return (
    <>
      <div ref={mountRef} />
      {togglePortal}
      {periodPortal}
      {utilitiesPortal}
      {reportsPortal}
    </>
  );
}
