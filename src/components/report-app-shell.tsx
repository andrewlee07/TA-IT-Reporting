"use client";

import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type ChangeEvent, type KeyboardEvent } from "react";
import { createPortal } from "react-dom";
import Chart from "chart.js/auto";

import { ReportAnnotationLayer } from "@/components/report-annotation-layer";
import { ReportAnnotationToolbar } from "@/components/report-annotation-toolbar";
import { ReportEmbeddedEditor } from "@/components/report-embedded-editor";
import { ReportPrepDrawer } from "@/components/report-prep-drawer";
import { DEFAULT_ANNOTATION_STYLE, createEmptyReportAnnotationsState, type ReportAnnotation, type ReportAnnotationsState, type ReportAnnotationTool } from "@/lib/annotations/types";
import type { EditableReportDraft, SectionId } from "@/lib/drafts/types";
import { EDITOR_SECTIONS_CONFIG } from "@/lib/editor/config";
import { getSectionPayload } from "@/lib/editor/sections";
import { REPORT_PAGES, getReportSlides, getSlideId, hasPageTabs, isExportablePageId, isValidPageId, resolveTabId } from "@/lib/report/blocks";
import { buildTemplateData, formatMonthLabel } from "@/lib/report/template-data";
import { initReportApp } from "@/lib/report/runtime";
import type { ExecSummaryState } from "@/lib/reports/exec-summary";
import type { ReportPrepView } from "@/lib/reports/prep-center";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

interface ReportListEntry {
  id: string;
  title: string;
  originalFilename: string;
  reportSeriesKey: string;
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
  initialAnnotations: ReportAnnotationsState;
  initialExecSummary: ExecSummaryState;
  initialMonth: string;
  initialPageId: string;
  initialTabId: string | null;
  templateBody: string;
}

interface PortalTargets {
  toggle: Element | null;
  period: Element | null;
  utilities: Element | null;
  reports: Element | null;
  summaryControls: Element | null;
  summaryEditor: Element | null;
  annotationRoots: Record<string, Element>;
  editorRoots: Partial<Record<SectionId, Element>>;
}

type ClientExportFormat = "png" | "jpeg";
type CollapsedNavStyle = "icons" | "initials";

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

interface ExecSummaryApiPayload {
  summary?: ExecSummaryState;
  error?: string;
}

interface PrepApiPayload {
  prep?: ReportPrepView;
  error?: string;
}

interface EditorApiPayload {
  draft?: EditableReportDraft;
  error?: string;
}

interface AnnotationApiPayload {
  annotationState?: ReportAnnotationsState;
  error?: string;
}

type EditorSaveState = "idle" | "dirty" | "saving" | "saved" | "conflict" | "error";
type AnnotationSaveState = "idle" | "dirty" | "saving" | "saved" | "conflict" | "error";

const DATA_ENTRY_PAGE_ID = "p-data";

function buildCanonicalUrl(reportId: string, month: string, pageId: string, tabId?: string | null): string {
  const params = new URLSearchParams();
  params.set("report", reportId);
  params.set("month", month);
  params.set("page", pageId);

  if (hasPageTabs(pageId)) {
    const resolvedTabId = resolveTabId(pageId, tabId);
    if (resolvedTabId) {
      params.set("tab", resolvedTabId);
    }
  }

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

function formatSidebarReportTitle(title: string, currentMonth: string): string {
  const monthSuffixPattern = new RegExp(`\\s*·\\s*${currentMonth.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`);
  return title.replace(monthSuffixPattern, "").trim();
}

function getCurrentMonthValue(): string {
  const today = new Date();
  return `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}`;
}

function buildDefaultBlankReportTitle(month: string): string {
  return `TA IT Report ${month}`;
}

function buildEditorNavMarkup(): string {
  return [
    `<div class="nav-link" data-page-id="${DATA_ENTRY_PAGE_ID}" title="Data Entry" onclick="showPage('${DATA_ENTRY_PAGE_ID}',this)">`,
    `<div class="nav-icon"><span class="nav-icon-label">DE</span><svg class="nav-icon-glyph" viewBox="0 0 16 16" aria-hidden="true"><path d="M3.5 3.5h9v9h-9z"></path><path d="M5.5 6h5"></path><path d="M5.5 8.5h5"></path><path d="M5.5 11h3"></path></svg></div>`,
    `<span class="nav-text">Data Entry</span>`,
    `<span class="nav-tooltip" role="tooltip">Data Entry</span>`,
    `</div>`,
  ].join("");
}

function buildEditorPageMarkup(section: { id: string; label: string; description: string }): string {
  return [
    `<div class="report-page" id="${DATA_ENTRY_PAGE_ID}-${section.id}" data-page-id="${DATA_ENTRY_PAGE_ID}" data-tab-id="${section.id}" data-export="false">`,
    `<div class="ph">`,
    `<div class="ph-brand"><div class="ph-mark">TA</div><div><div class="ph-org">TeacherActive</div><div class="ph-dept">Information Technology</div></div></div>`,
    `<div class="ph-meta">`,
    `<div><div class="ph-title">Data Entry</div><div class="ph-sub">Workbook-backed admin inputs in the same report shell</div></div>`,
    `<div><div class="ph-period-label">Reporting Period</div><div class="ph-period-val">June 2026</div></div>`,
    `</div>`,
    `</div>`,
    `<div class="pb">`,
    `<div class="sl"><div class="sl-tag">Admin Workspace</div><div class="sl-title">${section.label}</div><div class="sl-sub">${section.description}</div></div>`,
    `<div id="data-entry-root-${section.id}"></div>`,
    `</div>`,
    `<div class="pf"><div class="pf-source">Internal admin workspace · Not included in PDF/PPTX exports</div><div class="pf-page">DATA ENTRY · ${section.label}</div></div>`,
    `</div>`,
  ].join("");
}

function ensureEditorPages(shellRoot: HTMLElement): Partial<Record<SectionId, Element>> {
  const sidebar = shellRoot.querySelector(".sidebar");
  const main = shellRoot.querySelector(".main");
  const utilitiesHeading = Array.from(shellRoot.querySelectorAll(".nav-section")).find((section) => section.textContent?.trim() === "App Utilities");

  if (!sidebar || !main) {
    return {};
  }

  if (!shellRoot.querySelector(`.nav-link[data-page-id="${DATA_ENTRY_PAGE_ID}"]`) && utilitiesHeading?.parentNode) {
    utilitiesHeading.parentNode.insertBefore(document.createRange().createContextualFragment(buildEditorNavMarkup()), utilitiesHeading);
  }

  EDITOR_SECTIONS_CONFIG.forEach((section) => {
    if (!shellRoot.querySelector(`#${DATA_ENTRY_PAGE_ID}-${section.id}`)) {
      main.insertAdjacentHTML("beforeend", buildEditorPageMarkup(section));
    }
  });

  return Object.fromEntries(
    EDITOR_SECTIONS_CONFIG.map((section) => [section.id, shellRoot.querySelector(`#data-entry-root-${section.id}`)]),
  ) as Partial<Record<SectionId, Element>>;
}

function toReportListEntry(report: AppReportRecord): ReportListEntry {
  return {
    id: report.id,
    title: report.title,
    originalFilename: report.originalFilename,
    reportSeriesKey: report.reportSeriesKey,
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

function ensureTab(pageId: string, tabId: string | null | undefined): string | null {
  return resolveTabId(pageId, tabId);
}

async function fetchJson<T>(input: RequestInfo | URL, init?: RequestInit): Promise<T> {
  const response = await fetch(input, init);
  const payload = (await response.json()) as T & { error?: string };

  if (!response.ok) {
    throw new Error(payload.error ?? "Request failed.");
  }

  return payload;
}

interface ExecSummaryEditorProps {
  initialHtml: string;
  isSaving: boolean;
  onCancel: () => void;
  onSave: (contentHtml: string) => Promise<void>;
}

interface MonthPickerProps {
  availableMonths: string[];
  selectedMonth: string;
  onChange: (month: string) => void;
}

function createClientAnnotationId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }

  return `annotation-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function stripAnnotationAuthoringChrome(root: ParentNode): void {
  root
    .querySelectorAll(
      ".annotation-stage-capture, .annotation-drag-handle, .annotation-resize-handle, .annotation-tail-handle, .annotation-delete-pill, .annotation-toolbar-shell",
    )
    .forEach((node) => node.remove());
}

const MONTH_PICKER_LABEL_ID = "report-month-picker-label";
const MONTH_PICKER_TRIGGER_ID = "report-month-trigger";
const MONTH_PICKER_LISTBOX_ID = "report-month-listbox";

function MonthPicker({ availableMonths, selectedMonth, onChange }: MonthPickerProps) {
  const rootRef = useRef<HTMLDivElement | null>(null);
  const buttonRef = useRef<HTMLButtonElement | null>(null);
  const listboxRef = useRef<HTMLDivElement | null>(null);
  const selectedIndex = Math.max(availableMonths.indexOf(selectedMonth), 0);
  const [isOpen, setIsOpen] = useState(false);
  const [activeIndex, setActiveIndex] = useState(selectedIndex);

  const closePicker = useCallback((restoreFocus = true) => {
    setIsOpen(false);

    if (restoreFocus) {
      window.requestAnimationFrame(() => {
        buttonRef.current?.focus();
      });
    }
  }, []);

  const openPicker = useCallback(
    (nextIndex = selectedIndex) => {
      const boundedIndex = Math.min(Math.max(nextIndex, 0), Math.max(availableMonths.length - 1, 0));
      setActiveIndex(boundedIndex);
      setIsOpen(true);
    },
    [availableMonths.length, selectedIndex],
  );

  const commitSelection = useCallback(
    (index: number) => {
      const nextMonth = availableMonths[index];
      if (!nextMonth) {
        return;
      }

      onChange(nextMonth);
      setActiveIndex(index);
      closePicker();
    },
    [availableMonths, closePicker, onChange],
  );

  useEffect(() => {
    setActiveIndex(selectedIndex);
  }, [selectedIndex]);

  useEffect(() => {
    if (!isOpen) {
      return;
    }

    window.requestAnimationFrame(() => {
      listboxRef.current?.focus();
    });

    const handlePointerDown = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) {
        closePicker(false);
      }
    };

    document.addEventListener("mousedown", handlePointerDown);
    return () => document.removeEventListener("mousedown", handlePointerDown);
  }, [closePicker, isOpen]);

  const handleTriggerKeyDown = useCallback(
    (event: KeyboardEvent<HTMLButtonElement>) => {
      switch (event.key) {
        case "ArrowDown":
          event.preventDefault();
          openPicker(Math.min(selectedIndex + 1, availableMonths.length - 1));
          return;
        case "ArrowUp":
          event.preventDefault();
          openPicker(Math.max(selectedIndex - 1, 0));
          return;
        case "Enter":
        case " ":
          event.preventDefault();
          if (isOpen) {
            closePicker();
          } else {
            openPicker(selectedIndex);
          }
          return;
        case "Escape":
          if (isOpen) {
            event.preventDefault();
            closePicker();
          }
          return;
        default:
          return;
      }
    },
    [availableMonths.length, closePicker, isOpen, openPicker, selectedIndex],
  );

  const handleListboxKeyDown = useCallback(
    (event: KeyboardEvent<HTMLDivElement>) => {
      switch (event.key) {
        case "ArrowDown":
          event.preventDefault();
          setActiveIndex((current) => Math.min(current + 1, availableMonths.length - 1));
          return;
        case "ArrowUp":
          event.preventDefault();
          setActiveIndex((current) => Math.max(current - 1, 0));
          return;
        case "Home":
          event.preventDefault();
          setActiveIndex(0);
          return;
        case "End":
          event.preventDefault();
          setActiveIndex(Math.max(availableMonths.length - 1, 0));
          return;
        case "Enter":
        case " ":
          event.preventDefault();
          commitSelection(activeIndex);
          return;
        case "Escape":
          event.preventDefault();
          closePicker();
          return;
        case "Tab":
          closePicker(false);
          return;
        default:
          return;
      }
    },
    [activeIndex, availableMonths.length, closePicker, commitSelection],
  );

  return (
    <div className={`sidebar-month-picker${isOpen ? " is-open" : ""}`} ref={rootRef}>
      <label className="sidebar-field-label" id={MONTH_PICKER_LABEL_ID}>
        Reporting Period
      </label>
      <button
        aria-controls={MONTH_PICKER_LISTBOX_ID}
        aria-expanded={isOpen}
        aria-haspopup="listbox"
        aria-labelledby={`${MONTH_PICKER_LABEL_ID} ${MONTH_PICKER_TRIGGER_ID}`}
        className="month-picker-trigger"
        id={MONTH_PICKER_TRIGGER_ID}
        onClick={() => {
          if (isOpen) {
            closePicker(false);
            return;
          }
          openPicker(selectedIndex);
        }}
        onKeyDown={handleTriggerKeyDown}
        ref={buttonRef}
        type="button"
      >
        <span className="month-picker-trigger-value">{formatMonthLabel(selectedMonth)}</span>
        <span aria-hidden="true" className="month-picker-trigger-icon">
          <svg fill="none" viewBox="0 0 12 12" xmlns="http://www.w3.org/2000/svg">
            <path d="M2 4.25 6 8l4-3.75" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.6" />
          </svg>
        </span>
      </button>

      {isOpen ? (
        <div
          aria-activedescendant={`report-month-option-${availableMonths[activeIndex]}`}
          aria-labelledby={MONTH_PICKER_LABEL_ID}
          className="month-picker-panel"
          id={MONTH_PICKER_LISTBOX_ID}
          onKeyDown={handleListboxKeyDown}
          ref={listboxRef}
          role="listbox"
          tabIndex={-1}
        >
          {availableMonths.map((month, index) => {
            const isSelected = month === selectedMonth;
            const isActive = index === activeIndex;

            return (
              <button
                aria-selected={isSelected}
                className={`month-picker-option${isSelected ? " is-selected" : ""}${isActive ? " is-active" : ""}`}
                id={`report-month-option-${month}`}
                key={month}
                onClick={() => commitSelection(index)}
                onMouseEnter={() => setActiveIndex(index)}
                onMouseDown={(event) => event.preventDefault()}
                role="option"
                type="button"
              >
                <span className="month-picker-option-text">{formatMonthLabel(month)}</span>
                {isSelected ? <span className="month-picker-option-badge">Current</span> : null}
              </button>
            );
          })}
        </div>
      ) : null}
    </div>
  );
}

function ExecSummaryEditor({ initialHtml, isSaving, onCancel, onSave }: ExecSummaryEditorProps) {
  const editorRef = useRef<HTMLDivElement | null>(null);
  const [value, setValue] = useState(initialHtml);

  useEffect(() => {
    setValue(initialHtml);
  }, [initialHtml]);

  useEffect(() => {
    if (editorRef.current && editorRef.current.innerHTML !== value) {
      editorRef.current.innerHTML = value;
    }
  }, [value]);

  const runCommand = useCallback((command: string, commandValue?: string) => {
    editorRef.current?.focus();
    document.execCommand(command, false, commandValue);
    setValue(editorRef.current?.innerHTML ?? "");
  }, []);

  const handleLink = useCallback(() => {
    const nextUrl = window.prompt("Enter a full URL", "https://");
    if (!nextUrl) {
      return;
    }
    runCommand("createLink", nextUrl);
  }, [runCommand]);

  return (
    <div className="summary-editor-shell">
      <div className="summary-editor-toolbar">
        <button className="summary-editor-btn" onClick={() => runCommand("formatBlock", "H2")} type="button">
          Heading
        </button>
        <button className="summary-editor-btn" onClick={() => runCommand("formatBlock", "P")} type="button">
          Paragraph
        </button>
        <button className="summary-editor-btn" onClick={() => runCommand("bold")} type="button">
          Bold
        </button>
        <button className="summary-editor-btn" onClick={() => runCommand("insertUnorderedList")} type="button">
          Bullets
        </button>
        <button className="summary-editor-btn" onClick={handleLink} type="button">
          Link
        </button>
      </div>

      <div
        className="summary-editor"
        contentEditable
        onInput={() => setValue(editorRef.current?.innerHTML ?? "")}
        ref={editorRef}
        suppressContentEditableWarning
      />

      <div className="summary-editor-actions">
        <button className="summary-action-btn secondary" disabled={isSaving} onClick={onCancel} type="button">
          Cancel
        </button>
        <button className="summary-action-btn primary" disabled={isSaving} onClick={() => void onSave(value)} type="button">
          {isSaving ? "Saving..." : "Save"}
        </button>
      </div>
    </div>
  );
}

export function ReportAppShell({
  initialReport,
  initialReports,
  initialAnnotations,
  initialExecSummary,
  initialMonth,
  initialPageId,
  initialTabId,
  templateBody,
}: ReportAppShellProps) {
  const mountRef = useRef<HTMLDivElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const controllerRef = useRef<ReturnType<typeof initReportApp> | null>(null);
  const reportCacheRef = useRef(new Map<string, AppReportRecord>([[initialReport.id, initialReport]]));
  const annotationCacheRef = useRef(new Map<string, ReportAnnotationsState>([[`${initialReport.id}:${initialMonth}`, initialAnnotations]]));
  const execSummaryCacheRef = useRef(new Map<string, ExecSummaryState>([[`${initialReport.id}:${initialMonth}`, initialExecSummary]]));
  const prepCacheRef = useRef(new Map<string, ReportPrepView>());
  const activeReportRef = useRef(initialReport);
  const selectedMonthRef = useRef(initialMonth);
  const selectedPageRef = useRef(initialPageId);
  const selectedTabByPageRef = useRef<Record<string, string | null>>(
    initialTabId ? { [initialPageId]: initialTabId } : {},
  );
  const exportTargetsRef = useRef(new Map<string, ClientExportTarget>());

  const [reports, setReports] = useState<ReportListEntry[]>(initialReports);
  const [activeReport, setActiveReport] = useState<AppReportRecord>(initialReport);
  const [selectedMonth, setSelectedMonth] = useState(initialMonth);
  const [selectedPageId, setSelectedPageId] = useState(initialPageId);
  const [selectedTabByPage, setSelectedTabByPage] = useState<Record<string, string | null>>(
    initialTabId ? { [initialPageId]: initialTabId } : {},
  );
  const [targets, setTargets] = useState<PortalTargets>({
    toggle: null,
    period: null,
    utilities: null,
    reports: null,
    summaryControls: null,
    summaryEditor: null,
    annotationRoots: {},
    editorRoots: {},
  });
  const [statusMessage, setStatusMessage] = useState<string | null>(null);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [uploadIssues, setUploadIssues] = useState<string[]>([]);
  const [isCreateBlankOpen, setIsCreateBlankOpen] = useState(false);
  const [isCreatingBlank, setIsCreatingBlank] = useState(false);
  const [blankReportMonth, setBlankReportMonth] = useState(initialMonth || getCurrentMonthValue());
  const [blankReportTitle, setBlankReportTitle] = useState(buildDefaultBlankReportTitle(initialMonth || getCurrentMonthValue()));
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
  const [collapsedNavStyle, setCollapsedNavStyle] = useState<CollapsedNavStyle>("icons");
  const [annotationState, setAnnotationState] = useState<ReportAnnotationsState>(initialAnnotations);
  const [isAnnotationsLoading, setIsAnnotationsLoading] = useState(false);
  const [annotationSaveState, setAnnotationSaveState] = useState<AnnotationSaveState>("idle");
  const [annotationSaveMessage, setAnnotationSaveMessage] = useState("Annotations ready");
  const [annotationTool, setAnnotationTool] = useState<ReportAnnotationTool>("select");
  const [selectedAnnotationId, setSelectedAnnotationId] = useState<string | null>(null);
  const [execSummary, setExecSummary] = useState<ExecSummaryState>(initialExecSummary);
  const [isSummaryLoading, setIsSummaryLoading] = useState(false);
  const [isSummarySaving, setIsSummarySaving] = useState(false);
  const [isSummaryEditing, setIsSummaryEditing] = useState(false);
  const [summaryEditorHtml, setSummaryEditorHtml] = useState(initialExecSummary.contentHtml);
  const [prepView, setPrepView] = useState<ReportPrepView | null>(null);
  const [isPrepLoading, setIsPrepLoading] = useState(false);
  const [isPrepSaving, setIsPrepSaving] = useState(false);
  const [isPrepOpen, setIsPrepOpen] = useState(false);
  const [activePrepTab, setActivePrepTab] = useState<"readiness" | "rollover">("readiness");
  const [editorDraft, setEditorDraft] = useState<EditableReportDraft | null>(null);
  const [isEditorLoading, setIsEditorLoading] = useState(false);
  const [editorSaveState, setEditorSaveState] = useState<EditorSaveState>("idle");
  const [editorSaveMessage, setEditorSaveMessage] = useState("Ready");
  const editorBaseRevisionIdRef = useRef<string | null>(null);
  const editorDirtySectionRef = useRef<SectionId | null>(null);
  const editorSaveTimerRef = useRef<number | null>(null);
  const annotationBaseRevisionIdRef = useRef<string | null>(initialAnnotations.revisionId);
  const annotationSaveTimerRef = useRef<number | null>(null);

  const templateData = useMemo(
    () => buildTemplateData(activeReport.snapshot, selectedMonth, execSummary),
    [activeReport.snapshot, execSummary, selectedMonth],
  );
  const selectedTabId = useMemo(
    () => ensureTab(selectedPageId, selectedTabByPage[selectedPageId]),
    [selectedPageId, selectedTabByPage],
  );
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
  const selectedEditorSectionId = selectedPageId === DATA_ENTRY_PAGE_ID ? (selectedTabId as SectionId | null) : null;
  const pageIsExportable = useMemo(() => isExportablePageId(selectedPageId), [selectedPageId]);
  const activeSlideId = useMemo(() => getSlideId(selectedPageId, selectedTabId), [selectedPageId, selectedTabId]);
  const selectedAnnotation = useMemo(
    () => annotationState.annotations.find((annotation) => annotation.id === selectedAnnotationId) ?? null,
    [annotationState.annotations, selectedAnnotationId],
  );
  const canEditAnnotations = activeReport.id !== "demo" && pageIsExportable && !exportMode && !isSwitchingReport;

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
    selectedTabByPageRef.current = selectedTabByPage;
  }, [selectedTabByPage]);

  const loadExecSummary = useCallback(async (reportId: string, month: string) => {
    const cacheKey = `${reportId}:${month}`;
    const cached = execSummaryCacheRef.current.get(cacheKey);

    if (cached) {
      setExecSummary(cached);
      setSummaryEditorHtml(cached.contentHtml);
      setIsSummaryLoading(false);
      return;
    }

    setIsSummaryLoading(true);
    setExecSummary({
      mode: "loading",
      contentHtml: "",
      excerpt: "",
      updatedAt: null,
      sourceReportId: null,
    });

    try {
      const payload = await fetchJson<ExecSummaryApiPayload>(`/api/reports/${reportId}/exec-summary?month=${encodeURIComponent(month)}`);
      const nextSummary = payload.summary ?? {
        mode: "empty" as const,
        contentHtml: "",
        excerpt: "",
        updatedAt: null,
        sourceReportId: null,
      };

      execSummaryCacheRef.current.set(cacheKey, nextSummary);
      setExecSummary(nextSummary);
      setSummaryEditorHtml(nextSummary.contentHtml);
    } catch (error) {
      setUploadError(error instanceof Error ? error.message : "Failed to load exec summary.");
      setExecSummary({
        mode: "empty",
        contentHtml: "",
        excerpt: "",
        updatedAt: null,
        sourceReportId: null,
      });
      setSummaryEditorHtml("");
    } finally {
      setIsSummaryLoading(false);
    }
  }, []);

  const loadPrep = useCallback(async (reportId: string, month: string) => {
    const cacheKey = `${reportId}:${month}`;
    const cached = prepCacheRef.current.get(cacheKey);

    if (cached) {
      setPrepView(cached);
      setIsPrepLoading(false);
      return;
    }

    setIsPrepLoading(true);

    try {
      const payload = await fetchJson<PrepApiPayload>(`/api/reports/${reportId}/prep?month=${encodeURIComponent(month)}`);
      const nextPrep = payload.prep ?? null;
      if (nextPrep) {
        prepCacheRef.current.set(cacheKey, nextPrep);
      }
      setPrepView(nextPrep);
    } catch (error) {
      setUploadError(error instanceof Error ? error.message : "Failed to load readiness data.");
      setPrepView(null);
    } finally {
      setIsPrepLoading(false);
    }
  }, []);

  const loadAnnotations = useCallback(async (reportId: string, month: string) => {
    const cacheKey = `${reportId}:${month}`;
    const cached = annotationCacheRef.current.get(cacheKey);

    if (cached) {
      annotationBaseRevisionIdRef.current = cached.revisionId;
      setAnnotationState(cached);
      setAnnotationSaveState("idle");
      setAnnotationSaveMessage(cached.annotations.length === 0 ? "No annotations yet" : "Annotations ready");
      setSelectedAnnotationId(null);
      setAnnotationTool("select");
      setIsAnnotationsLoading(false);
      return;
    }

    setIsAnnotationsLoading(true);

    try {
      const payload = await fetchJson<AnnotationApiPayload>(`/api/reports/${reportId}/annotations?month=${encodeURIComponent(month)}`);
      const nextState = payload.annotationState ?? createEmptyReportAnnotationsState(reportId, month);
      annotationCacheRef.current.set(cacheKey, nextState);
      annotationBaseRevisionIdRef.current = nextState.revisionId;
      setAnnotationState(nextState);
      setAnnotationSaveState("idle");
      setAnnotationSaveMessage(nextState.annotations.length === 0 ? "No annotations yet" : "Annotations ready");
      setSelectedAnnotationId(null);
      setAnnotationTool("select");
    } catch (error) {
      setUploadError(error instanceof Error ? error.message : "Failed to load annotations.");
      const emptyState = createEmptyReportAnnotationsState(reportId, month);
      annotationBaseRevisionIdRef.current = emptyState.revisionId;
      setAnnotationState(emptyState);
      setAnnotationSaveState("error");
      setAnnotationSaveMessage("Unable to load annotations");
      setSelectedAnnotationId(null);
      setAnnotationTool("select");
    } finally {
      setIsAnnotationsLoading(false);
    }
  }, []);

  const syncReportFromDraft = useCallback((draft: EditableReportDraft) => {
    const nextUpdatedAt = new Date().toISOString();

    setActiveReport((current) => {
      if (current.id !== draft.manifest.reportId) {
        return current;
      }

      const nextReport: AppReportRecord = {
        ...current,
        title: draft.manifest.title,
        reportSeriesKey: draft.manifest.reportSeriesKey,
        currentMonth: draft.snapshot.currentMonth,
        availableMonths: draft.snapshot.availableMonths,
        snapshot: draft.snapshot,
        updatedAt: nextUpdatedAt,
      };

      reportCacheRef.current.set(nextReport.id, nextReport);
      activeReportRef.current = nextReport;
      return nextReport;
    });

    setReports((current) =>
      current.map((report) =>
        report.id === draft.manifest.reportId
          ? {
              ...report,
              title: draft.manifest.title,
              reportSeriesKey: draft.manifest.reportSeriesKey,
              currentMonth: draft.snapshot.currentMonth,
              availableMonths: draft.snapshot.availableMonths,
              updatedAt: nextUpdatedAt,
            }
          : report,
      ),
    );
  }, []);

  const loadEditorDraft = useCallback(
    async (reportId: string, month: string) => {
      if (reportId === "demo") {
        setEditorDraft(null);
        return;
      }

      setIsEditorLoading(true);
      try {
        const payload = await fetchJson<EditorApiPayload>(`/api/reports/${reportId}/editor?month=${encodeURIComponent(month)}`);
        if (!payload.draft) {
          throw new Error("Editor draft not available.");
        }

        setEditorDraft(payload.draft);
        editorBaseRevisionIdRef.current = payload.draft.manifest.currentRevision.revisionId;
        setEditorSaveState("idle");
        setEditorSaveMessage("Ready");
        syncReportFromDraft(payload.draft);
      } catch (error) {
        setUploadError(error instanceof Error ? error.message : "Failed to load data entry workspace.");
      } finally {
        setIsEditorLoading(false);
      }
    },
    [syncReportFromDraft],
  );

  const updateEditorDraft = useCallback(
    (updater: (current: EditableReportDraft) => EditableReportDraft) => {
      setEditorDraft((current) => {
        if (!current) {
          return current;
        }

        const nextDraft = updater(current);
        syncReportFromDraft(nextDraft);
        return nextDraft;
      });
    },
    [syncReportFromDraft],
  );

  const markEditorDirty = useCallback((sectionId: SectionId) => {
    editorDirtySectionRef.current = sectionId;
    setEditorSaveState("dirty");
    setEditorSaveMessage("Unsaved changes");
  }, []);

  const markAnnotationsDirty = useCallback(() => {
    setAnnotationSaveState("dirty");
    setAnnotationSaveMessage("Unsaved annotations");
  }, []);

  const updateAnnotations = useCallback(
    (updater: (annotations: ReportAnnotation[]) => ReportAnnotation[]) => {
      setAnnotationState((current) => ({
        ...current,
        annotations: updater(current.annotations),
      }));
      markAnnotationsDirty();
    },
    [markAnnotationsDirty],
  );

  const updateAnnotationById = useCallback(
    (annotationId: string, updater: (annotation: ReportAnnotation) => ReportAnnotation) => {
      updateAnnotations((annotations) => annotations.map((annotation) => (annotation.id === annotationId ? updater(annotation) : annotation)));
    },
    [updateAnnotations],
  );

  const handleCreateAnnotation = useCallback(
    (slideId: string, type: "bubble" | "text", point: { x: number; y: number }) => {
      const maxZIndex = annotationState.annotations.reduce((highest, annotation) => Math.max(highest, annotation.zIndex), 0);
      const width = type === "bubble" ? 0.24 : 0.18;
      const height = type === "bubble" ? 0.16 : 0.11;
      const nextAnnotation: ReportAnnotation = {
        id: createClientAnnotationId(),
        slideId,
        type,
        x: Math.max(0, Math.min(point.x - width / 2, 1 - width)),
        y: Math.max(0, Math.min(point.y - height / 2, 1 - height)),
        width,
        height,
        text: "",
        zIndex: maxZIndex + 1,
        style: DEFAULT_ANNOTATION_STYLE,
        tailAnchor:
          type === "bubble"
            ? {
                x: Math.max(0, Math.min(point.x + 0.04, 1)),
                y: Math.max(0, Math.min(point.y + 0.11, 1)),
              }
            : null,
      };

      setAnnotationState((current) => ({
        ...current,
        annotations: [...current.annotations, nextAnnotation],
      }));
      setSelectedAnnotationId(nextAnnotation.id);
      setAnnotationTool("select");
      markAnnotationsDirty();
    },
    [annotationState.annotations, markAnnotationsDirty],
  );

  const handleDeleteAnnotation = useCallback(
    (annotationId: string) => {
      updateAnnotations((annotations) => annotations.filter((annotation) => annotation.id !== annotationId));
      setSelectedAnnotationId((current) => (current === annotationId ? null : current));
    },
    [updateAnnotations],
  );

  useEffect(() => {
    setIsSummaryEditing(false);
    void loadExecSummary(activeReport.id, selectedMonth);
  }, [activeReport.id, loadExecSummary, selectedMonth]);

  useEffect(() => {
    void loadAnnotations(activeReport.id, selectedMonth);
  }, [activeReport.id, loadAnnotations, selectedMonth]);

  useEffect(() => {
    void loadPrep(activeReport.id, selectedMonth);
  }, [activeReport.id, loadPrep, selectedMonth]);

  useEffect(() => {
    if (selectedAnnotationId && !annotationState.annotations.some((annotation) => annotation.id === selectedAnnotationId && annotation.slideId === activeSlideId)) {
      setSelectedAnnotationId(null);
    }
  }, [activeSlideId, annotationState.annotations, selectedAnnotationId]);

  useEffect(() => {
    if (!pageIsExportable || exportMode) {
      setAnnotationTool("select");
      setSelectedAnnotationId(null);
    }
  }, [exportMode, pageIsExportable]);

  useEffect(() => {
    if (selectedPageId !== DATA_ENTRY_PAGE_ID || activeReport.id === "demo") {
      return;
    }

    if (editorDraft?.manifest.reportId === activeReport.id && editorDraft.snapshot.currentMonth === selectedMonth) {
      return;
    }

    void loadEditorDraft(activeReport.id, selectedMonth);
  }, [activeReport.id, editorDraft, loadEditorDraft, selectedMonth, selectedPageId]);

  useEffect(() => {
    if (selectedPageId !== DATA_ENTRY_PAGE_ID || !editorDraft || !selectedEditorSectionId) {
      return;
    }

    if (editorSaveTimerRef.current) {
      window.clearTimeout(editorSaveTimerRef.current);
    }

    if (editorSaveState !== "dirty" || !editorDirtySectionRef.current) {
      return;
    }

    editorSaveTimerRef.current = window.setTimeout(async () => {
      const dirtySection = editorDirtySectionRef.current;
      if (!dirtySection || !editorDraft) {
        return;
      }

      setEditorSaveState("saving");
      setEditorSaveMessage("Saving changes...");

      try {
        const payload = getSectionPayload(
          editorDraft.snapshot,
          dirtySection,
          editorDraft.manifest.title,
          editorDraft.manifest.reportSeriesKey,
        );
        const response = await fetch(`/api/reports/${editorDraft.manifest.reportId}/editor/sections/${dirtySection}?month=${selectedMonth}`, {
          method: "PUT",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            baseRevisionId: editorBaseRevisionIdRef.current,
            payload,
          }),
        });

        if (response.status === 409) {
          const body = (await response.json()) as { changedSections?: string[] };
          setEditorSaveState("conflict");
          setEditorSaveMessage(`Someone else changed ${body.changedSections?.join(", ") ?? "this section"}. Refresh to continue.`);
          return;
        }

        if (!response.ok) {
          const body = (await response.json()) as { error?: string };
          throw new Error(body.error ?? "Save failed.");
        }

        const body = (await response.json()) as EditorApiPayload;
        if (!body.draft) {
          throw new Error("Save succeeded but no updated draft was returned.");
        }

        editorDirtySectionRef.current = null;
        editorBaseRevisionIdRef.current = body.draft.manifest.currentRevision.revisionId;
        setEditorDraft(body.draft);
        syncReportFromDraft(body.draft);
        setEditorSaveState("saved");
        setEditorSaveMessage(`Saved revision ${body.draft.manifest.currentRevision.revisionNumber}`);
      } catch (error) {
        setEditorSaveState("error");
        setEditorSaveMessage(error instanceof Error ? error.message : "Save failed.");
      }
    }, 900);

    return () => {
      if (editorSaveTimerRef.current) {
        window.clearTimeout(editorSaveTimerRef.current);
      }
    };
  }, [editorDraft, editorSaveState, selectedEditorSectionId, selectedMonth, selectedPageId, syncReportFromDraft]);

  useEffect(() => {
    if (activeReport.id === "demo" || isAnnotationsLoading) {
      return;
    }

    if (annotationSaveTimerRef.current) {
      window.clearTimeout(annotationSaveTimerRef.current);
    }

    if (annotationSaveState !== "dirty") {
      return;
    }

    const reportId = activeReport.id;
    const reportingMonth = selectedMonth;
    const annotations = annotationState.annotations;

    annotationSaveTimerRef.current = window.setTimeout(async () => {
      setAnnotationSaveState("saving");
      setAnnotationSaveMessage("Saving annotations...");

      try {
        const response = await fetch(`/api/reports/${reportId}/annotations?month=${encodeURIComponent(reportingMonth)}`, {
          method: "PUT",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            baseRevisionId: annotationBaseRevisionIdRef.current,
            annotations,
          }),
        });

        if (response.status === 409) {
          setAnnotationSaveState("conflict");
          setAnnotationSaveMessage("Annotations changed elsewhere. Reload to continue.");
          return;
        }

        if (!response.ok) {
          const body = (await response.json()) as { error?: string };
          throw new Error(body.error ?? "Annotation save failed.");
        }

        const body = (await response.json()) as AnnotationApiPayload;
        if (!body.annotationState) {
          throw new Error("Save succeeded but no annotation state was returned.");
        }

        annotationBaseRevisionIdRef.current = body.annotationState.revisionId;
        annotationCacheRef.current.set(`${reportId}:${reportingMonth}`, body.annotationState);
        setAnnotationState(body.annotationState);
        setAnnotationSaveState("saved");
        setAnnotationSaveMessage("Annotations saved");
      } catch (error) {
        setAnnotationSaveState("error");
        setAnnotationSaveMessage(error instanceof Error ? error.message : "Annotation save failed.");
      }
    }, 650);

    return () => {
      if (annotationSaveTimerRef.current) {
        window.clearTimeout(annotationSaveTimerRef.current);
      }
    };
  }, [activeReport.id, annotationSaveState, annotationState.annotations, isAnnotationsLoading, selectedMonth]);

  useEffect(() => {
    if (selectedPageId !== DATA_ENTRY_PAGE_ID || !editorDraft) {
      return;
    }

    const reportId = editorDraft.manifest.reportId;
    let cancelled = false;

    async function heartbeat() {
      try {
        const response = await fetch(`/api/reports/${reportId}/presence?month=${selectedMonth}`, {
          method: "PUT",
        });

        if (!response.ok || cancelled) {
          return;
        }

        const body = (await response.json()) as { activePresence: EditableReportDraft["activePresence"] };
        setEditorDraft((current) => (current ? { ...current, activePresence: body.activePresence } : current));
      } catch {
        // Presence is best-effort.
      }
    }

    void heartbeat();
    const interval = window.setInterval(heartbeat, 45000);

    return () => {
      cancelled = true;
      window.clearInterval(interval);
    };
  }, [editorDraft, selectedMonth, selectedPageId]);

  useEffect(() => {
    if (selectedPageId !== DATA_ENTRY_PAGE_ID || !editorDraft || editorDraft.manifest.artifactSyncStatus.state !== "pending") {
      return;
    }

    const interval = window.setInterval(async () => {
      try {
        const payload = await fetchJson<EditorApiPayload>(
          `/api/reports/${editorDraft.manifest.reportId}/editor?month=${encodeURIComponent(selectedMonth)}`,
        );

        if (!payload.draft) {
          return;
        }

        setEditorDraft(payload.draft);
        syncReportFromDraft(payload.draft);
      } catch {
        // Polling is best-effort.
      }
    }, 5000);

    return () => window.clearInterval(interval);
  }, [editorDraft, selectedMonth, selectedPageId, syncReportFromDraft]);

  useEffect(() => {
    if (selectedPageId !== "p-summary") {
      setIsSummaryEditing(false);
    }
  }, [selectedPageId]);

  useEffect(() => {
    try {
      const storedValue = window.localStorage.getItem("ta-it-reporting-sidebar-collapsed");
      if (storedValue === "true") {
        setIsSidebarCollapsed(true);
      }

      const storedNavStyle = window.localStorage.getItem("ta-it-reporting-collapsed-nav-style");
      if (storedNavStyle === "icons" || storedNavStyle === "initials") {
        setCollapsedNavStyle(storedNavStyle);
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

  useEffect(() => {
    try {
      window.localStorage.setItem("ta-it-reporting-collapsed-nav-style", collapsedNavStyle);
    } catch {
      // localStorage access is optional
    }
  }, [collapsedNavStyle]);

  useEffect(() => {
    const handleUploadRequest = () => {
      setIsSidebarCollapsed(false);
      window.requestAnimationFrame(() => {
        fileInputRef.current?.click();
      });
    };

    window.addEventListener("ta:request-upload", handleUploadRequest);
    return () => window.removeEventListener("ta:request-upload", handleUploadRequest);
  }, []);

  const syncUrl = useCallback(
    (reportId: string, month: string, pageId: string, tabId: string | null, historyMode: "push" | "replace" = "push") => {
      const url = buildCanonicalUrl(reportId, month, pageId, tabId);
      const method = historyMode === "replace" ? "replaceState" : "pushState";
      window.history[method]({}, "", url);
    },
    [],
  );

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
        tabId?: string | null;
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
        const nextTabId = ensureTab(nextPageId, options.tabId);

        setActiveReport(report);
        setSelectedMonth(nextMonth);
        setSelectedPageId(nextPageId);
        setSelectedTabByPage((current) => ({
          ...current,
          [nextPageId]: nextTabId,
        }));
        setStatusMessage(`Viewing ${report.title}`);

        if (options.historyMode !== "none") {
          syncUrl(report.id, nextMonth, nextPageId, nextTabId, options.historyMode ?? "push");
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
    const nextTabId = ensureTab(nextPageId, params.get("tab"));

    if (nextReportId !== currentReport.id) {
      await activateReport(nextReportId, {
        month: params.get("month"),
        pageId: nextPageId,
        tabId: nextTabId,
        historyMode: "none",
      });
      return;
    }

    setSelectedMonth(ensureMonth(currentReport, params.get("month")));
    setSelectedPageId(nextPageId);
    setSelectedTabByPage((current) => ({
      ...current,
      [nextPageId]: nextTabId,
    }));
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
  }, [activeReport.id, selectedMonth, selectedPageId, selectedTabId]);

  const handlePageChange = useCallback(
    (pageId: string, tabId: string | null) => {
      const resolvedTabId = ensureTab(pageId, tabId);
      const currentTabId = ensureTab(pageId, selectedTabByPageRef.current[pageId]);

      if (selectedPageRef.current === pageId && currentTabId === resolvedTabId) {
        return;
      }

      selectedPageRef.current = pageId;
      setSelectedPageId(pageId);
      setSelectedTabByPage((current) => ({
        ...current,
        [pageId]: resolvedTabId,
      }));
      syncUrl(activeReportRef.current.id, selectedMonthRef.current, pageId, resolvedTabId);
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
    const editorRoots = ensureEditorPages(shellRoot as HTMLElement);
    const annotationRoots = Object.fromEntries(
      getReportSlides()
        .map((slide) => [slide.id, mountNode.querySelector(`#${slide.id}`)])
        .filter((entry): entry is [string, Element] => Boolean(entry[1])),
    );

    setTargets({
      toggle: mountNode.querySelector("#sidebar-toggle-slot"),
      period: mountNode.querySelector("#sidebar-period-slot"),
      utilities: mountNode.querySelector("#sidebar-app-utilities-slot"),
      reports: mountNode.querySelector("#sidebar-report-list-slot"),
      summaryControls: mountNode.querySelector("#summary-controls-slot"),
      summaryEditor: mountNode.querySelector("#summary-editor-slot"),
      annotationRoots,
      editorRoots,
    });

    controllerRef.current = initReportApp(shellRoot, {
      ChartLib: Chart,
      data: templateData,
      activeMonth: selectedMonth,
      initialPageId: selectedPageRef.current,
      initialTabId: ensureTab(selectedPageRef.current, selectedTabByPageRef.current[selectedPageRef.current]),
      showAllPages: false,
      attachGlobals: true,
      onPageChange: handlePageChange,
    });

    return () => {
      controllerRef.current?.destroy();
      controllerRef.current = null;
      mountNode.innerHTML = "";
      setTargets({
        toggle: null,
        period: null,
        utilities: null,
        reports: null,
        summaryControls: null,
        summaryEditor: null,
        annotationRoots: {},
        editorRoots: {},
      });
    };
  }, [handlePageChange, selectedMonth, templateBody, templateData]);

  useEffect(() => {
    const shellRoot = mountRef.current?.querySelector(".shell");
    if (!(shellRoot instanceof HTMLElement)) {
      return;
    }

    shellRoot.classList.toggle("sidebar-collapsed", isSidebarCollapsed);
    shellRoot.classList.toggle("sidebar-use-icons", collapsedNavStyle === "icons");
    shellRoot.classList.toggle("sidebar-use-initials", collapsedNavStyle === "initials");
  }, [collapsedNavStyle, isSidebarCollapsed]);

  useEffect(() => {
    const summaryBlock = mountRef.current?.querySelector("#summary-content-block");
    if (!(summaryBlock instanceof HTMLElement)) {
      return;
    }

    summaryBlock.classList.toggle("summary-is-editing", isSummaryEditing);
  }, [isSummaryEditing, selectedPageId]);

  useEffect(() => {
    controllerRef.current?.showPage(selectedPageId, selectedTabId);
  }, [selectedPageId, selectedTabId]);

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
      onclone: (clonedDocument, clonedElement) => {
        clonedElement.querySelectorAll?.(".export-icon").forEach((icon) => icon.remove());
        stripAnnotationAuthoringChrome(clonedDocument);
        stripAnnotationAuthoringChrome(clonedElement);
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
        "font-family:Inter,-apple-system,BlinkMacSystemFont,sans-serif",
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
            <div style="font-size:15px;font-weight:700;color:#005292;">TeacherActive · IT Reporting</div>
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
        stripAnnotationAuthoringChrome(clone);
        clone.classList.remove("exportable", "selected");
        clone.style.width = "100%";
        clone.style.position = "relative";
        wrapper.appendChild(clone);
      });

      const footer = document.createElement("div");
      footer.style.cssText =
        "border-top:1px solid #E5E7EB;padding-top:12px;display:flex;justify-content:space-between;font-size:10px;color:#9CA3AF;font-family:Inter,-apple-system,BlinkMacSystemFont,sans-serif;";
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
            `${getSlideId(selectedPageRef.current, ensureTab(selectedPageRef.current, selectedTabByPageRef.current[selectedPageRef.current]))}-selection`,
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

    const activePage = shellRoot.querySelector(`#${getSlideId(selectedPageId, selectedTabId)}`);
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
  }, [activeReport.id, exportMode, exportSingleTarget, selectedExportIds, selectedMonth, selectedPageId, selectedTabId]);

  const handleMonthChange = useCallback(
    (month: string) => {
      const nextMonth = ensureMonth(activeReportRef.current, month);
      setSelectedMonth(nextMonth);
      syncUrl(
        activeReportRef.current.id,
        nextMonth,
        selectedPageRef.current,
        ensureTab(selectedPageRef.current, selectedTabByPageRef.current[selectedPageRef.current]),
      );
    },
    [syncUrl],
  );

  const handleReportSelect = useCallback(
    async (reportId: string) => {
      await activateReport(reportId, {
        month: selectedMonthRef.current,
        pageId: selectedPageRef.current,
        tabId: selectedTabByPageRef.current[selectedPageRef.current],
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
        setSelectedTabByPage({});
        setStatusMessage(`Imported ${payload.report.originalFilename}. Open Edit data to continue in the admin workspace.`);
        syncUrl(payload.report.id, payload.report.currentMonth, nextPageId, null);
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

  const openEditor = useCallback(() => {
    if (activeReportRef.current.id === "demo") {
      setUploadError("Create or import a saved report before opening the editor.");
      return;
    }

    const nextTabId = resolveTabId(DATA_ENTRY_PAGE_ID, selectedTabByPageRef.current[DATA_ENTRY_PAGE_ID] ?? "overview-setup") ?? "overview-setup";
    setSelectedPageId(DATA_ENTRY_PAGE_ID);
    setSelectedTabByPage((current) => ({
      ...current,
      [DATA_ENTRY_PAGE_ID]: nextTabId,
    }));
    syncUrl(activeReportRef.current.id, selectedMonthRef.current, DATA_ENTRY_PAGE_ID, nextTabId);
    setStatusMessage("Opened the data entry workspace.");
  }, [syncUrl]);

  const openBlankDraftForm = useCallback(() => {
    const suggestedMonth = selectedMonthRef.current || getCurrentMonthValue();
    setBlankReportMonth(suggestedMonth);
    setBlankReportTitle(buildDefaultBlankReportTitle(suggestedMonth));
    setIsCreateBlankOpen(true);
    setUploadError(null);
    setUploadIssues([]);
    setStatusMessage(null);
  }, []);

  const createBlankReport = useCallback(async () => {
    const trimmedTitle = blankReportTitle.trim();

    if (!trimmedTitle) {
      setUploadError("Blank reports need a title.");
      return;
    }

    setIsCreatingBlank(true);
    setUploadError(null);
    setUploadIssues([]);
    setStatusMessage(null);

    try {
      const response = await fetch("/api/reports", {
        method: "POST",
        headers: {
          "content-type": "application/json",
        },
        body: JSON.stringify({
          title: trimmedTitle,
          initialMonth: blankReportMonth,
        }),
      });

      const payload = (await response.json()) as ReportApiPayload;

      if (!response.ok || !payload.report) {
        setUploadError(payload.error ?? "Unable to create a blank report.");
        setUploadIssues(payload.issues ?? []);
        return;
      }

      reportCacheRef.current.set(payload.report.id, payload.report);
      await refreshReportList(payload.report);
      setIsCreateBlankOpen(false);
      setActiveReport(payload.report);
      setSelectedMonth(payload.report.currentMonth);
      setSelectedPageId(DATA_ENTRY_PAGE_ID);
      setSelectedTabByPage({
        [DATA_ENTRY_PAGE_ID]: "overview-setup",
      });
      setEditorDraft(null);
      setStatusMessage(`Created ${payload.report.title}. Data entry is ready.`);
      syncUrl(payload.report.id, payload.report.currentMonth, DATA_ENTRY_PAGE_ID, "overview-setup");
    } catch (error) {
      setUploadError(error instanceof Error ? error.message : "Unable to create a blank report.");
    } finally {
      setIsCreatingBlank(false);
    }
  }, [blankReportMonth, blankReportTitle, refreshReportList, syncUrl]);

  const downloadExport = useCallback(async (exportType: "page-png" | "full-pdf" | "full-pptx" | "full-pptx-editable" | "full-xlsx" | "full-json") => {
    setBusyExport(exportType);
    setExportError(null);

    try {
      const activeTabId = ensureTab(selectedPageRef.current, selectedTabByPageRef.current[selectedPageRef.current]);
      const payload: Record<string, string> = {
        exportType,
        month: selectedMonthRef.current,
      };

      if (
        exportType !== "full-pdf" &&
        exportType !== "full-pptx" &&
        exportType !== "full-pptx-editable" &&
        exportType !== "full-xlsx" &&
        exportType !== "full-json"
      ) {
        payload.pageId = selectedPageRef.current;
        if (activeTabId) {
          payload.tabId = activeTabId;
        }
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

  const openSummaryEditor = useCallback(() => {
    setSummaryEditorHtml(execSummary.contentHtml);
    setIsSummaryEditing(true);
  }, [execSummary.contentHtml]);

  const cancelSummaryEditor = useCallback(() => {
    setSummaryEditorHtml(execSummary.contentHtml);
    setIsSummaryEditing(false);
  }, [execSummary.contentHtml]);

  const saveSummaryEditor = useCallback(
    async (contentHtml: string) => {
      setIsSummarySaving(true);
      setUploadError(null);
      setStatusMessage(null);

      try {
        const payload = await fetchJson<ExecSummaryApiPayload>(
          `/api/reports/${activeReportRef.current.id}/exec-summary?month=${encodeURIComponent(selectedMonthRef.current)}`,
          {
            method: "PUT",
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ contentHtml }),
          },
        );

        if (!payload.summary) {
          throw new Error("Failed to save exec summary.");
        }

        const cacheKey = `${activeReportRef.current.id}:${selectedMonthRef.current}`;
        execSummaryCacheRef.current.set(cacheKey, payload.summary);
        setExecSummary(payload.summary);
        setSummaryEditorHtml(payload.summary.contentHtml);
        setIsSummaryEditing(false);
        prepCacheRef.current.delete(cacheKey);
        void loadPrep(activeReportRef.current.id, selectedMonthRef.current);
        setStatusMessage("Exec summary saved.");
      } catch (error) {
        setUploadError(error instanceof Error ? error.message : "Failed to save exec summary.");
      } finally {
        setIsSummarySaving(false);
      }
    },
    [loadPrep],
  );

  const togglePrepDrawer = useCallback(() => {
    setIsSidebarCollapsed(false);
    setIsPrepOpen((current) => !current);
  }, []);

  const jumpToPrepTarget = useCallback(
    (pageId: string, tabId: string | null) => {
      handlePageChange(pageId, tabId);
    },
    [handlePageChange],
  );

  const saveAcknowledgedChecks = useCallback(
    async (acknowledgedCheckIds: string[]) => {
      if (activeReportRef.current.id === "demo") {
        return;
      }

      setIsPrepSaving(true);
      setUploadError(null);

      try {
        const payload = await fetchJson<PrepApiPayload>(
          `/api/reports/${activeReportRef.current.id}/prep?month=${encodeURIComponent(selectedMonthRef.current)}`,
          {
            method: "PUT",
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ acknowledgedCheckIds }),
          },
        );

        if (!payload.prep) {
          throw new Error("Failed to save readiness review state.");
        }

        const cacheKey = `${activeReportRef.current.id}:${selectedMonthRef.current}`;
        prepCacheRef.current.set(cacheKey, payload.prep);
        setPrepView(payload.prep);
      } catch (error) {
        setUploadError(error instanceof Error ? error.message : "Failed to save readiness review state.");
      } finally {
        setIsPrepSaving(false);
      }
    },
    [],
  );

  const toggleAcknowledgedCheck = useCallback(
    (checkId: string, nextAcknowledged: boolean) => {
      const currentPrep = prepView;
      if (!currentPrep) {
        return;
      }

      const nextIds = nextAcknowledged
        ? [...currentPrep.acknowledgedCheckIds, checkId]
        : currentPrep.acknowledgedCheckIds.filter((id) => id !== checkId);

      void saveAcknowledgedChecks(nextIds);
    },
    [prepView, saveAcknowledgedChecks],
  );

  const copyPreviousMonthSummary = useCallback(() => {
    const previousSummary = prepView?.rollover.previousExecSummary;
    if (!previousSummary?.available || activeReportRef.current.id === "demo") {
      return;
    }

    setSummaryEditorHtml(previousSummary.contentHtml);
    setIsSummaryEditing(true);
    selectedPageRef.current = "p-summary";
    selectedTabByPageRef.current = {
      ...selectedTabByPageRef.current,
      "p-summary": null,
    };
    setSelectedPageId("p-summary");
    setSelectedTabByPage((current) => ({
      ...current,
      "p-summary": null,
    }));
    syncUrl(activeReportRef.current.id, selectedMonthRef.current, "p-summary", null);
    setIsPrepOpen(false);
    setStatusMessage(`Copied ${previousSummary.monthLabel} summary into the editor.`);
  }, [prepView, syncUrl]);

  const periodPortal =
    targets.period &&
    createPortal(
      <MonthPicker availableMonths={activeReport.availableMonths} onChange={handleMonthChange} selectedMonth={selectedMonth} />,
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
          <span className="sidebar-field-label">Navigation</span>
          <div className="sidebar-inline">
            <button
              className={`sidebar-button secondary ${collapsedNavStyle === "icons" ? "is-active" : ""}`}
              onClick={() => setCollapsedNavStyle("icons")}
              type="button"
            >
              Icons
            </button>
            <button
              className={`sidebar-button secondary ${collapsedNavStyle === "initials" ? "is-active" : ""}`}
              onClick={() => setCollapsedNavStyle("initials")}
              type="button"
            >
              Initials
            </button>
          </div>
          <div className="sidebar-meta">Collapsed rail style</div>
        </div>

        <div className="sidebar-stack-tight">
          <span className="sidebar-field-label">Author Workspace</span>
          <button
            className={`sidebar-button ${isPrepOpen ? "primary is-active" : "secondary"}`}
            disabled={isUploading || isSwitchingReport || isPrepSaving}
            onClick={togglePrepDrawer}
            type="button"
          >
            {isPrepOpen ? "Close Readiness Center" : "Readiness & Rollover"}
          </button>
          <div className="sidebar-meta">
            {isPrepLoading || !prepView
              ? "Loading author checks..."
              : prepView.readiness.summary.status === "ready"
                ? "Ready to export"
                : `${prepView.readiness.summary.blockingCount} blocking · ${prepView.readiness.summary.warningCount} warning${prepView.readiness.summary.warningCount === 1 ? "" : "s"}`}
          </div>
        </div>

        <div className="sidebar-stack-tight">
          <span className="sidebar-field-label">Data Entry</span>
          <button
            className={`sidebar-button ${activeReport.id === "demo" ? "secondary" : "primary"}`}
            disabled={activeReport.id === "demo" || isUploading || isSwitchingReport || isCreatingBlank}
            onClick={openEditor}
            type="button"
          >
            {activeReport.id === "demo" ? "Import or Create First" : "Edit data"}
          </button>
          <div className="sidebar-meta">
            Open the tabbed admin workspace with autosave, revision tracking, and workbook sync.
          </div>
        </div>

        <div className="sidebar-stack-tight">
          <span className="sidebar-field-label">Start or Import</span>
          {isCreateBlankOpen ? (
            <div className="sidebar-stack-tight">
              <label className="sidebar-stack-tight">
                <span className="sidebar-meta">Draft title</span>
                <input
                  className="sidebar-input"
                  disabled={isCreatingBlank || isUploading || isSwitchingReport}
                  onChange={(event) => setBlankReportTitle(event.target.value)}
                  type="text"
                  value={blankReportTitle}
                />
              </label>
              <label className="sidebar-stack-tight">
                <span className="sidebar-meta">Starting month</span>
                <input
                  className="sidebar-input"
                  disabled={isCreatingBlank || isUploading || isSwitchingReport}
                  onChange={(event) => setBlankReportMonth(event.target.value)}
                  type="month"
                  value={blankReportMonth}
                />
              </label>
              <div className="sidebar-inline">
                <button
                  className="sidebar-button primary"
                  disabled={isCreatingBlank || isUploading || isSwitchingReport}
                  onClick={() => void createBlankReport()}
                  type="button"
                >
                  {isCreatingBlank ? "Creating..." : "Create blank"}
                </button>
                <button
                  className="sidebar-button secondary"
                  disabled={isCreatingBlank || isUploading || isSwitchingReport}
                  onClick={() => setIsCreateBlankOpen(false)}
                  type="button"
                >
                  Cancel
                </button>
              </div>
            </div>
          ) : (
            <button
              className="sidebar-button primary"
              disabled={isCreatingBlank || isUploading || isSwitchingReport}
              onClick={openBlankDraftForm}
              type="button"
            >
              Create blank report
            </button>
          )}
          <span className="sidebar-field-label">Workbook Upload</span>
          <input
            accept=".xlsx"
            className="sidebar-input-hidden"
            onChange={handleFileSelection}
            ref={fileInputRef}
            type="file"
          />
          <button
            className="sidebar-button secondary"
            disabled={isUploading || isSwitchingReport || isCreatingBlank}
            onClick={() => fileInputRef.current?.click()}
            type="button"
          >
            {isUploading ? "Uploading workbook..." : "Upload workbook"}
          </button>
          <a className="sidebar-link" href="/templates/IT_Exec_Reporting_Ingestion_Template_master.xlsx">
            Download master template
          </a>
        </div>

        <div className="sidebar-stack-tight">
          <span className="sidebar-field-label">Exports</span>
          <div className="sidebar-inline">
            <button
              className="sidebar-button secondary"
              disabled={busyExport !== null || busyClientExport !== null || !pageIsExportable}
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
          <div className="sidebar-inline">
            <button
              className="sidebar-button secondary"
              disabled={busyExport !== null || busyClientExport !== null}
              onClick={() => void downloadExport("full-xlsx")}
              type="button"
            >
              {busyExport === "full-xlsx" ? "Preparing..." : "Workbook"}
            </button>
            <button
              className="sidebar-button secondary"
              disabled={busyExport !== null || busyClientExport !== null}
              onClick={() => void downloadExport("full-json")}
              type="button"
            >
              {busyExport === "full-json" ? "Preparing..." : "JSON"}
            </button>
          </div>
          <div className="sidebar-inline">
            <button
              className="sidebar-button secondary"
              disabled={busyExport !== null || busyClientExport !== null}
              onClick={() => void downloadExport("full-pptx")}
              type="button"
            >
              {busyExport === "full-pptx" ? "Rendering..." : "Visual PPTX"}
            </button>
            <button
              className="sidebar-button secondary"
              disabled={busyExport !== null || busyClientExport !== null}
              onClick={() => void downloadExport("full-pptx-editable")}
              type="button"
            >
              {busyExport === "full-pptx-editable" ? "Rendering..." : "Editable PPTX"}
            </button>
          </div>
          <div className="sidebar-meta">Annotations are included in visual exports now. Editable PPTX support is phase 2.</div>
          <button
            className={`sidebar-button ${exportMode ? "primary is-active" : "secondary"}`}
            disabled={busyExport !== null || busyClientExport !== null || activeExportTargets.length === 0 || !pageIsExportable}
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
          <div className="sidebar-report-meta">
            <div className="sidebar-report-sub">Prototype snapshot · 2026-06</div>
            <div className="sidebar-report-chip">Demo</div>
          </div>
        </button>

        {reports.length === 0 ? (
          <div className="sidebar-stack-tight">
            <div className="sidebar-empty">No saved reports yet. Start blank for direct UI entry, or import a workbook to seed the draft.</div>
            <div className="sidebar-inline">
              <button
                className="sidebar-button primary"
                disabled={isUploading || isSwitchingReport || isCreatingBlank}
                onClick={openBlankDraftForm}
                type="button"
              >
                Create blank
              </button>
              <button
                className="sidebar-button secondary"
                disabled={isUploading || isSwitchingReport || isCreatingBlank}
                onClick={() => fileInputRef.current?.click()}
                type="button"
              >
                Import workbook
              </button>
            </div>
          </div>
        ) : (
          reports.map((report) => (
            <button
              className={`sidebar-report-item ${activeReport.id === report.id ? "active" : ""}`}
              disabled={isSwitchingReport || isUploading}
              key={report.id}
              onClick={() => void handleReportSelect(report.id)}
              title={report.title}
              type="button"
            >
              <div className="sidebar-report-title">{formatSidebarReportTitle(report.title, report.currentMonth)}</div>
              <div className="sidebar-report-meta">
                <div className="sidebar-report-sub">{formatMonthLabel(report.currentMonth)}</div>
                <div className="sidebar-report-chip">v{report.templateVersion}</div>
              </div>
            </button>
          ))
        )}
      </div>,
      targets.reports,
    );

  const summaryIsReadOnly = activeReport.id === "demo" || execSummary.mode === "demo-readonly";
  const summaryControlsPortal =
    targets.summaryControls &&
    createPortal(
      <div className="summary-controls">
        {summaryIsReadOnly ? (
          <span className="summary-readonly-pill">Bundled example · read only</span>
        ) : isSummaryEditing ? null : (
          <button
            className="summary-action-btn primary"
            disabled={isSummaryLoading || isSummarySaving || isSwitchingReport}
            onClick={openSummaryEditor}
            type="button"
          >
            {execSummary.mode === "empty"
              ? "Add exec summary"
              : execSummary.mode === "carried-forward"
                ? "Review inherited draft"
                : "Edit summary"}
          </button>
        )}
      </div>,
      targets.summaryControls,
    );

  const summaryEditorPortal =
    targets.summaryEditor &&
    createPortal(
      isSummaryEditing ? (
        <ExecSummaryEditor
          initialHtml={summaryEditorHtml}
          isSaving={isSummarySaving}
          onCancel={cancelSummaryEditor}
          onSave={saveSummaryEditor}
        />
      ) : null,
      targets.summaryEditor,
    );

  const activeEditorRoot = selectedEditorSectionId ? targets.editorRoots[selectedEditorSectionId] ?? null : null;
  const editorPortal =
    activeEditorRoot &&
    createPortal(
      activeReport.id === "demo" ? (
        <div className="block">
          <div className="bh">
            <div>
              <div className="bh-title">Data Entry</div>
              <div className="bh-sub">Create or import a saved report to use the data-entry workspace.</div>
            </div>
          </div>
          <div className="bb">
            <div className="prep-empty-copy">The bundled demo stays read-only. Create a blank report or import a workbook to start editing.</div>
          </div>
        </div>
      ) : isEditorLoading || !editorDraft || !selectedEditorSectionId ? (
        <div className="block">
          <div className="bh">
            <div>
              <div className="bh-title">Loading Data Entry</div>
              <div className="bh-sub">Fetching the latest draft, revision metadata, and presence state for this report.</div>
            </div>
          </div>
          <div className="bb">
            <div className="prep-empty-copy">Loading the admin workspace…</div>
          </div>
        </div>
      ) : (
        <ReportEmbeddedEditor
          draft={editorDraft}
          key={selectedEditorSectionId}
          onDraftChange={updateEditorDraft}
          onMarkDirty={markEditorDirty}
          onSelectedMonthChange={handleMonthChange}
          saveMessage={editorSaveMessage}
          saveState={editorSaveState}
          sectionId={selectedEditorSectionId}
          selectedMonth={selectedMonth}
        />
      ),
      activeEditorRoot,
    );

  const annotationPortals = Object.entries(targets.annotationRoots).map(([slideId, root]) =>
    createPortal(
      <ReportAnnotationLayer
        activeTool={slideId === activeSlideId ? annotationTool : "select"}
        annotations={annotationState.annotations}
        canEdit={canEditAnnotations && slideId === activeSlideId && !isAnnotationsLoading}
        onCreate={handleCreateAnnotation}
        onDelete={handleDeleteAnnotation}
        onSelect={setSelectedAnnotationId}
        onUpdate={updateAnnotationById}
        selectedAnnotationId={slideId === activeSlideId ? selectedAnnotationId : null}
        slideId={slideId}
      />,
      root,
      slideId,
    ),
  );

  return (
    <>
      <div ref={mountRef} />
      {pageIsExportable ? (
        <ReportAnnotationToolbar
          activeTool={annotationTool}
          canEdit={canEditAnnotations && !isAnnotationsLoading}
          onDeleteSelected={() => {
            if (selectedAnnotationId) {
              handleDeleteAnnotation(selectedAnnotationId);
            }
          }}
          onToolChange={setAnnotationTool}
          onUpdateSelected={(updater) => {
            if (selectedAnnotationId) {
              updateAnnotationById(selectedAnnotationId, updater);
            }
          }}
          saveMessage={isAnnotationsLoading ? "Loading annotations..." : annotationSaveMessage}
          saveState={isAnnotationsLoading ? "idle" : annotationSaveState}
          selectedAnnotation={selectedAnnotation?.slideId === activeSlideId ? selectedAnnotation : null}
        />
      ) : null}
      {togglePortal}
      {periodPortal}
      {utilitiesPortal}
      {reportsPortal}
      {summaryControlsPortal}
      {summaryEditorPortal}
      {editorPortal}
      {annotationPortals}
      <ReportPrepDrawer
        activeTab={activePrepTab}
        isLoading={isPrepLoading}
        isOpen={isPrepOpen}
        isSaving={isPrepSaving}
        onClose={() => setIsPrepOpen(false)}
        onCopyPreviousSummary={copyPreviousMonthSummary}
        onJump={jumpToPrepTarget}
        onTabChange={setActivePrepTab}
        onToggleAcknowledged={toggleAcknowledgedCheck}
        prep={prepView}
      />
    </>
  );
}
