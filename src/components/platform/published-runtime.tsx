"use client";

import Image from "next/image";
import Link from "next/link";
import { useCallback, useEffect, useEffectEvent, useMemo, useRef, useState, useTransition, type ReactNode } from "react";

import { getSampleFieldValue } from "@/lib/platform/designer";
import { formatPlatformDateTime } from "@/lib/platform/format";
import { findLayoutForPage, findObjectDefinition, findPageDefinition } from "@/lib/platform/manifest";
import { getThemeCssVariables } from "@/lib/platform/theme";
import type {
  AgentDefinition,
  FieldDefinition,
  LayoutComponentDefinition,
  MenuItemDefinition,
  ObjectDefinition,
  PlatformManifest,
  PlatformRecord,
  PlatformRole,
  PlatformViewAsState,
  PlatformWorkflowRunRecord,
} from "@/lib/platform/types";

import styles from "./platform-shell.module.css";

type ToastKind = "success" | "error" | "info";

interface ToastAction {
  label: string;
  onClick: () => void;
}

interface ToastRecord {
  id: string;
  kind: ToastKind;
  text: string;
  createdAt: number;
  dismissed: boolean;
  paused: boolean;
  durationMs: number;
  remainingMs: number;
  lastResumedAt: number;
  cycle: number;
  action?: ToastAction;
  onExpire?: () => Promise<void> | void;
}

interface RuntimeTableSort {
  fieldKey: string;
  direction: "asc" | "desc";
}

interface PublishedRuntimeProps {
  manifest: PlatformManifest;
  requestedRoute?: string;
  tenantSlug: string;
  mode?: "published" | "draft-preview" | "admin-preview";
  viewAs?: PlatformViewAsState | null;
}

function metricValue(manifest: PlatformManifest, metric: string): string {
  switch (metric) {
    case "objects":
      return String(manifest.objects.length);
    case "pages":
      return String(manifest.pages.length);
    case "workflows":
      return String(manifest.workflows.length);
    case "agents":
      return String(manifest.agents.length);
    default:
      return "0";
  }
}

function fieldInputType(field: FieldDefinition): string {
  if (field.type === "number" || field.type === "currency") {
    return "number";
  }

  if (field.type === "date") {
    return "date";
  }

  if (field.type === "datetime") {
    return "datetime-local";
  }

  if (field.type === "boolean") {
    return "checkbox";
  }

  return "text";
}

async function fetchJson<T>(input: RequestInfo, init?: RequestInit): Promise<T> {
  const response = await fetch(input, init);
  const payload = (await response.json()) as T & { error?: string };

  if (!response.ok) {
    throw new Error(payload.error ?? "Request failed.");
  }

  return payload;
}

function getInitialDrafts(objectKeys: string[]): Record<string, Record<string, unknown>> {
  return Object.fromEntries(objectKeys.map((objectKey) => [objectKey, {}]));
}

function renderFieldValue(value: unknown): string {
  if (value === null || value === undefined || value === "") {
    return "—";
  }

  if (typeof value === "boolean") {
    return value ? "Yes" : "No";
  }

  if (typeof value === "number") {
    return Number.isInteger(value) ? String(value) : value.toFixed(2);
  }

  return String(value);
}

function glyphLabel(value: string): string {
  const words = value
    .split(/[^a-z0-9]+/i)
    .filter(Boolean)
    .slice(0, 2);

  if (words.length === 0) {
    return "RT";
  }

  if (words.length === 1) {
    return words[0].slice(0, 2).toUpperCase();
  }

  return words.map((word) => word[0]?.toUpperCase() ?? "").join("");
}

export function PublishedRuntime({ manifest, requestedRoute, tenantSlug, mode = "published", viewAs = null }: PublishedRuntimeProps) {
  const runtimeRole: PlatformRole = viewAs?.role ?? (mode === "published" ? "USER" : "SUPER_ADMIN");
  const orderedMenus = [...manifest.menus]
    .filter((menu) => !menu.visibleToRoles?.length || menu.visibleToRoles.includes(runtimeRole))
    .sort((left, right) => left.order - right.order);
  const brandLogo = manifest.branding.assets.find((asset) => asset.id === manifest.branding.logoAssetId);
  const firstRoute =
    manifest.appShell.defaultLandingPageKey ??
    manifest.pages.find((page) => page.isHome)?.key ??
    orderedMenus[0]?.pageKey;
  const requestedPage =
    manifest.pages.find((page) => page.route === requestedRoute) ??
    (requestedRoute ? findPageDefinition(manifest, requestedRoute) : undefined) ??
    (firstRoute ? findPageDefinition(manifest, firstRoute) : undefined) ??
    manifest.pages[0];
  const [activePageKey, setActivePageKey] = useState(requestedPage?.key ?? manifest.pages[0]?.key ?? "");
  const [recordsByObject, setRecordsByObject] = useState<Record<string, PlatformRecord[]>>({});
  const [workflowRunsByWorkflow, setWorkflowRunsByWorkflow] = useState<Record<string, PlatformWorkflowRunRecord[]>>({});
  const [draftsByObject, setDraftsByObject] = useState<Record<string, Record<string, unknown>>>(() =>
    getInitialDrafts(manifest.objects.map((objectDefinition) => objectDefinition.key)),
  );
  const [toasts, setToasts] = useState<ToastRecord[]>([]);
  const [collapsedRuntimeSections, setCollapsedRuntimeSections] = useState<Record<string, boolean>>({
    announcements: false,
    navigation: false,
    status: false,
    manifest: true,
    actions: false,
  });
  const [scrolledPast, setScrolledPast] = useState(false);
  const [loadingRecordObjectKeys, setLoadingRecordObjectKeys] = useState<string[]>([]);
  const [loadingWorkflowKeys, setLoadingWorkflowKeys] = useState<string[]>([]);
  const [recordQueryByObject, setRecordQueryByObject] = useState<Record<string, string>>({});
  const [visibleColumnsByObject, setVisibleColumnsByObject] = useState<Record<string, string[]>>({});
  const [sortByObject, setSortByObject] = useState<Record<string, RuntimeTableSort>>({});
  const [openColumnPickerObjectKey, setOpenColumnPickerObjectKey] = useState<string | null>(null);
  const [pendingDeleteRecordIds, setPendingDeleteRecordIds] = useState<Record<string, string>>({});
  const [isPending, startTransition] = useTransition();
  const toastTimersRef = useRef<Map<string, number>>(new Map());
  const toastsRef = useRef<ToastRecord[]>([]);
  const runtimeUrlRef = useRef<string | null>(null);

  const activePage = manifest.pages.find((page) => page.key === activePageKey) ?? manifest.pages[0];
  const activeLayout = activePage ? findLayoutForPage(manifest, activePage) : undefined;
  const publishedLabel = formatPlatformDateTime(manifest.metadata.publishedAt);
  const pageObjectKeys = useMemo(() => {
    if (!activeLayout) {
      return [];
    }

    return Array.from(
      new Set(
        activeLayout.sections.flatMap((section) =>
          section.components
            .map((component) => component.binding?.objectKey ?? component.objectKey)
            .filter((objectKey): objectKey is string => Boolean(objectKey)),
        ),
      ),
    );
  }, [activeLayout]);
  const pageWorkflowIds = useMemo(() => {
    if (!activeLayout) {
      return [];
    }

    return Array.from(
      new Set(
        activeLayout.sections.flatMap((section) =>
          section.components
            .map((component) => component.binding?.workflowKey ?? component.workflowKey)
            .filter((workflowKey): workflowKey is string => Boolean(workflowKey)),
        ),
      ),
    );
  }, [activeLayout]);
  const previewRecordsByObject = useMemo(() => {
    if (mode === "published") {
      return {};
    }

    return Object.fromEntries(
      pageObjectKeys.map((objectKey) => {
        const objectDefinition = findObjectDefinition(manifest, objectKey);
        if (!objectDefinition) {
          return [objectKey, []];
        }

        const rows: PlatformRecord[] = Array.from({ length: 3 }).map((_, index) => {
          const data = Object.fromEntries(
            objectDefinition.fields.map((field) => [field.key, getSampleFieldValue(objectDefinition, index, field.key)]),
          );

          return {
            id: `${objectKey}-preview-${index + 1}`,
            objectKey,
            data,
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString(),
          };
        });

        return [objectKey, rows];
      }),
    );
  }, [manifest, mode, pageObjectKeys]);

  async function clearViewAs(): Promise<void> {
    await fetch(`/api/platform/tenants/${tenantSlug}/view-as`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
      },
      body: JSON.stringify({
        active: false,
      }),
    });
    window.location.reload();
  }

  const runtimeBasePath = mode === "published" ? `/platform/runtime/${tenantSlug}` : mode === "admin-preview" ? `/platform/admin-preview/${tenantSlug}` : `/platform/preview/${tenantSlug}`;

  const buildRuntimeUrl = useCallback(
    (pageKey: string): string => {
      const page = findPageDefinition(manifest, pageKey);
      const route = page?.route ? `/${page.route}` : "";
      return `${runtimeBasePath}${route}`;
    },
    [manifest, runtimeBasePath],
  );

  function dismissToast(toastId: string, immediate = false) {
    const clearTimer = toastTimersRef.current.get(toastId);
    if (clearTimer) {
      window.clearTimeout(clearTimer);
      toastTimersRef.current.delete(toastId);
    }

    setToasts((current) => current.map((toast) => (toast.id === toastId ? { ...toast, dismissed: true, paused: true } : toast)));

    const removeToast = () => {
      setToasts((current) => current.filter((toast) => toast.id !== toastId));
    };

    if (immediate) {
      removeToast();
      return;
    }

    window.setTimeout(removeToast, 220);
  }

  async function expireToast(toastId: string): Promise<void> {
    const toast = toastsRef.current.find((candidate) => candidate.id === toastId);
    if (!toast) {
      return;
    }

    try {
      await toast.onExpire?.();
    } catch (caughtError) {
      addToast("error", caughtError instanceof Error ? caughtError.message : "Action failed.");
    } finally {
      dismissToast(toastId);
    }
  }

  function scheduleToastExpiry(toastId: string, delay: number): void {
    const existingTimer = toastTimersRef.current.get(toastId);
    if (existingTimer) {
      window.clearTimeout(existingTimer);
    }

    const nextTimer = window.setTimeout(() => {
      void expireToast(toastId);
    }, delay);
    toastTimersRef.current.set(toastId, nextTimer);
  }

  function addToast(kind: ToastKind, text: string, options?: { durationMs?: number; action?: ToastAction; onExpire?: () => Promise<void> | void }) {
    const toastId = `runtime-toast-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const durationMs = options?.durationMs ?? 6000;
    const now = Date.now();
    setToasts((current) => [
      ...current,
      {
        id: toastId,
        kind,
        text,
        createdAt: now,
        dismissed: false,
        paused: false,
        durationMs,
        remainingMs: durationMs,
        lastResumedAt: now,
        cycle: 0,
        action: options?.action,
        onExpire: options?.onExpire,
      },
    ]);
    scheduleToastExpiry(toastId, durationMs);
    return toastId;
  }

  function pauseToast(toastId: string): void {
    const toast = toastsRef.current.find((candidate) => candidate.id === toastId);
    if (!toast || toast.paused || toast.dismissed) {
      return;
    }

    const elapsed = Date.now() - toast.lastResumedAt;
    const remainingMs = Math.max(0, toast.remainingMs - elapsed);
    const existingTimer = toastTimersRef.current.get(toastId);
    if (existingTimer) {
      window.clearTimeout(existingTimer);
      toastTimersRef.current.delete(toastId);
    }

    setToasts((current) =>
      current.map((candidate) => (candidate.id === toastId ? { ...candidate, paused: true, remainingMs } : candidate)),
    );
  }

  function resumeToast(toastId: string): void {
    const toast = toastsRef.current.find((candidate) => candidate.id === toastId);
    if (!toast || !toast.paused || toast.dismissed) {
      return;
    }

    const now = Date.now();
    setToasts((current) =>
      current.map((candidate) =>
        candidate.id === toastId
          ? {
              ...candidate,
              paused: false,
              lastResumedAt: now,
              cycle: candidate.cycle + 1,
            }
          : candidate,
      ),
    );
    scheduleToastExpiry(toastId, toast.remainingMs);
  }

  function setMessage(text: string | null) {
    if (text) {
      addToast("success", text);
    }
  }

  function setError(text: string | null) {
    if (text) {
      addToast("error", text);
    }
  }

  function setPageAndSync(pageKey: string): void {
    setActivePageKey(pageKey);
  }

  const reportRuntimeError = useEffectEvent((text: string) => {
    addToast("error", text);
  });

  useEffect(() => {
    toastsRef.current = toasts;
  }, [toasts]);

  useEffect(() => {
    const timers = toastTimersRef.current;
    return () => {
      for (const timer of timers.values()) {
        window.clearTimeout(timer);
      }
    };
  }, []);

  useEffect(() => {
    function handleScroll() {
      setScrolledPast(window.scrollY > 100);
    }

    handleScroll();
    window.addEventListener("scroll", handleScroll, { passive: true });
    return () => window.removeEventListener("scroll", handleScroll);
  }, []);

  useEffect(() => {
    if (typeof window === "undefined") {
      return;
    }

    const applyPageFromLocation = () => {
      const pathname = window.location.pathname;
      const locatedPage =
        manifest.pages.find((page) => pathname === buildRuntimeUrl(page.key)) ??
        manifest.pages.find((page) => pathname.endsWith(`/${page.route}`));
      if (locatedPage) {
        setActivePageKey(locatedPage.key);
      }
      runtimeUrlRef.current = `${window.location.pathname}${window.location.search}`;
    };

    applyPageFromLocation();
    const handlePopState = () => applyPageFromLocation();
    window.addEventListener("popstate", handlePopState);
    return () => window.removeEventListener("popstate", handlePopState);
  }, [buildRuntimeUrl, manifest.pages, mode, tenantSlug]);

  useEffect(() => {
    if (typeof window === "undefined") {
      return;
    }

    const nextUrl = buildRuntimeUrl(activePageKey);
    if (runtimeUrlRef.current === nextUrl) {
      return;
    }

    window.history.pushState({}, "", nextUrl);
    runtimeUrlRef.current = nextUrl;
  }, [activePageKey, buildRuntimeUrl, mode, tenantSlug]);

  useEffect(() => {
    function handlePointerDown(event: MouseEvent) {
      const target = event.target;
      if (!(target instanceof HTMLElement) || !target.closest("[data-column-picker-root]")) {
        setOpenColumnPickerObjectKey(null);
      }
    }

    if (!openColumnPickerObjectKey) {
      return;
    }

    document.addEventListener("mousedown", handlePointerDown);
    return () => document.removeEventListener("mousedown", handlePointerDown);
  }, [openColumnPickerObjectKey]);

  useEffect(() => {
    if (mode !== "published") {
      return;
    }

    let cancelled = false;

    async function loadRecords(): Promise<void> {
      try {
        setLoadingRecordObjectKeys(pageObjectKeys);
        const entries = await Promise.all(
          pageObjectKeys.map(async (objectKey) => {
            const payload = await fetchJson<{ records: PlatformRecord[] }>(`/api/platform/tenants/${tenantSlug}/records/${objectKey}`);
            return [objectKey, payload.records] as const;
          }),
        );

        if (!cancelled) {
          setRecordsByObject((current) => ({
            ...current,
            ...Object.fromEntries(entries),
          }));
        }
      } catch (caughtError) {
        if (!cancelled) {
          reportRuntimeError(caughtError instanceof Error ? caughtError.message : "Failed to load runtime records.");
        }
      } finally {
        if (!cancelled) {
          setLoadingRecordObjectKeys([]);
        }
      }
    }

    if (pageObjectKeys.length > 0) {
      void loadRecords();
    }

    return () => {
      cancelled = true;
    };
  }, [mode, pageObjectKeys, tenantSlug]);

  useEffect(() => {
    if (mode !== "published") {
      return;
    }

    let cancelled = false;

    async function loadWorkflowRuns(): Promise<void> {
      try {
        setLoadingWorkflowKeys(pageWorkflowIds);
        const entries = await Promise.all(
          pageWorkflowIds.map(async (workflowId) => {
            const payload = await fetchJson<{ runs: PlatformWorkflowRunRecord[] }>(
              `/api/platform/tenants/${tenantSlug}/workflows/${workflowId}/runs`,
            );
            return [workflowId, payload.runs] as const;
          }),
        );

        if (!cancelled) {
          setWorkflowRunsByWorkflow((current) => ({
            ...current,
            ...Object.fromEntries(entries),
          }));
        }
      } catch (caughtError) {
        if (!cancelled) {
          reportRuntimeError(caughtError instanceof Error ? caughtError.message : "Failed to load workflow runs.");
        }
      } finally {
        if (!cancelled) {
          setLoadingWorkflowKeys([]);
        }
      }
    }

    if (pageWorkflowIds.length > 0) {
      void loadWorkflowRuns();
    }

    return () => {
      cancelled = true;
    };
  }, [mode, pageWorkflowIds, tenantSlug]);

  async function refreshRecords(objectKey: string): Promise<void> {
    setLoadingRecordObjectKeys((current) => Array.from(new Set([...current, objectKey])));
    try {
      const payload = await fetchJson<{ records: PlatformRecord[] }>(`/api/platform/tenants/${tenantSlug}/records/${objectKey}`);
      setRecordsByObject((current) => ({
        ...current,
        [objectKey]: payload.records,
      }));
    } finally {
      setLoadingRecordObjectKeys((current) => current.filter((candidate) => candidate !== objectKey));
    }
  }

  async function handleCreateRecord(objectDefinition: ObjectDefinition): Promise<void> {
    if (mode !== "published") {
      setMessage("Draft preview is read-only. Publish before creating records.");
      return;
    }

    const draft = draftsByObject[objectDefinition.key] ?? {};

    try {
      setError(null);
      setMessage(null);
      await fetchJson<{ record: PlatformRecord }>(`/api/platform/tenants/${tenantSlug}/records/${objectDefinition.key}`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
        },
        body: JSON.stringify({
          data: draft,
        }),
      });

      setDraftsByObject((current) => ({
        ...current,
        [objectDefinition.key]: {},
      }));
      await refreshRecords(objectDefinition.key);
      setMessage(`Created ${objectDefinition.label}.`);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : `Failed to create ${objectDefinition.label}.`);
    }
  }

  async function handleDeleteRecord(objectKey: string, recordId: string): Promise<void> {
    if (mode !== "published") {
      setMessage("Draft preview is read-only. Publish before deleting records.");
      return;
    }

    const pendingKey = `${objectKey}:${recordId}`;
    setPendingDeleteRecordIds((current) => ({
      ...current,
      [pendingKey]: recordId,
    }));

    addToast("info", "Record deleted.", {
      durationMs: 5000,
      action: {
        label: "Undo",
        onClick: () => {
          setPendingDeleteRecordIds((current) => {
            const next = { ...current };
            delete next[pendingKey];
            return next;
          });
        },
      },
      onExpire: async () => {
        try {
          await fetchJson<{ ok: true }>(`/api/platform/tenants/${tenantSlug}/records/${objectKey}/${recordId}`, {
            method: "DELETE",
          });
          await refreshRecords(objectKey);
          setPendingDeleteRecordIds((current) => {
            const next = { ...current };
            delete next[pendingKey];
            return next;
          });
        } catch (caughtError) {
          setPendingDeleteRecordIds((current) => {
            const next = { ...current };
            delete next[pendingKey];
            return next;
          });
          throw caughtError;
        }
      },
    });
  }

  async function refreshWorkflowRuns(workflowId: string): Promise<void> {
    setLoadingWorkflowKeys((current) => Array.from(new Set([...current, workflowId])));
    try {
      const payload = await fetchJson<{ runs: PlatformWorkflowRunRecord[] }>(`/api/platform/tenants/${tenantSlug}/workflows/${workflowId}/runs`);
      setWorkflowRunsByWorkflow((current) => ({
        ...current,
        [workflowId]: payload.runs,
      }));
    } finally {
      setLoadingWorkflowKeys((current) => current.filter((candidate) => candidate !== workflowId));
    }
  }

  async function handleQueueWorkflow(workflowId: string): Promise<void> {
    if (mode !== "published") {
      setMessage("Draft preview is read-only. Publish before queueing workflows.");
      return;
    }

    try {
      setError(null);
      setMessage(null);
      await fetchJson(`/api/platform/tenants/${tenantSlug}/workflows/${workflowId}/runs`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
        },
        body: JSON.stringify({
          payload: {
            launchedFrom: "runtime",
            requestedAt: new Date().toISOString(),
          },
        }),
      });
      await refreshWorkflowRuns(workflowId);
      setMessage("Workflow queued.");
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Failed to queue workflow.");
    }
  }

  function updateDraft(objectKey: string, fieldKey: string, value: unknown): void {
    setDraftsByObject((current) => ({
      ...current,
      [objectKey]: {
        ...(current[objectKey] ?? {}),
        [fieldKey]: value,
      },
    }));
  }

  const resolvedRecordsByObject = mode === "published" ? recordsByObject : previewRecordsByObject;
  const groupedMenus = manifest.appShell.menuGroups.map((group) => ({
    ...group,
    items: orderedMenus.filter((menu) => (menu.groupKey ?? menu.group.toLowerCase()) === group.key || menu.group === group.label),
  })).filter((group) => group.items.length > 0);

  function componentSpan(component: LayoutComponentDefinition): number {
    return component.placement?.responsive?.desktopSpan ?? component.placement?.span ?? component.width ?? 12;
  }

  function componentStyle(component: LayoutComponentDefinition) {
    return { gridColumn: `span ${componentSpan(component)}` };
  }

  function resolveComponentObjectKey(component: LayoutComponentDefinition): string | undefined {
    return component.binding?.objectKey ?? component.objectKey;
  }

  function resolveComponentWorkflowKey(component: LayoutComponentDefinition): string | undefined {
    return component.binding?.workflowKey ?? component.workflowKey;
  }

  function resolveComponentAgent(component: LayoutComponentDefinition): AgentDefinition | undefined {
    const agentId = component.binding?.agentId ?? component.agentId;
    if (!agentId) {
      return undefined;
    }

    return manifest.agents.find((candidate) => candidate.id === agentId || candidate.key === agentId);
  }

  function resolveBadgeValue(menu: MenuItemDefinition): string | null {
    const binding = manifest.appShell.badgeBindings.find((candidate) => candidate.key === menu.badgeBindingKey);
    if (!binding) {
      return null;
    }

    if (binding.metric === "draft_changes") {
      return mode === "published" ? null : "draft";
    }

    if (binding.metric === "records" && binding.objectKey) {
      return String((resolvedRecordsByObject[binding.objectKey] ?? []).length);
    }

    if (binding.workflowKey) {
      const runs = workflowRunsByWorkflow[binding.workflowKey] ?? [];
      if (binding.metric === "queued_runs") {
        return String(runs.filter((run) => run.status === "QUEUED").length);
      }
      if (binding.metric === "failed_runs") {
        return String(runs.filter((run) => run.status === "FAILED").length);
      }
    }

    return null;
  }

  function renderBindingError(component: LayoutComponentDefinition, title: string, detail: string) {
    return (
      <section className={styles.runtimeCard} key={component.id} style={componentStyle(component)}>
        <div className={styles.sectionHeader}>
          <div>
            <p className={styles.cardEyebrow}>Binding issue</p>
            <h2>{title}</h2>
          </div>
        </div>
        <div className={styles.errorBanner}>{detail}</div>
      </section>
    );
  }

  function renderEmptyState(title: string, hint: string, ctaLabel?: string, ctaAction?: () => void) {
    return (
      <div className={styles.emptyState}>
        <div className={styles.emptyStateIcon} aria-hidden="true">
          <svg fill="none" viewBox="0 0 24 24">
            <path d="M4 6h16v10H15l-2 3-2-3H4z" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.6" />
          </svg>
        </div>
        <p>{title}</p>
        <p className={styles.emptyStateHint}>{hint}</p>
        {ctaLabel && ctaAction ? (
          <button className={styles.emptyStateCta} onClick={ctaAction} type="button">
            {ctaLabel}
          </button>
        ) : null}
      </div>
    );
  }

  function renderToastStack() {
    if (toasts.length === 0) {
      return null;
    }

    return (
      <div aria-live="polite" className={styles.toastStack}>
        {toasts.map((toast) => (
          <article
            className={[
              styles.toast,
              toast.kind === "success" ? styles.toastSuccess : toast.kind === "error" ? styles.toastError : styles.toastInfo,
              toast.dismissed ? styles.toastExit : styles.toastEnter,
            ].join(" ")}
            key={toast.id}
            onMouseEnter={() => pauseToast(toast.id)}
            onMouseLeave={() => resumeToast(toast.id)}
          >
            <div className={styles.toastCopy}>
              <strong>{toast.kind === "success" ? "Success" : toast.kind === "error" ? "Error" : "Heads up"}</strong>
              <p>{toast.text}</p>
            </div>
            <div className={styles.toastActions}>
              {toast.action ? (
                <button
                  className={styles.toastUndoButton}
                  onClick={() => {
                    toast.action?.onClick();
                    dismissToast(toast.id, true);
                  }}
                  type="button"
                >
                  {toast.action.label}
                </button>
              ) : null}
              <button aria-label="Dismiss notification" className={styles.toastDismiss} onClick={() => dismissToast(toast.id, Boolean(toast.onExpire))} type="button">
                <svg fill="none" viewBox="0 0 16 16">
                  <path d="m4 4 8 8M12 4 4 12" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
                </svg>
              </button>
            </div>
            <div
              className={styles.toastProgressBar}
              key={`${toast.id}-${toast.cycle}`}
              style={{
                animationDuration: `${toast.remainingMs}ms`,
                animationPlayState: toast.paused ? "paused" : "running",
              }}
            />
          </article>
        ))}
      </div>
    );
  }

  function renderComponent(component: LayoutComponentDefinition) {
    if (component.kind === "hero") {
      return (
        <section className={styles.runtimeHero} key={component.id} style={componentStyle(component)}>
          <p className={styles.eyebrow}>{String(component.props.eyebrow ?? manifest.tenant.name)}</p>
          <h1>{component.title}</h1>
          {component.description ? <p>{component.description}</p> : null}
        </section>
      );
    }

    if (component.kind === "text" || component.kind === "rich_text" || component.kind === "callout") {
      return (
        <section className={component.kind === "callout" ? styles.runtimeCallout : styles.runtimeCard} key={component.id} style={componentStyle(component)}>
          <h2>{component.title}</h2>
          <p>{String(component.props.body ?? component.description ?? "")}</p>
          {component.kind === "callout" ? <span className={styles.inlineTag}>{String(component.props.tone ?? "info")}</span> : null}
        </section>
      );
    }

    if (component.kind === "stats" || component.kind === "stat_tiles") {
      const metrics = Array.isArray(component.props.metrics)
        ? (component.props.metrics as Array<{ label: string; metric: string }>)
        : [];

      return (
        <section className={styles.runtimeCard} key={component.id} style={componentStyle(component)}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>{mode === "published" ? "Published Runtime" : "Draft Preview"}</p>
              <h2>{component.title}</h2>
            </div>
          </div>
          <div className={styles.metricGrid}>
            {metrics.map((metric) => (
              <article className={`${styles.metricCard} ${styles.staggerChild}`} key={metric.label}>
                <span>{metric.label}</span>
                <strong>{metricValue(manifest, metric.metric)}</strong>
              </article>
            ))}
          </div>
        </section>
      );
    }

    if ((component.kind === "record_table" || component.kind === "related_records") && resolveComponentObjectKey(component)) {
      const objectDefinition = findObjectDefinition(manifest, resolveComponentObjectKey(component)!);
      if (!objectDefinition) {
        return renderBindingError(component, component.title, "The bound object is missing from this manifest.");
      }

      const view = objectDefinition.views[0];
      const defaultVisibleKeys = view?.visibleFieldKeys ?? objectDefinition.fields.slice(0, 4).map((field) => field.key);
      const visibleKeys = visibleColumnsByObject[objectDefinition.key] ?? defaultVisibleKeys;
      const rawRecords = resolvedRecordsByObject[objectDefinition.key] ?? [];
      const isLoading = loadingRecordObjectKeys.includes(objectDefinition.key);
      const query = recordQueryByObject[objectDefinition.key]?.trim().toLowerCase() ?? "";
      const activeSort = sortByObject[objectDefinition.key];
      const records = rawRecords
        .filter((record) => !pendingDeleteRecordIds[`${objectDefinition.key}:${record.id}`])
        .filter((record) =>
          !query
            ? true
            : visibleKeys.some((fieldKey) => renderFieldValue(record.data[fieldKey]).toLowerCase().includes(query)),
        )
        .sort((left, right) => {
          if (!activeSort) {
            return 0;
          }

          const leftValue = renderFieldValue(left.data[activeSort.fieldKey]).toLowerCase();
          const rightValue = renderFieldValue(right.data[activeSort.fieldKey]).toLowerCase();
          return activeSort.direction === "asc" ? leftValue.localeCompare(rightValue) : rightValue.localeCompare(leftValue);
        });

      return (
        <section className={styles.runtimeCard} key={component.id} style={componentStyle(component)}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>{component.kind === "related_records" ? "Related records" : objectDefinition.pluralLabel}</p>
              <h2>{component.title}</h2>
            </div>
            <span className={styles.badge}>{records.length} records</span>
          </div>
          <div className={styles.tableToolbar}>
            <input
              className={`${styles.input} ${styles.tableSearchBar}`}
              onChange={(event) => setRecordQueryByObject((current) => ({ ...current, [objectDefinition.key]: event.target.value }))}
              placeholder="Search visible fields"
              value={recordQueryByObject[objectDefinition.key] ?? ""}
            />
            <div className={styles.columnPickerWrap} data-column-picker-root="">
              <button
                className={styles.columnPickerTrigger}
                onClick={() => setOpenColumnPickerObjectKey((current) => (current === objectDefinition.key ? null : objectDefinition.key))}
                type="button"
              >
                Columns
              </button>
              {openColumnPickerObjectKey === objectDefinition.key ? (
                <div className={styles.columnPickerDropdown}>
                  {objectDefinition.fields.map((field) => {
                    const isVisible = visibleKeys.includes(field.key);
                    return (
                      <label className={styles.checkboxField} key={field.id}>
                        <input
                          checked={isVisible}
                          onChange={(event) =>
                            setVisibleColumnsByObject((current) => {
                              const next = current[objectDefinition.key] ?? defaultVisibleKeys;
                              return {
                                ...current,
                                [objectDefinition.key]: event.target.checked
                                  ? Array.from(new Set([...next, field.key]))
                                  : next.filter((candidate) => candidate !== field.key),
                              };
                            })
                          }
                          type="checkbox"
                        />
                        <span>{field.label}</span>
                      </label>
                    );
                  })}
                </div>
              ) : null}
            </div>
          </div>
          <div className={styles.tableWrap}>
            <table className={styles.table}>
              <thead>
                <tr>
                  {visibleKeys.map((fieldKey) => (
                    <th key={fieldKey}>
                      <button
                        className={styles.sortableHeader}
                        onClick={() =>
                          setSortByObject((current) => {
                            const currentSort = current[objectDefinition.key];
                            if (!currentSort || currentSort.fieldKey !== fieldKey) {
                              return { ...current, [objectDefinition.key]: { fieldKey, direction: "asc" } };
                            }

                            return {
                              ...current,
                              [objectDefinition.key]: { fieldKey, direction: currentSort.direction === "asc" ? "desc" : "asc" },
                            };
                          })
                        }
                        type="button"
                      >
                        {objectDefinition.fields.find((field) => field.key === fieldKey)?.label ?? fieldKey}
                        <span className={styles.sortIndicator}>
                          {activeSort?.fieldKey === fieldKey ? (activeSort.direction === "asc" ? "↑" : "↓") : "↕"}
                        </span>
                      </button>
                    </th>
                  ))}
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {isLoading ? (
                  Array.from({ length: 4 }, (_, index) => (
                    <tr key={`runtime-record-skeleton-${index}`}>
                      <td colSpan={visibleKeys.length + 1}>
                        <div className={styles.skeletonCard}>
                          <div className={`${styles.skeleton} ${styles.skeletonWide}`} />
                          <div className={`${styles.skeleton} ${styles.skeletonNarrow}`} />
                        </div>
                      </td>
                    </tr>
                  ))
                ) : records.length > 0 ? (
                  records.map((record) => (
                    <tr key={record.id}>
                      {visibleKeys.map((fieldKey) => (
                        <td key={`${record.id}-${fieldKey}`}>{renderFieldValue(record.data[fieldKey])}</td>
                      ))}
                      <td>
                        <button className={styles.ghostButton} onClick={() => void handleDeleteRecord(objectDefinition.key, record.id)} type="button">
                          Delete
                        </button>
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={visibleKeys.length + 1}>
                      {renderEmptyState(
                        mode === "published" ? "No records yet." : "Preview records only.",
                        mode === "published"
                          ? "Create one from the form beside this table."
                          : "Draft preview uses sample records until a published runtime is active.",
                      )}
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </section>
      );
    }

    if (component.kind === "record_form" && resolveComponentObjectKey(component)) {
      const objectDefinition = findObjectDefinition(manifest, resolveComponentObjectKey(component)!);
      if (!objectDefinition) {
        return renderBindingError(component, component.title, "The bound object is missing from this manifest.");
      }

      const draft = draftsByObject[objectDefinition.key] ?? {};

      return (
        <section className={styles.runtimeCard} key={component.id} style={componentStyle(component)}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>{mode === "published" ? "Runtime Form" : "Draft Form Preview"}</p>
              <h2>{component.title}</h2>
            </div>
          </div>
          <form
            className={styles.formStack}
            onSubmit={(event) => {
              event.preventDefault();
              startTransition(() => {
                void handleCreateRecord(objectDefinition);
              });
            }}
          >
            {objectDefinition.fields
              .filter((field) => field.type !== "computed")
              .map((field) => (
                <label className={styles.formField} key={field.id}>
                  <span>{field.label}</span>
                  {field.type === "long_text" ? (
                    <textarea
                      className={styles.textarea}
                      onChange={(event) => updateDraft(objectDefinition.key, field.key, event.target.value)}
                      placeholder={field.placeholder}
                      value={String(draft[field.key] ?? "")}
                    />
                  ) : field.type === "select" ? (
                    <select
                      className={styles.select}
                      onChange={(event) => updateDraft(objectDefinition.key, field.key, event.target.value)}
                      value={String(draft[field.key] ?? field.defaultValue ?? "")}
                    >
                      <option value="">Choose an option</option>
                      {field.options?.map((option) => (
                        <option key={option} value={option}>
                          {option}
                        </option>
                      ))}
                    </select>
                  ) : field.type === "boolean" ? (
                    <input
                      checked={Boolean(draft[field.key] ?? field.defaultValue ?? false)}
                      className={styles.checkbox}
                      onChange={(event) => updateDraft(objectDefinition.key, field.key, event.target.checked)}
                      type="checkbox"
                    />
                  ) : (
                    <input
                      className={styles.input}
                      onChange={(event) => updateDraft(objectDefinition.key, field.key, event.target.value)}
                      placeholder={field.placeholder}
                      type={fieldInputType(field)}
                      value={String(draft[field.key] ?? field.defaultValue ?? "")}
                    />
                  )}
                </label>
              ))}
            <button className={styles.primaryButton} disabled={isPending} type="submit">
              {mode === "published" ? (isPending ? "Saving..." : `Create ${objectDefinition.label}`) : `Preview ${objectDefinition.label}`}
            </button>
          </form>
        </section>
      );
    }

    if (component.kind === "workflow_launcher") {
      const workflows = manifest.workflows.filter((workflow) =>
        resolveComponentWorkflowKey(component)
          ? workflow.key === resolveComponentWorkflowKey(component) || workflow.id === resolveComponentWorkflowKey(component)
          : workflow.objectKey === resolveComponentObjectKey(component) || !resolveComponentObjectKey(component),
      );
      return (
        <section className={styles.runtimeCard} key={component.id} style={componentStyle(component)}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>Workflow Tools</p>
              <h2>{component.title}</h2>
            </div>
          </div>
          <div className={styles.listStack}>
            {workflows.map((workflow) => (
              <article className={`${styles.runCard} ${styles.staggerChild}`} key={workflow.id}>
                <div className={styles.runCardHeader}>
                  <div>
                    <strong>{workflow.name}</strong>
                    <p>{workflow.description ?? "Published workflow"}</p>
                  </div>
                  <button className={styles.secondaryButton} disabled={mode !== "published"} onClick={() => void handleQueueWorkflow(workflow.id)} type="button">
                    {mode === "published" ? "Run" : "Preview"}
                  </button>
                </div>
                <div className={styles.inlineList}>
                  {workflow.nodes.map((node) => (
                    <span className={styles.inlineTag} key={node.id}>
                      {node.type}
                    </span>
                  ))}
                </div>
                {loadingWorkflowKeys.includes(workflow.id) ? (
                  <div className={styles.skeletonRow}>
                    <div className={`${styles.skeleton} ${styles.skeletonWide}`} />
                    <div className={`${styles.skeleton} ${styles.skeletonNarrow}`} />
                  </div>
                ) : (
                  (workflowRunsByWorkflow[workflow.id] ?? []).slice(0, 3).map((run) => (
                    <div className={styles.logRow} key={run.id}>
                      <strong>{run.status}</strong>
                      <span>{formatPlatformDateTime(run.createdAt)}</span>
                    </div>
                  ))
                )}
              </article>
            ))}
          </div>
          <p className={styles.helperCopy}>This launcher reads only the active published manifest and shows recent operator-visible runs.</p>
        </section>
      );
    }

    if (component.kind === "agent_summary" || component.kind === "agent_panel") {
      const agent =
        resolveComponentAgent(component) ??
        manifest.agents.find((candidate) =>
          resolveComponentObjectKey(component) ? candidate.objectKeys.includes(resolveComponentObjectKey(component)!) : true,
        );

      if (!agent) {
        return renderBindingError(component, component.title, "The selected agent is missing or outside the current page scope.");
      }

      const provider = manifest.modelProviders.find((candidate) => candidate.id === agent.modelProviderId || candidate.key === agent.modelProviderId);
      return (
        <section className={styles.runtimeCard} key={component.id} style={componentStyle(component)}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>Agent surface</p>
              <h2>{component.title}</h2>
            </div>
            <span className={styles.badge}>{agent.scope}</span>
          </div>
          <div className={styles.scopeGrid}>
            <article className={styles.scopeCard}>
              <p className={styles.cardEyebrow}>Model</p>
              <strong>{provider?.name ?? "Unknown provider"}</strong>
              <p>{provider?.model ?? "Unbound model"}</p>
              <div className={styles.inlineList}>
                <span className={styles.inlineTag}>{agent.zeroRetentionRequired ? "zero retention" : "standard retention"}</span>
                <span className={styles.inlineTag}>{agent.objectKeys.length} objects in scope</span>
              </div>
            </article>
            <article className={styles.scopeCard}>
              <p className={styles.cardEyebrow}>{component.kind === "agent_panel" ? "Prompt posture" : "Summary"}</p>
              <p>{agent.description ?? "Tenant-scoped agent ready for governed invocation."}</p>
              <div className={styles.sidebarPanel}>
                {component.kind === "agent_panel"
                  ? agent.prompt
                  : `Allowed tools: ${agent.allowedToolIds.length || 0}. Masked data policy remains enforced at the gateway.`}
              </div>
            </article>
          </div>
        </section>
      );
    }

    if (component.kind === "activity_feed") {
      const items = Object.values(workflowRunsByWorkflow)
        .flat()
        .slice(0, 6)
        .map((run) => ({
          id: run.id,
          title: `${run.workflowKey} ${run.status.toLowerCase()}`,
          timestamp: run.createdAt,
        }));
      const previewItems =
        items.length > 0
          ? items
          : [
              { id: "preview-1", title: "Draft preview opened", timestamp: new Date().toISOString() },
              { id: "preview-2", title: "Page layout awaiting publish", timestamp: new Date().toISOString() },
            ];

      return (
        <section className={styles.runtimeCard} key={component.id} style={componentStyle(component)}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>Operations</p>
              <h2>{component.title}</h2>
            </div>
          </div>
          <div className={styles.listStack}>
            {previewItems.map((item) => (
              <article className={styles.auditRow} key={item.id}>
                <div>
                  <strong>{item.title}</strong>
                  <p>{mode === "published" ? "Runtime activity" : "Draft preview"}</p>
                </div>
                <span>{formatPlatformDateTime(item.timestamp)}</span>
              </article>
            ))}
          </div>
        </section>
      );
    }

    return null;
  }

  function renderSidebarSection(sectionKey: string, label: string, content: ReactNode) {
    const isCollapsed = collapsedRuntimeSections[sectionKey];
    return (
      <div className={`${styles.sidebarSection} ${styles.sidebarSectionCollapsible}`}>
        <button
          aria-expanded={!isCollapsed}
          className={styles.sidebarSectionToggle}
          onClick={() => setCollapsedRuntimeSections((current) => ({ ...current, [sectionKey]: !current[sectionKey] }))}
          type="button"
        >
          <span>{label}</span>
          <svg className={styles.navGroupChevron} fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" viewBox="0 0 10 10">
            <path d="M3 1l4 4-4 4" />
          </svg>
        </button>
        {!isCollapsed ? content : null}
      </div>
    );
  }

  return (
    <div className={styles.runtimeShell} style={getThemeCssVariables(manifest)}>
      <aside className={styles.runtimeSidebar}>
        <div className={styles.brandBlock}>
          <div className={styles.brandRow}>
            {brandLogo ? (
              <Image
                alt={`${manifest.tenant.name} logo`}
                className={styles.brandImage}
                height={40}
                src={brandLogo.url}
                unoptimized
                width={40}
              />
            ) : (
              <div className={styles.brandMark}>TA</div>
            )}
            <div className={styles.brandCopy}>
              <p className={styles.sidebarTitle}>{manifest.appShell.productName || manifest.tenant.name}</p>
              <p className={styles.sidebarSub}>
                {mode === "published" ? "Published runtime" : mode === "admin-preview" ? "Admin preview" : "Draft preview"}
              </p>
            </div>
          </div>
          <div className={styles.sidebarModePill}>
            {mode === "published"
              ? "Runtime shell · published manifest"
              : mode === "admin-preview"
                ? "Runtime shell · admin preview"
                : "Runtime shell · draft preview"}
          </div>
        </div>
        {manifest.appShell.announcementSlots.some((slot) => slot.active) ? (
          renderSidebarSection(
            "announcements",
            "Announcements",
            <div className={styles.listStack}>
              {manifest.appShell.announcementSlots
                .filter((slot) => slot.active)
                .map((slot) => (
                  <div className={styles.sidebarPanel} key={slot.key}>
                    <strong>{slot.label}</strong>
                    <p>{slot.message}</p>
                  </div>
                ))}
            </div>,
          )
        ) : null}
        {renderSidebarSection(
          "navigation",
          "Navigation",
          <nav className={styles.navStack}>
            {(groupedMenus.length > 0
              ? groupedMenus.flatMap((group) => [
                  <div className={styles.sidebarMeta} key={`${group.key}-label`}>
                    <span>{group.label}</span>
                    <strong>{group.items.length}</strong>
                  </div>,
                  ...group.items,
                ])
              : orderedMenus
            ).map((menuOrNode) => {
              if (!("id" in menuOrNode)) {
                return menuOrNode;
              }
              const menu = menuOrNode;
              const page = findPageDefinition(manifest, menu.pageKey);
              if (!page) {
                return null;
              }
              const badgeValue = resolveBadgeValue(menu);

              return (
                <button
                  className={menu.pageKey === activePageKey ? styles.activeNavItem : styles.navItem}
                  key={menu.id}
                  onClick={() => setPageAndSync(menu.pageKey)}
                  type="button"
                >
                  <span className={styles.navIcon}>{glyphLabel(menu.label)}</span>
                  <span className={styles.navCopy}>
                    <span>{menu.label}</span>
                    <small>{menu.description ?? menu.group}</small>
                  </span>
                  {badgeValue ? <span className={styles.inlineTag}>{badgeValue}</span> : null}
                </button>
              );
            })}
          </nav>,
        )}
        {renderSidebarSection(
          "status",
          "Status",
          <div className={styles.sidebarMetaCompactRow}>
            <div className={styles.sidebarMetaCompact}>
              <span>Environment</span>
              <strong>{manifest.environment.name}</strong>
            </div>
            <div className={styles.sidebarMetaCompact}>
              <span>Published</span>
              <strong>{mode === "published" ? publishedLabel : mode === "admin-preview" ? "Admin preview" : "Draft preview"}</strong>
            </div>
          </div>,
        )}
        {renderSidebarSection(
          "manifest",
          "Manifest",
          <div className={styles.sidebarPanel}>
            {mode === "published"
              ? "This runtime only reflects published metadata. Draft changes stay hidden until the next activated version."
              : mode === "admin-preview"
                ? "This admin lens renders draft metadata with diagnostics, while user-facing runtime remains unchanged until publish."
                : "This view renders the current draft manifest for builders only. Data mutations stay disabled until publish."}
          </div>,
        )}
        {renderSidebarSection(
          "actions",
          "Quick actions",
          <div className={styles.sidebarActions}>
            {manifest.appShell.quickActions.map((action) => {
              const targetPageKey = action.pageKey;
              const workflow = action.workflowKey
                ? manifest.workflows.find((candidate) => candidate.key === action.workflowKey || candidate.id === action.workflowKey)
                : null;
              return (
                <button
                  className={action.tone === "accent" ? styles.primaryButton : styles.secondaryButton}
                  key={action.key}
                  onClick={() => {
                    if (targetPageKey) {
                      setPageAndSync(targetPageKey);
                    } else if (workflow) {
                      void handleQueueWorkflow(workflow.id);
                    }
                  }}
                  type="button"
                >
                  {action.label}
                </button>
              );
            })}
          </div>,
        )}
        <div className={styles.sidebarSection}>
          <Link className={styles.secondaryLink} href={`/platform/${tenantSlug}`}>
            Open builder
          </Link>
        </div>
        <div className={styles.sidebarFooter}>
          {mode === "published" ? "Internal runtime preview." : mode === "admin-preview" ? "Internal admin preview." : "Internal draft preview."}
          <br />
          Route: /platform/{mode === "published" ? "runtime" : mode === "admin-preview" ? "admin-preview" : "preview"}/{tenantSlug}
        </div>
      </aside>

      <main className={styles.runtimeMain}>
        {manifest.appShell.navigationMode === "topbar" ? (
          <div className={styles.runtimeLensBar}>
            <strong>{manifest.appShell.productName}</strong>
            <div className={styles.inlineList}>
              {orderedMenus.map((menu) => (
                <button className={menu.pageKey === activePageKey ? styles.secondaryButton : styles.ghostButton} key={`topbar-${menu.id}`} onClick={() => setPageAndSync(menu.pageKey)} type="button">
                  {menu.label}
                </button>
              ))}
            </div>
          </div>
        ) : null}
        {mode === "admin-preview" ? (
          <div className={styles.runtimeLensBar}>
            <strong>Admin preview</strong>
            <span>Draft rendering with diagnostics enabled.</span>
          </div>
        ) : null}
        {viewAs?.active ? (
          <div className={styles.runtimeLensBar}>
            <strong>Viewing as {viewAs.personaLabel}</strong>
            <span>{viewAs.role}</span>
            <button className={styles.secondaryButton} onClick={() => void clearViewAs()} type="button">
              Exit view-as
            </button>
          </div>
        ) : null}
        <header className={[styles.runtimeHeader, scrolledPast ? styles.headerCompact : ""].join(" ")}>
          <div className={styles.headerLead}>
            <p className={styles.eyebrow}>
              {mode === "published" ? manifest.environment.name : mode === "admin-preview" ? "Admin preview" : "Draft preview"}
            </p>
            <h1>{activePage?.title ?? "Published runtime"}</h1>
            <p className={styles.headerCopy}>
              {activePage?.description ??
                (mode === "published"
                  ? "This surface is rendered from the current published manifest with no hand-authored page code."
                  : mode === "admin-preview"
                    ? "This surface renders the current draft manifest with admin diagnostics and a view-as lens before publish."
                    : "This surface renders the current draft manifest so builders can validate layout, bindings, and responsiveness before publish.")}
            </p>
          </div>
          <div className={styles.headerRail}>
            <div className={styles.headerStats}>
              <article className={styles.metricCard}>
                <span>Objects</span>
                <strong>{manifest.objects.length}</strong>
              </article>
              <article className={styles.metricCard}>
                <span>Workflows</span>
                <strong>{manifest.workflows.length}</strong>
              </article>
              <article className={styles.metricCard}>
                <span>Agents</span>
                <strong>{manifest.agents.length}</strong>
              </article>
            </div>
            <div className={styles.headerNote}>
              <span>{mode === "published" ? "Published manifest" : mode === "admin-preview" ? "Admin lens" : "Draft manifest"}</span>
              <strong>{manifest.tenant.slug}</strong>
              <p>{mode === "published" ? publishedLabel : manifest.metadata.draftUpdatedAt}</p>
            </div>
          </div>
        </header>
        <div className={styles.runtimePageEnter} key={activePageKey}>
        {activeLayout ? (
          activeLayout.sections.map((section) => (
            <section className={styles.runtimeSection} key={section.id}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>{activePage?.title}</p>
                  <h2>{section.title}</h2>
                </div>
              </div>
              <div className={styles.runtimeGrid}>
                {section.components.map((component) => renderComponent(component))}
              </div>
            </section>
          ))
        ) : (
          renderEmptyState("This page has no published layout.", "Add sections in the platform studio, then publish to make them available here.")
        )}
        </div>
      </main>
      {renderToastStack()}
    </div>
  );
}
