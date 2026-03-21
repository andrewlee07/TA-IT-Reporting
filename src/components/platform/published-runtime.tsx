"use client";

import Link from "next/link";
import { useEffect, useMemo, useState, useTransition } from "react";

import { getSampleFieldValue } from "@/lib/platform/designer";
import { formatPlatformDateTime } from "@/lib/platform/format";
import { findLayoutForPage, findObjectDefinition, findPageDefinition } from "@/lib/platform/manifest";
import type {
  AgentDefinition,
  FieldDefinition,
  LayoutComponentDefinition,
  ObjectDefinition,
  PlatformManifest,
  PlatformRecord,
  PlatformWorkflowRunRecord,
} from "@/lib/platform/types";

import styles from "./platform-shell.module.css";

interface PublishedRuntimeProps {
  manifest: PlatformManifest;
  requestedRoute?: string;
  tenantSlug: string;
  mode?: "published" | "draft-preview";
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

export function PublishedRuntime({ manifest, requestedRoute, tenantSlug, mode = "published" }: PublishedRuntimeProps) {
  const orderedMenus = [...manifest.menus].sort((left, right) => left.order - right.order);
  const firstRoute = manifest.pages.find((page) => page.isHome)?.key ?? orderedMenus[0]?.pageKey;
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
  const [message, setMessage] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isPending, startTransition] = useTransition();

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

  useEffect(() => {
    if (mode !== "published") {
      return;
    }

    let cancelled = false;

    async function loadRecords(): Promise<void> {
      try {
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
          setError(caughtError instanceof Error ? caughtError.message : "Failed to load runtime records.");
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
          setError(caughtError instanceof Error ? caughtError.message : "Failed to load workflow runs.");
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
    const payload = await fetchJson<{ records: PlatformRecord[] }>(`/api/platform/tenants/${tenantSlug}/records/${objectKey}`);
    setRecordsByObject((current) => ({
      ...current,
      [objectKey]: payload.records,
    }));
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

    try {
      setError(null);
      setMessage(null);
      await fetchJson<{ ok: true }>(`/api/platform/tenants/${tenantSlug}/records/${objectKey}/${recordId}`, {
        method: "DELETE",
      });
      await refreshRecords(objectKey);
      setMessage("Record deleted.");
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Failed to delete record.");
    }
  }

  async function refreshWorkflowRuns(workflowId: string): Promise<void> {
    const payload = await fetchJson<{ runs: PlatformWorkflowRunRecord[] }>(`/api/platform/tenants/${tenantSlug}/workflows/${workflowId}/runs`);
    setWorkflowRunsByWorkflow((current) => ({
      ...current,
      [workflowId]: payload.runs,
    }));
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
              <article className={styles.metricCard} key={metric.label}>
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
      const visibleKeys = view?.visibleFieldKeys ?? objectDefinition.fields.slice(0, 4).map((field) => field.key);
      const records = resolvedRecordsByObject[objectDefinition.key] ?? [];

      return (
        <section className={styles.runtimeCard} key={component.id} style={componentStyle(component)}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>{component.kind === "related_records" ? "Related records" : objectDefinition.pluralLabel}</p>
              <h2>{component.title}</h2>
            </div>
            <span className={styles.badge}>{records.length} records</span>
          </div>
          <div className={styles.tableWrap}>
            <table className={styles.table}>
              <thead>
                <tr>
                  {visibleKeys.map((fieldKey) => (
                    <th key={fieldKey}>{objectDefinition.fields.find((field) => field.key === fieldKey)?.label ?? fieldKey}</th>
                  ))}
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {records.length > 0 ? (
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
                      <div className={styles.emptyState}>
                        {mode === "published"
                          ? "No records yet. Create one from the form beside this table."
                          : "Draft preview uses sample records until a published runtime is active."}
                      </div>
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
              <article className={styles.runCard} key={workflow.id}>
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
                {(workflowRunsByWorkflow[workflow.id] ?? []).slice(0, 3).map((run) => (
                  <div className={styles.logRow} key={run.id}>
                    <strong>{run.status}</strong>
                    <span>{formatPlatformDateTime(run.createdAt)}</span>
                  </div>
                ))}
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

  return (
    <div className={styles.runtimeShell}>
      <aside className={styles.runtimeSidebar}>
        <div className={styles.brandBlock}>
          <div className={styles.brandRow}>
            <div className={styles.brandMark}>TA</div>
            <div className={styles.brandCopy}>
              <p className={styles.sidebarTitle}>{manifest.tenant.name}</p>
              <p className={styles.sidebarSub}>{mode === "published" ? "Published runtime" : "Draft preview"}</p>
            </div>
          </div>
          <div className={styles.sidebarModePill}>
            {mode === "published" ? "Runtime shell · published manifest" : "Runtime shell · draft preview"}
          </div>
        </div>
        <div className={styles.sidebarSection}>
          <p className={styles.sidebarLabel}>Navigation</p>
          <nav className={styles.navStack}>
            {manifest.menus
              .slice()
              .sort((left, right) => left.order - right.order)
              .map((menu) => {
                const page = findPageDefinition(manifest, menu.pageKey);
                if (!page) {
                  return null;
                }

                return (
                  <button
                    className={menu.pageKey === activePageKey ? styles.activeNavItem : styles.navItem}
                    key={menu.id}
                    onClick={() => setActivePageKey(menu.pageKey)}
                    type="button"
                  >
                    <span className={styles.navIcon}>{glyphLabel(menu.label)}</span>
                    <span className={styles.navCopy}>
                      <span>{menu.label}</span>
                      <small>{menu.group}</small>
                    </span>
                  </button>
                );
              })}
          </nav>
        </div>
        <div className={styles.sidebarSection}>
          <p className={styles.sidebarLabel}>Status</p>
          <div className={styles.sidebarMeta}>
            <span>Environment</span>
            <strong>{manifest.environment.name}</strong>
          </div>
          <div className={styles.sidebarMeta}>
            <span>Published</span>
            <strong>{mode === "published" ? publishedLabel : "Draft preview"}</strong>
          </div>
        </div>
        <div className={styles.sidebarSection}>
          <p className={styles.sidebarLabel}>Manifest</p>
          <div className={styles.sidebarPanel}>
            {mode === "published"
              ? "This runtime only reflects published metadata. Draft changes stay hidden until the next activated version."
              : "This view renders the current draft manifest for builders only. Data mutations stay disabled until publish."}
          </div>
        </div>
        <div className={styles.sidebarSection}>
          <Link className={styles.secondaryLink} href={`/platform/${tenantSlug}`}>
            Open builder
          </Link>
        </div>
        <div className={styles.sidebarFooter}>
          {mode === "published" ? "Internal runtime preview." : "Internal draft preview."}
          <br />
          Route: /platform/{mode === "published" ? "runtime" : "preview"}/{tenantSlug}
        </div>
      </aside>

      <main className={styles.runtimeMain}>
        {message ? <div className={styles.successBanner}>{message}</div> : null}
        {error ? <div className={styles.errorBanner}>{error}</div> : null}
        <header className={styles.runtimeHeader}>
          <div className={styles.headerLead}>
            <p className={styles.eyebrow}>{mode === "published" ? manifest.environment.name : "Draft preview"}</p>
            <h1>{activePage?.title ?? "Published runtime"}</h1>
            <p className={styles.headerCopy}>
              {activePage?.description ??
                (mode === "published"
                  ? "This surface is rendered from the current published manifest with no hand-authored page code."
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
              <span>{mode === "published" ? "Published manifest" : "Draft manifest"}</span>
              <strong>{manifest.tenant.slug}</strong>
              <p>{mode === "published" ? publishedLabel : manifest.metadata.draftUpdatedAt}</p>
            </div>
          </div>
        </header>
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
          <div className={styles.emptyState}>This page has no published layout.</div>
        )}
      </main>
    </div>
  );
}
