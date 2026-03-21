"use client";

import Link from "next/link";
import { useCallback, useEffect, useEffectEvent, useState, useTransition } from "react";

import {
  createComponentFromPreset,
  createDefaultPlacement,
  createPageTemplateInstance,
  createSectionFromTemplate,
  normalizeLayoutComponentDefinition,
  normalizeLayoutSectionDefinition,
} from "@/lib/platform/designer";
import { formatPlatformDateTime } from "@/lib/platform/format";
import type {
  AgentDefinition,
  LayoutComponentDefinition,
  LayoutDefinition,
  LayoutSectionDefinition,
  MenuItemDefinition,
  ModelProviderDefinition,
  ObjectDefinition,
  PageDefinition,
  PlatformAgentPreview,
  PlatformBootstrap,
  PlatformPublishPreview,
  PlatformRole,
  PlatformWorkflowRunRecord,
  SecurityPolicyDefinition,
  WorkflowDefinition,
  WorkflowEdgeDefinition,
  WorkflowNodeDefinition,
  WorkflowNodeType,
} from "@/lib/platform/types";

import styles from "./platform-shell.module.css";

type WorkspaceKey = "data-model" | "pages" | "navigation" | "workflows" | "agents" | "models" | "security" | "audit";

const WORKSPACES: Array<{ key: WorkspaceKey; label: string; note: string; code: string }> = [
  { key: "data-model", label: "Data Model", note: "Objects, fields, rules, formulas", code: "DM" },
  { key: "pages", label: "Pages", note: "Page definitions, layout sections, runtime components", code: "PG" },
  { key: "navigation", label: "Navigation", note: "Menus, ordering, route exposure", code: "NV" },
  { key: "workflows", label: "Workflows", note: "Visual graph metadata and execution scaffolding", code: "WF" },
  { key: "agents", label: "Agents", note: "Prompt assets, scope, model assignment", code: "AG" },
  { key: "models", label: "Models", note: "Provider registry and zero-retention controls", code: "ML" },
  { key: "security", label: "Security", note: "Masking defaults and protected-model policy", code: "SC" },
  { key: "audit", label: "Audit", note: "Publish history and admin activity", code: "AU" },
];

const WORKSPACE_TABS: Record<WorkspaceKey, string[]> = {
  "data-model": ["Objects", "Fields", "Validation"],
  pages: ["Designer", "Pages", "Templates"],
  navigation: ["Menu Items", "Routes"],
  workflows: ["Definitions", "Runs"],
  agents: ["Definitions", "Prompts", "Scope"],
  models: ["Providers", "Configuration"],
  security: ["Policies", "Access", "Masking Rules"],
  audit: ["History", "Activity"],
};

function createClientId(prefix: string): string {
  return `${prefix}-${Math.random().toString(36).slice(2, 10)}`;
}

function fetchJson<T>(input: RequestInfo, init?: RequestInit): Promise<T> {
  return fetch(input, init).then(async (response) => {
    const payload = (await response.json()) as T & { error?: string };

    if (!response.ok) {
      throw new Error(payload.error ?? "Request failed.");
    }

    return payload;
  });
}

function createBlankLayout(pageKey: string, title: string): LayoutDefinition {
  return {
    id: createClientId("layout"),
    key: pageKey,
    name: title,
    pageKey,
    mobileColumns: 1,
    tabletColumns: 2,
    desktopColumns: 12,
    sections: [
      {
        id: createClientId("section"),
        title,
        kind: "grid",
        columns: 12,
        placement: createDefaultPlacement({
          zone: "main",
          span: 12,
          minHeight: 420,
        }),
        components: [
          {
            id: createClientId("component"),
            kind: "rich_text",
            title,
            width: 12,
            stylePreset: "editorial",
            placement: createDefaultPlacement({
              zone: "main",
              span: 12,
              minHeight: 260,
            }),
            props: {
              body: "Add components to turn this page into a working runtime surface.",
            },
          },
        ],
      },
    ],
  };
}

function createBlankWorkflow(): WorkflowDefinition {
  return {
    id: createClientId("wf"),
    key: "",
    name: "",
    description: "",
    status: "draft",
    triggers: [
      {
        id: createClientId("trigger"),
        type: "manual",
        label: "Manual launch",
        config: {
          notes: "Runs from the operator console or runtime launcher.",
        },
      },
    ],
    nodes: [],
    edges: [],
  };
}

function createWorkflowNode(type: WorkflowNodeType, index: number): WorkflowNodeDefinition {
  const position = {
    x: 48 + (index % 2) * 280,
    y: 88 + Math.floor(index / 2) * 152,
  };

  switch (type) {
    case "condition":
      return {
        id: createClientId("node"),
        type,
        label: "Decision",
        config: {
          expression: "total_value >= 5000",
        },
        position,
      };
    case "crud":
      return {
        id: createClientId("node"),
        type,
        label: "Record action",
        config: {
          operation: "update",
          objectKey: "booking_request",
          targetFieldKey: "status",
          valueExpression: "'Review'",
        },
        position,
      };
    case "formula":
      return {
        id: createClientId("node"),
        type,
        label: "Formula",
        config: {
          expression: "daily_rate * days_requested",
          outputKey: "estimated_value",
        },
        position,
      };
    case "webhook":
      return {
        id: createClientId("node"),
        type,
        label: "Webhook",
        config: {
          method: "POST",
          url: "https://api.example.com/workflows/notify",
          bodyTemplate: '{"recordId":"{{record.id}}"}',
        },
        position,
      };
    case "notification":
      return {
        id: createClientId("node"),
        type,
        label: "Notification",
        config: {
          channel: "email",
          recipient: "ops@teacheractive.com",
          message: "A workflow branch requires attention.",
        },
        position,
      };
    case "wait":
      return {
        id: createClientId("node"),
        type,
        label: "Wait",
        config: {
          durationMinutes: 60,
        },
        position,
      };
    case "approval":
      return {
        id: createClientId("node"),
        type,
        label: "Approval",
        config: {
          approverRole: "BUILDER_ADMIN",
          instructions: "Review before escalating to the regional team.",
        },
        position,
      };
    case "model_call":
      return {
        id: createClientId("node"),
        type,
        label: "Agent step",
        config: {
          agentId: "",
          objectKey: "booking_request",
          promptAsset: "Use the configured agent prompt.",
        },
        position,
      };
  }
}

const WORKFLOW_NODE_LIBRARY: Array<{ type: WorkflowNodeType; title: string; note: string }> = [
  { type: "condition", title: "Condition", note: "Route branches from an expression." },
  { type: "crud", title: "CRUD", note: "Create, update, or delete a record." },
  { type: "formula", title: "Formula", note: "Compute a derived value." },
  { type: "webhook", title: "Webhook", note: "Call an outbound REST endpoint." },
  { type: "notification", title: "Notification", note: "Send an email, task, or Slack style event." },
  { type: "wait", title: "Wait", note: "Pause the workflow for a duration." },
  { type: "approval", title: "Approval", note: "Insert a governed approval checkpoint." },
  { type: "model_call", title: "Model call", note: "Run a tenant agent at a node step." },
];

function createBlankAgent(): AgentDefinition {
  return {
    id: createClientId("agent"),
    key: "",
    name: "",
    description: "",
    scope: "workspace",
    modelProviderId: "",
    prompt: "",
    allowedToolIds: [],
    objectKeys: [],
    zeroRetentionRequired: true,
  };
}

function createWorkflowDraftFromDefinition(workflowDefinition?: WorkflowDefinition) {
  return workflowDefinition ? structuredClone(workflowDefinition) : createBlankWorkflow();
}

function createAgentDraftFromDefinition(agentDefinition?: AgentDefinition) {
  return agentDefinition ? structuredClone(agentDefinition) : createBlankAgent();
}

function createBlankProvider(): ModelProviderDefinition {
  return {
    id: createClientId("provider"),
    key: "",
    name: "",
    provider: "openai",
    model: "",
    endpoint: "",
    apiKeySecretRef: "env:OPENAI_API_KEY",
    supportsZeroRetention: true,
    allowedForSensitiveData: false,
    status: "active",
  };
}

function createBlankObjectDraft() {
  return {
    id: "",
    key: "",
    label: "",
    pluralLabel: "",
    description: "",
    primaryFieldKey: "",
    icon: "layout",
    allowCreate: true,
    allowUpdate: true,
    allowDelete: false,
  };
}

function createObjectDraftFromDefinition(objectDefinition?: ObjectDefinition) {
  if (!objectDefinition) {
    return createBlankObjectDraft();
  }

  return {
    id: objectDefinition.id,
    key: objectDefinition.key,
    label: objectDefinition.label,
    pluralLabel: objectDefinition.pluralLabel,
    description: objectDefinition.description ?? "",
    primaryFieldKey: objectDefinition.primaryFieldKey,
    icon: objectDefinition.icon,
    allowCreate: objectDefinition.allowCreate,
    allowUpdate: objectDefinition.allowUpdate,
    allowDelete: objectDefinition.allowDelete,
  };
}

function createBlankPageDraft() {
  return {
    id: "",
    key: "",
    title: "",
    route: "",
    description: "",
    objectKey: "",
    isHome: false,
    previewNote: "",
  };
}

function createPageDraftFromDefinition(pageDefinition?: PageDefinition) {
  if (!pageDefinition) {
    return createBlankPageDraft();
  }

  return {
    id: pageDefinition.id,
    key: pageDefinition.key,
    title: pageDefinition.title,
    route: pageDefinition.route,
    description: pageDefinition.description ?? "",
    objectKey: pageDefinition.objectKey ?? "",
    isHome: pageDefinition.isHome ?? false,
    previewNote: pageDefinition.previewNote ?? "",
  };
}

function normalizeLayoutDraftForSave(
  layoutDraft: LayoutDefinition,
  pageKey: string,
  layoutKey: string,
  title: string,
  objectKey?: string,
): LayoutDefinition {
  return {
    ...layoutDraft,
    key: layoutKey,
    name: layoutDraft.name || title,
    pageKey,
    sections: layoutDraft.sections.map((section) =>
      normalizeLayoutSectionDefinition({
        ...section,
        title: section.title || title,
        placement: createDefaultPlacement({
          ...section.placement,
          span: section.placement?.span ?? 12,
        }),
        components: section.components.map((component) =>
          normalizeLayoutComponentDefinition({
            ...component,
            objectKey: component.objectKey ?? component.binding?.objectKey ?? objectKey,
            placement: createDefaultPlacement({
              ...component.placement,
              span: component.placement?.span ?? component.width ?? 12,
            }),
          }),
        ),
      }),
    ),
  };
}

export function PlatformStudio({
  initialBootstrap,
  initialWorkspace,
}: {
  initialBootstrap: PlatformBootstrap;
  initialWorkspace: string;
}) {
  type PreviewDevice = "desktop" | "tablet" | "mobile";

  const initialObject = initialBootstrap.draftManifest.objects[0];
  const initialPage = initialBootstrap.draftManifest.pages[0];
  const initialWorkflow = initialBootstrap.draftManifest.workflows[0];
  const initialAgent = initialBootstrap.draftManifest.agents[0];
  const initialLayout =
    initialPage
      ? structuredClone(
          initialBootstrap.draftManifest.layouts.find((candidate) => candidate.key === initialPage.layoutKey) ??
            createBlankLayout(initialPage.key, initialPage.title),
        )
      : null;
  const [bootstrap, setBootstrap] = useState(initialBootstrap);
  const [workspace, setWorkspace] = useState<WorkspaceKey>(
    WORKSPACES.some((entry) => entry.key === initialWorkspace) ? (initialWorkspace as WorkspaceKey) : "data-model",
  );
  const [selectedObjectId, setSelectedObjectId] = useState(initialBootstrap.draftManifest.objects[0]?.id ?? "");
  const [selectedPageId, setSelectedPageId] = useState(initialBootstrap.draftManifest.pages[0]?.id ?? "");
  const [selectedWorkflowId, setSelectedWorkflowId] = useState(initialWorkflow?.id ?? "");
  const [selectedWorkflowNodeId, setSelectedWorkflowNodeId] = useState(initialWorkflow?.nodes[0]?.id ?? "");
  const [selectedAgentId, setSelectedAgentId] = useState(initialAgent?.id ?? "");
  const [objectDraft, setObjectDraft] = useState(createObjectDraftFromDefinition(initialObject));
  const [fieldDraft, setFieldDraft] = useState({
    id: "",
    key: "",
    label: "",
    type: "text",
    sensitivity: "internal",
    required: false,
    unique: false,
    placeholder: "",
  });
  const [pageDraft, setPageDraft] = useState(createPageDraftFromDefinition(initialPage));
  const [layoutDraft, setLayoutDraft] = useState<LayoutDefinition | null>(initialLayout);
  const [selectedSectionId, setSelectedSectionId] = useState(initialLayout?.sections[0]?.id ?? "");
  const [selectedComponentId, setSelectedComponentId] = useState(initialLayout?.sections[0]?.components[0]?.id ?? "");
  const [previewDevice, setPreviewDevice] = useState<PreviewDevice>("desktop");
  const [designerHistory, setDesignerHistory] = useState<Array<{ pageDraft: ReturnType<typeof createBlankPageDraft>; layoutDraft: LayoutDefinition | null }>>([]);
  const [designerDirty, setDesignerDirty] = useState(false);
  const [autosaveStatus, setAutosaveStatus] = useState<"idle" | "saving" | "saved" | "error">("idle");
  const [menuDraft, setMenuDraft] = useState({
    id: "",
    key: "",
    label: "",
    icon: "dot",
    pageKey: "",
    group: "Workspace",
    order: 0,
  });
  const [workflowDraft, setWorkflowDraft] = useState<WorkflowDefinition>(createWorkflowDraftFromDefinition(initialWorkflow));
  const [workflowRuns, setWorkflowRuns] = useState<PlatformWorkflowRunRecord[]>([]);
  const [workflowEdgeDraft, setWorkflowEdgeDraft] = useState({ sourceId: "", targetId: "", label: "" });
  const [agentDraft, setAgentDraft] = useState<AgentDefinition>(createAgentDraftFromDefinition(initialAgent));
  const [agentPreview, setAgentPreview] = useState<PlatformAgentPreview | null>(null);
  const [publishPreview, setPublishPreview] = useState<PlatformPublishPreview | null>(null);
  const [providerDraft, setProviderDraft] = useState<ModelProviderDefinition>(createBlankProvider());
  const [inviteDraft, setInviteDraft] = useState({
    email: "",
    role: "BUILDER_ADMIN" as PlatformRole,
    expiresInDays: 7,
  });
  const [securityDraft, setSecurityDraft] = useState<SecurityPolicyDefinition>(initialBootstrap.draftManifest.securityPolicy);
  const [draggedComponentId, setDraggedComponentId] = useState<string | null>(null);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [hasLoadedSidebarPreference, setHasLoadedSidebarPreference] = useState(false);
  const [activeTab, setActiveTab] = useState(0);
  const [message, setMessage] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isPending, startTransition] = useTransition();

  const manifest = bootstrap.draftManifest;
  const actor = bootstrap.actor;
  const selectedObject = manifest.objects.find((objectDefinition) => objectDefinition.id === selectedObjectId) ?? manifest.objects[0];
  const selectedPage = manifest.pages.find((pageDefinition) => pageDefinition.id === selectedPageId) ?? manifest.pages[0];
  const selectedWorkflow = manifest.workflows.find((workflowDefinition) => workflowDefinition.id === selectedWorkflowId) ?? manifest.workflows[0];
  const selectedWorkflowNode =
    workflowDraft.nodes.find((node) => node.id === selectedWorkflowNodeId) ??
    selectedWorkflow?.nodes[0] ??
    workflowDraft.nodes[0];
  const selectedAgent = manifest.agents.find((agentDefinition) => agentDefinition.id === selectedAgentId) ?? manifest.agents[0];
  const selectedSection = layoutDraft?.sections.find((section) => section.id === selectedSectionId) ?? layoutDraft?.sections[0] ?? null;
  const selectedComponent =
    selectedSection?.components.find((component) => component.id === selectedComponentId) ??
    selectedSection?.components[0] ??
    null;
  const sortedMenus = [...manifest.menus].sort((left, right) => left.order - right.order);
  const activeWorkspace = WORKSPACES.find((entry) => entry.key === workspace) ?? WORKSPACES[0];
  const publishedLabel = bootstrap.activeVersion ? `v${bootstrap.activeVersion.versionNumber}` : "Draft only";
  const agentActivity = bootstrap.auditEvents.filter(
    (event) => event.resourceType === "agent" && event.resourceId === (selectedAgent?.id ?? agentDraft.id),
  );

  useEffect(() => {
    if (typeof window === "undefined") {
      return;
    }

    const nextUrl = new URL(window.location.href);
    nextUrl.searchParams.set("workspace", workspace);
    window.history.replaceState({}, "", nextUrl);
  }, [workspace]);

  useEffect(() => {
    try {
      setSidebarCollapsed(window.localStorage.getItem("ta-platform-sidebar-collapsed") === "true");
    } catch {
      setSidebarCollapsed(false);
    } finally {
      setHasLoadedSidebarPreference(true);
    }
  }, []);

  useEffect(() => {
    if (!hasLoadedSidebarPreference) {
      return;
    }

    try {
      window.localStorage.setItem("ta-platform-sidebar-collapsed", String(sidebarCollapsed));
    } catch {
      // localStorage access is optional
    }
  }, [hasLoadedSidebarPreference, sidebarCollapsed]);

  const refreshPublishPreview = useCallback(async (): Promise<void> => {
    try {
      const payload = await fetchJson<{ preview: PlatformPublishPreview }>(`/api/platform/tenants/${bootstrap.tenant.slug}/publish/preview`);
      setPublishPreview(payload.preview);
    } catch {
      setPublishPreview(null);
    }
  }, [bootstrap.tenant.slug]);

  const refreshWorkflowRuns = useCallback(async (workflowId: string): Promise<void> => {
    try {
      const payload = await fetchJson<{ runs: PlatformWorkflowRunRecord[] }>(
        `/api/platform/tenants/${bootstrap.tenant.slug}/workflows/${workflowId}/runs`,
      );
      setWorkflowRuns(payload.runs);
    } catch {
      setWorkflowRuns([]);
    }
  }, [bootstrap.tenant.slug]);

  async function refreshBootstrap(): Promise<void> {
    const payload = await fetchJson<PlatformBootstrap>(`/api/platform/tenants/${bootstrap.tenant.slug}/bootstrap`);
    setBootstrap(payload);
    const nextObject =
      payload.draftManifest.objects.find((objectDefinition) => objectDefinition.id === selectedObjectId) ?? payload.draftManifest.objects[0];
    const nextPage =
      payload.draftManifest.pages.find((pageDefinition) => pageDefinition.id === selectedPageId) ?? payload.draftManifest.pages[0];
    const nextWorkflow =
      payload.draftManifest.workflows.find((workflowDefinition) => workflowDefinition.id === selectedWorkflowId) ?? payload.draftManifest.workflows[0];
    const nextAgent =
      payload.draftManifest.agents.find((agentDefinition) => agentDefinition.id === selectedAgentId) ?? payload.draftManifest.agents[0];

    setSelectedObjectId(nextObject?.id ?? "");
    setSelectedPageId(nextPage?.id ?? "");
    setSelectedWorkflowId(nextWorkflow?.id ?? "");
    setSelectedWorkflowNodeId(nextWorkflow?.nodes[0]?.id ?? "");
    setSelectedAgentId(nextAgent?.id ?? "");
    setObjectDraft(createObjectDraftFromDefinition(nextObject));
    setPageDraft(createPageDraftFromDefinition(nextPage));
    setWorkflowDraft(createWorkflowDraftFromDefinition(nextWorkflow));
    setWorkflowEdgeDraft({
      sourceId: nextWorkflow?.nodes[0]?.id ?? "",
      targetId: nextWorkflow?.nodes[1]?.id ?? nextWorkflow?.nodes[0]?.id ?? "",
      label: "",
    });
    setAgentDraft(createAgentDraftFromDefinition(nextAgent));
    const nextLayout =
      nextPage
        ? structuredClone(
            payload.draftManifest.layouts.find((candidate) => candidate.key === nextPage.layoutKey) ??
              createBlankLayout(nextPage.key, nextPage.title),
          )
        : null;
    setLayoutDraft(nextLayout);
    setSelectedSectionId(nextLayout?.sections[0]?.id ?? "");
    setSelectedComponentId(nextLayout?.sections[0]?.components[0]?.id ?? "");
    setSecurityDraft(payload.draftManifest.securityPolicy);
    await refreshPublishPreview();
    if (nextWorkflow) {
      await refreshWorkflowRuns(nextWorkflow.id);
    } else {
      setWorkflowRuns([]);
    }
  }

  async function executeAction(action: () => Promise<void>, successMessage: string): Promise<void> {
    try {
      setError(null);
      setMessage(null);
      await action();
      await refreshBootstrap();
      setMessage(successMessage);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Action failed.");
    }
  }

  function snapshotDesignerState() {
    return {
      pageDraft: structuredClone(pageDraft),
      layoutDraft: layoutDraft ? structuredClone(layoutDraft) : null,
    };
  }

  function pushDesignerHistory(): void {
    setDesignerHistory((current) => [...current.slice(-11), snapshotDesignerState()]);
  }

  function markDesignerDirty(): void {
    setDesignerDirty(true);
    setAutosaveStatus("idle");
  }

  function updatePageDraftState(updater: (current: typeof pageDraft) => typeof pageDraft): void {
    pushDesignerHistory();
    setPageDraft((current) => updater(current));
    markDesignerDirty();
  }

  function updateLayoutDraftState(updater: (current: LayoutDefinition | null) => LayoutDefinition | null): void {
    pushDesignerHistory();
    setLayoutDraft((current) => updater(current));
    markDesignerDirty();
  }

  function applySelectedPage(pageDefinition?: PageDefinition): void {
    const nextPage = pageDefinition ?? manifest.pages[0];
    const nextLayout = nextPage
      ? structuredClone(manifest.layouts.find((candidate) => candidate.key === nextPage.layoutKey) ?? createBlankLayout(nextPage.key, nextPage.title))
      : null;

    setSelectedPageId(nextPage?.id ?? "");
    setPageDraft(createPageDraftFromDefinition(nextPage));
    setLayoutDraft(nextLayout);
    setSelectedSectionId(nextLayout?.sections[0]?.id ?? "");
    setSelectedComponentId(nextLayout?.sections[0]?.components[0]?.id ?? "");
    setDesignerHistory([]);
    setDesignerDirty(false);
    setAutosaveStatus("idle");
  }

  async function persistDesignerDraft(options?: { quiet?: boolean }): Promise<void> {
    if (!pageDraft.title.trim()) {
      return;
    }

    const savedPagePayload = await fetchJson<{ page: PageDefinition }>(`/api/platform/tenants/${bootstrap.tenant.slug}/pages`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
      },
      body: JSON.stringify(pageDraft),
    });

    const savedPage = savedPagePayload.page;
    setSelectedPageId(savedPage.id);

    const nextLayout = normalizeLayoutDraftForSave(
      layoutDraft ?? createBlankLayout(savedPage.key, savedPage.title),
      savedPage.key,
      savedPage.layoutKey,
      savedPage.title,
      pageDraft.objectKey,
    );
    setLayoutDraft(nextLayout);

    await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/layouts`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
      },
      body: JSON.stringify(nextLayout),
    });

    await refreshBootstrap();
    setDesignerDirty(false);
    setAutosaveStatus("saved");
    if (!options?.quiet) {
      setMessage(`Saved page ${savedPage.title}.`);
    }
  }

  const triggerDesignerAutosave = useEffectEvent(() => {
    startTransition(() => {
      setAutosaveStatus("saving");
      void persistDesignerDraft({ quiet: true }).catch((caughtError) => {
        setAutosaveStatus("error");
        setError(caughtError instanceof Error ? caughtError.message : "Autosave failed.");
      });
    });
  });

  function handleUndoDesignerChange(): void {
    const previous = designerHistory.at(-1);
    if (!previous) {
      return;
    }

    setDesignerHistory((current) => current.slice(0, -1));
    setPageDraft(previous.pageDraft);
    setLayoutDraft(previous.layoutDraft);
    setSelectedSectionId(previous.layoutDraft?.sections[0]?.id ?? "");
    setSelectedComponentId(previous.layoutDraft?.sections[0]?.components[0]?.id ?? "");
    markDesignerDirty();
  }

  useEffect(() => {
    void refreshPublishPreview();
  }, [refreshPublishPreview, bootstrap.draftManifest.metadata.draftUpdatedAt, bootstrap.activeVersion?.id]);

  useEffect(() => {
    if (!selectedWorkflow) {
      setWorkflowRuns([]);
      return;
    }

    void refreshWorkflowRuns(selectedWorkflow.id);
  }, [selectedWorkflow, refreshWorkflowRuns, bootstrap.activeVersion?.id]);

  useEffect(() => {
    if (workspace !== "pages" || !designerDirty || !layoutDraft || !pageDraft.title.trim()) {
      return;
    }

    const timer = window.setTimeout(() => {
      triggerDesignerAutosave();
    }, 900);

    return () => window.clearTimeout(timer);
  }, [designerDirty, layoutDraft, pageDraft, workspace]);

  async function handlePublish(): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/publish`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            notes: "Published from the adaptive platform studio.",
          }),
        });
      },
      "Draft published and written to versioned manifest output.",
    );
  }

  async function handleObjectSave(): Promise<void> {
    await executeAction(
      async () => {
        const payload = await fetchJson<{ object: ObjectDefinition }>(`/api/platform/tenants/${bootstrap.tenant.slug}/objects`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify(objectDraft),
        });
        setSelectedObjectId(payload.object.id);
      },
      `Saved object ${objectDraft.label}.`,
    );
  }

  async function handleObjectDelete(objectId: string): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/objects/${objectId}`, {
          method: "DELETE",
        });
      },
      "Object deleted.",
    );
  }

  async function handleFieldSave(): Promise<void> {
    if (!selectedObject) {
      return;
    }

    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/fields`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            objectId: selectedObject.id,
            field: {
              ...fieldDraft,
              validations: fieldDraft.required
                ? [
                    {
                      id: createClientId("val"),
                      type: "required",
                      message: `${fieldDraft.label} is required.`,
                    },
                  ]
                : [],
            },
          }),
        });
        setFieldDraft({
          id: "",
          key: "",
          label: "",
          type: "text",
          sensitivity: "internal",
          required: false,
          unique: false,
          placeholder: "",
        });
      },
      `Saved field ${fieldDraft.label}.`,
    );
  }

  async function handleFieldDelete(fieldId: string): Promise<void> {
    if (!selectedObject) {
      return;
    }

    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/fields/${fieldId}?objectId=${selectedObject.id}`, {
          method: "DELETE",
        });
      },
      "Field deleted.",
    );
  }

  async function handlePageSave(): Promise<void> {
    try {
      setError(null);
      setMessage(null);
      await persistDesignerDraft();
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Failed to save page.");
    }
  }

  async function handleLayoutSave(): Promise<void> {
    await handlePageSave();
  }

  async function handleMenuSave(): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/menus`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify(menuDraft),
        });
        setMenuDraft({
          id: "",
          key: "",
          label: "",
          icon: "dot",
          pageKey: "",
          group: "Workspace",
          order: manifest.menus.length,
        });
      },
      `Saved menu item ${menuDraft.label}.`,
    );
  }

  async function handleWorkflowSave(): Promise<void> {
    await executeAction(
      async () => {
        const payload = await fetchJson<{ workflow: WorkflowDefinition }>(`/api/platform/tenants/${bootstrap.tenant.slug}/workflows`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify(workflowDraft),
        });
        setSelectedWorkflowId(payload.workflow.id);
        setSelectedWorkflowNodeId(payload.workflow.nodes[0]?.id ?? "");
      },
      `Saved workflow ${workflowDraft.name}.`,
    );
  }

  async function handleQueueWorkflowRun(workflowId: string): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/workflows/${workflowId}/runs`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            payload: {
              launchedFrom: "studio",
              requestedAt: new Date().toISOString(),
            },
          }),
        });
        await refreshWorkflowRuns(workflowId);
      },
      "Workflow run queued.",
    );
  }

  async function handleAgentSave(): Promise<void> {
    await executeAction(
      async () => {
        const payload = await fetchJson<{ agent: AgentDefinition }>(`/api/platform/tenants/${bootstrap.tenant.slug}/agents`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify(agentDraft),
        });
        setSelectedAgentId(payload.agent.id);
      },
      `Saved agent ${agentDraft.name}.`,
    );
  }

  async function handlePreviewAgent(agentId: string): Promise<void> {
    const scopedObjectKey = agentDraft.objectKeys[0] ?? manifest.objects[0]?.key;
    if (!scopedObjectKey) {
      setError("Agent preview requires at least one object to be in scope.");
      return;
    }

    try {
      setError(null);
      setMessage(null);
      const payload = await fetchJson<{ preview: PlatformAgentPreview }>(
        `/api/platform/tenants/${bootstrap.tenant.slug}/agents/${agentId}/preview`,
        {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            objectKey: scopedObjectKey,
            sampleSize: 3,
          }),
        },
      );
      setAgentPreview(payload.preview);
      await refreshBootstrap();
      setMessage(`Prepared masked preview for ${agentDraft.name}.`);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Failed to preview agent.");
    }
  }

  async function handleProviderSave(): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/models`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify(providerDraft),
        });
        setProviderDraft(createBlankProvider());
      },
      `Saved provider ${providerDraft.name}.`,
    );
  }

  async function handleSecuritySave(): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/security`, {
          method: "PUT",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify(securityDraft),
        });
      },
      "Security policy updated.",
    );
  }

  async function handleInviteCreate(): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/invites`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify(inviteDraft),
        });
        setInviteDraft({
          email: "",
          role: "BUILDER_ADMIN",
          expiresInDays: 7,
        });
      },
      `Created access invite for ${inviteDraft.email}.`,
    );
  }

  function updateLayoutSection(sectionId: string, updater: (section: LayoutSectionDefinition) => LayoutSectionDefinition): void {
    updateLayoutDraftState((current) => {
      if (!current) {
        return current;
      }

      return {
        ...current,
        sections: current.sections.map((section) =>
          section.id === sectionId ? normalizeLayoutSectionDefinition(updater(section)) : section,
        ),
      };
    });
  }

  function addLayoutSection(templateKey?: string): void {
    const preparedSection = templateKey
      ? createSectionFromTemplate({
          templateKey,
          idFactory: createClientId,
          objectKey: pageDraft.objectKey || undefined,
        })
      : normalizeLayoutSectionDefinition({
          id: createClientId("section"),
          title: `Section ${(layoutDraft?.sections.length ?? 0) + 1}`,
          kind: "grid",
          columns: 12,
          placement: createDefaultPlacement({
            zone: "main",
            span: 12,
            minHeight: 280,
          }),
          components: [],
        });

    updateLayoutDraftState((current) => {
      if (!current) {
        return current;
      }

      return {
        ...current,
        sections: [...current.sections, preparedSection],
      };
    });
    setSelectedSectionId(preparedSection.id);
    setSelectedComponentId(preparedSection.components[0]?.id ?? "");
  }

  function addComponentToSection(sectionId: string, presetKey = "rich_text"): void {
    const component = createComponentFromPreset({
      presetKey,
      idFactory: createClientId,
      objectKey: pageDraft.objectKey || undefined,
    });
    updateLayoutSection(sectionId, (section) => ({
      ...section,
      components: [...section.components, component],
    }));
    setSelectedSectionId(sectionId);
    setSelectedComponentId(component.id);
  }

  function updateComponent(sectionId: string, componentId: string, updater: (component: LayoutComponentDefinition) => LayoutComponentDefinition): void {
    updateLayoutSection(sectionId, (section) => ({
      ...section,
      components: section.components.map((component) =>
        component.id === componentId ? normalizeLayoutComponentDefinition(updater(component)) : component,
      ),
    }));
  }

  function removeSection(sectionId: string): void {
    updateLayoutDraftState((current) => {
      if (!current) {
        return current;
      }

      const nextSections = current.sections.filter((section) => section.id !== sectionId);
      return {
        ...current,
        sections: nextSections,
      };
    });
    setSelectedSectionId("");
    setSelectedComponentId("");
  }

  function duplicateSection(sectionId: string): void {
    updateLayoutDraftState((current) => {
      if (!current) {
        return current;
      }

      const source = current.sections.find((section) => section.id === sectionId);
      if (!source) {
        return current;
      }

      const duplicated = normalizeLayoutSectionDefinition({
        ...structuredClone(source),
        id: createClientId("section"),
        title: `${source.title} Copy`,
        components: source.components.map((component) => ({
          ...structuredClone(component),
          id: createClientId("component"),
        })),
      });

      return {
        ...current,
        sections: [...current.sections, duplicated],
      };
    });
  }

  function removeComponent(sectionId: string, componentId: string): void {
    updateLayoutSection(sectionId, (section) => ({
      ...section,
      components: section.components.filter((component) => component.id !== componentId),
    }));
    setSelectedComponentId("");
  }

  function duplicateComponent(sectionId: string, componentId: string): void {
    updateLayoutSection(sectionId, (section) => {
      const source = section.components.find((component) => component.id === componentId);
      if (!source) {
        return section;
      }

      return {
        ...section,
        components: [
          ...section.components,
          normalizeLayoutComponentDefinition({
            ...structuredClone(source),
            id: createClientId("component"),
            title: `${source.title} Copy`,
          }),
        ],
      };
    });
  }

  function applyPageTemplate(templateKey: string): void {
    const template = bootstrap.designerCatalog.pageTemplates.find((candidate) => candidate.key === templateKey);
    if (!template) {
      return;
    }

    pushDesignerHistory();
    const title = template.page.title;
    const pageKey = `${template.key}_${createClientId("page").slice(-4)}`;
    const route = pageKey.replace(/_/g, "-");
    const instance = createPageTemplateInstance({
      template,
      idFactory: createClientId,
      pageKey,
      title,
      route,
      objectKey: pageDraft.objectKey || template.page.objectKey,
    });

    setSelectedPageId("");
    setPageDraft(createPageDraftFromDefinition(instance.page));
    setLayoutDraft({
      id: createClientId("layout"),
      key: pageKey,
      name: title,
      pageKey,
      mobileColumns: template.layout.mobileColumns,
      tabletColumns: template.layout.tabletColumns,
      desktopColumns: template.layout.desktopColumns,
      sections: instance.layoutSections,
    });
    setSelectedSectionId(instance.layoutSections[0]?.id ?? "");
    setSelectedComponentId(instance.layoutSections[0]?.components[0]?.id ?? "");
    markDesignerDirty();
  }

  function reorderDraggedComponent(sectionId: string, targetComponentId: string): void {
    if (!draggedComponentId) {
      return;
    }

    updateLayoutSection(sectionId, (section) => {
      const draggedComponent = section.components.find((component) => component.id === draggedComponentId);
      const targetIndex = section.components.findIndex((component) => component.id === targetComponentId);
      const sourceIndex = section.components.findIndex((component) => component.id === draggedComponentId);

      if (!draggedComponent || targetIndex === -1 || sourceIndex === -1) {
        return section;
      }

      const nextComponents = [...section.components];
      nextComponents.splice(sourceIndex, 1);
      nextComponents.splice(targetIndex, 0, draggedComponent);
      return {
        ...section,
        components: nextComponents,
      };
    });
    setDraggedComponentId(null);
  }

  function moveMenu(menu: MenuItemDefinition, direction: -1 | 1): void {
    const orderedMenus = [...sortedMenus];
    const currentIndex = orderedMenus.findIndex((candidate) => candidate.id === menu.id);
    const nextIndex = currentIndex + direction;
    if (currentIndex === -1 || nextIndex < 0 || nextIndex >= orderedMenus.length) {
      return;
    }

    const swapped = orderedMenus[nextIndex];
    const nextMenu = { ...menu, order: swapped.order };
    const nextSwapped = { ...swapped, order: menu.order };

    startTransition(() => {
      void executeAction(
        async () => {
          await Promise.all([
            fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/menus`, {
              method: "POST",
              headers: { "content-type": "application/json" },
              body: JSON.stringify(nextMenu),
            }),
            fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/menus`, {
              method: "POST",
              headers: { "content-type": "application/json" },
              body: JSON.stringify(nextSwapped),
            }),
          ]);
        },
        "Navigation order updated.",
      );
    });
  }

  function selectWorkflow(workflow?: WorkflowDefinition): void {
    if (!workflow) {
      setSelectedWorkflowId("");
      setSelectedWorkflowNodeId("");
      setWorkflowDraft(createBlankWorkflow());
      setWorkflowEdgeDraft({ sourceId: "", targetId: "", label: "" });
      return;
    }

    setSelectedWorkflowId(workflow.id);
    setSelectedWorkflowNodeId(workflow.nodes[0]?.id ?? "");
    setWorkflowDraft(createWorkflowDraftFromDefinition(workflow));
    setWorkflowEdgeDraft({
      sourceId: workflow.nodes[0]?.id ?? "",
      targetId: workflow.nodes[1]?.id ?? workflow.nodes[0]?.id ?? "",
      label: "",
    });
  }

  function selectAgent(agent?: AgentDefinition): void {
    if (!agent) {
      setSelectedAgentId("");
      setAgentDraft(createBlankAgent());
      setAgentPreview(null);
      return;
    }

    setSelectedAgentId(agent.id);
    setAgentDraft(createAgentDraftFromDefinition(agent));
    setAgentPreview(null);
  }

  function updateWorkflowNode(nodeId: string, updater: (node: WorkflowNodeDefinition) => WorkflowNodeDefinition): void {
    setWorkflowDraft((current) => ({
      ...current,
      nodes: current.nodes.map((node) => (node.id === nodeId ? updater(node) : node)),
    }));
  }

  function addWorkflowNode(type: WorkflowNodeType): void {
    setWorkflowDraft((current) => {
      const nextNode = createWorkflowNode(type, current.nodes.length);
      setSelectedWorkflowNodeId(nextNode.id);
      setWorkflowEdgeDraft((edgeDraft) => ({
        ...edgeDraft,
        sourceId: edgeDraft.sourceId || nextNode.id,
        targetId: edgeDraft.targetId || nextNode.id,
      }));
      return {
        ...current,
        nodes: [...current.nodes, nextNode],
      };
    });
  }

  function removeWorkflowNode(nodeId: string): void {
    setWorkflowDraft((current) => ({
      ...current,
      nodes: current.nodes.filter((node) => node.id !== nodeId),
      edges: current.edges.filter((edge) => edge.sourceId !== nodeId && edge.targetId !== nodeId),
    }));
    if (selectedWorkflowNodeId === nodeId) {
      setSelectedWorkflowNodeId("");
    }
  }

  function addWorkflowEdge(): void {
    if (!workflowEdgeDraft.sourceId || !workflowEdgeDraft.targetId) {
      return;
    }

    const edge: WorkflowEdgeDefinition = {
      id: createClientId("edge"),
      sourceId: workflowEdgeDraft.sourceId,
      targetId: workflowEdgeDraft.targetId,
      label: workflowEdgeDraft.label || undefined,
    };

    setWorkflowDraft((current) => ({
      ...current,
      edges: [...current.edges, edge],
    }));
    setWorkflowEdgeDraft((current) => ({ ...current, label: "" }));
  }

  function removeWorkflowEdge(edgeId: string): void {
    setWorkflowDraft((current) => ({
      ...current,
      edges: current.edges.filter((edge) => edge.id !== edgeId),
    }));
  }

  function toggleAgentObject(objectKey: string): void {
    setAgentDraft((current) => ({
      ...current,
      objectKeys: current.objectKeys.includes(objectKey)
        ? current.objectKeys.filter((candidate) => candidate !== objectKey)
        : [...current.objectKeys, objectKey],
    }));
  }

  function toggleAgentTool(toolId: string): void {
    setAgentDraft((current) => ({
      ...current,
      allowedToolIds: current.allowedToolIds.includes(toolId)
        ? current.allowedToolIds.filter((candidate) => candidate !== toolId)
        : [...current.allowedToolIds, toolId],
    }));
  }

  function renderTabBar() {
    const tabs = WORKSPACE_TABS[workspace];
    return (
      <div className={styles.tabBar}>
        {tabs.map((tabLabel, index) => (
          <button
            className={index === activeTab ? styles.activeTab : styles.tab}
            key={tabLabel}
            onClick={() => setActiveTab(index)}
            type="button"
          >
            {tabLabel}
          </button>
        ))}
      </div>
    );
  }

  function renderDataModelWorkspace() {
    return (
      <div className={styles.workspaceGrid}>
        <section className={styles.panel}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>Objects</p>
              <h2>Metadata objects</h2>
            </div>
            <button
              className={styles.secondaryButton}
              onClick={() => {
                setSelectedObjectId("");
                setObjectDraft(createBlankObjectDraft());
              }}
              type="button"
            >
              New object
            </button>
          </div>
          <div className={styles.listStack}>
            {manifest.objects.map((objectDefinition) => (
              <button
                className={objectDefinition.id === selectedObject?.id ? styles.activeListItem : styles.listItem}
                key={objectDefinition.id}
                onClick={() => {
                  setSelectedObjectId(objectDefinition.id);
                  setObjectDraft(createObjectDraftFromDefinition(objectDefinition));
                }}
                type="button"
              >
                <span>{objectDefinition.label}</span>
                <small>{objectDefinition.fields.length} fields</small>
              </button>
            ))}
          </div>
        </section>

        <section className={styles.panelWide}>
          {activeTab === 0 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Object editor</p>
                  <h2>{selectedObject ? selectedObject.label : "Create object"}</h2>
                </div>
                {selectedObject ? (
                  <button className={styles.ghostButtonDanger} onClick={() => void handleObjectDelete(selectedObject.id)} type="button">
                    Delete
                  </button>
                ) : null}
              </div>
              <div className={styles.formGrid}>
                <label className={styles.formField}>
                  <span>Label</span>
                  <input className={styles.input} onChange={(event) => setObjectDraft((current) => ({ ...current, label: event.target.value }))} value={objectDraft.label} />
                </label>
                <label className={styles.formField}>
                  <span>Plural label</span>
                  <input className={styles.input} onChange={(event) => setObjectDraft((current) => ({ ...current, pluralLabel: event.target.value }))} value={objectDraft.pluralLabel} />
                </label>
                <label className={styles.formField}>
                  <span>Key</span>
                  <input className={styles.input} onChange={(event) => setObjectDraft((current) => ({ ...current, key: event.target.value }))} value={objectDraft.key} />
                </label>
                <label className={styles.formField}>
                  <span>Primary field key</span>
                  <input
                    className={styles.input}
                    onChange={(event) => setObjectDraft((current) => ({ ...current, primaryFieldKey: event.target.value }))}
                    value={objectDraft.primaryFieldKey}
                  />
                </label>
                <label className={styles.formFieldSpan}>
                  <span>Description</span>
                  <textarea className={styles.textarea} onChange={(event) => setObjectDraft((current) => ({ ...current, description: event.target.value }))} value={objectDraft.description} />
                </label>
              </div>
              <div className={styles.inlineList}>
                <label className={styles.checkboxField}>
                  <input checked={objectDraft.allowCreate} onChange={(event) => setObjectDraft((current) => ({ ...current, allowCreate: event.target.checked }))} type="checkbox" />
                  <span>Allow create</span>
                </label>
                <label className={styles.checkboxField}>
                  <input checked={objectDraft.allowUpdate} onChange={(event) => setObjectDraft((current) => ({ ...current, allowUpdate: event.target.checked }))} type="checkbox" />
                  <span>Allow update</span>
                </label>
                <label className={styles.checkboxField}>
                  <input checked={objectDraft.allowDelete} onChange={(event) => setObjectDraft((current) => ({ ...current, allowDelete: event.target.checked }))} type="checkbox" />
                  <span>Allow delete</span>
                </label>
              </div>
              <div className={styles.actionsRow}>
                <button className={styles.primaryButton} onClick={() => void handleObjectSave()} type="button">
                  Save object
                </button>
              </div>
            </>
          ) : null}

          {activeTab === 1 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Fields</p>
                  <h3>Field model{selectedObject ? ` — ${selectedObject.label}` : ""}</h3>
                </div>
              </div>
              <div className={styles.listStack}>
                {selectedObject?.fields.map((field) => (
                  <div className={styles.fieldCard} key={field.id}>
                    <div>
                      <strong>{field.label}</strong>
                      <p>
                        {field.key} · {field.type} · {field.sensitivity}
                      </p>
                    </div>
                    <button className={styles.ghostButtonDanger} onClick={() => void handleFieldDelete(field.id)} type="button">
                      Delete
                    </button>
                  </div>
                ))}
              </div>
              <div className={styles.formGridTight}>
                <label className={styles.formField}>
                  <span>Field label</span>
                  <input className={styles.input} onChange={(event) => setFieldDraft((current) => ({ ...current, label: event.target.value }))} value={fieldDraft.label} />
                </label>
                <label className={styles.formField}>
                  <span>Key</span>
                  <input className={styles.input} onChange={(event) => setFieldDraft((current) => ({ ...current, key: event.target.value }))} value={fieldDraft.key} />
                </label>
                <label className={styles.formField}>
                  <span>Type</span>
                  <select className={styles.select} onChange={(event) => setFieldDraft((current) => ({ ...current, type: event.target.value }))} value={fieldDraft.type}>
                    <option value="text">Text</option>
                    <option value="long_text">Long text</option>
                    <option value="number">Number</option>
                    <option value="currency">Currency</option>
                    <option value="boolean">Boolean</option>
                    <option value="date">Date</option>
                    <option value="datetime">DateTime</option>
                    <option value="select">Select</option>
                    <option value="computed">Computed</option>
                  </select>
                </label>
                <label className={styles.formField}>
                  <span>Sensitivity</span>
                  <select className={styles.select} onChange={(event) => setFieldDraft((current) => ({ ...current, sensitivity: event.target.value }))} value={fieldDraft.sensitivity}>
                    <option value="public">Public</option>
                    <option value="internal">Internal</option>
                    <option value="pii">PII</option>
                    <option value="sensitive">Sensitive</option>
                  </select>
                </label>
                <label className={styles.formFieldSpan}>
                  <span>Placeholder</span>
                  <input className={styles.input} onChange={(event) => setFieldDraft((current) => ({ ...current, placeholder: event.target.value }))} value={fieldDraft.placeholder} />
                </label>
              </div>
              <div className={styles.inlineList}>
                <label className={styles.checkboxField}>
                  <input checked={fieldDraft.required} onChange={(event) => setFieldDraft((current) => ({ ...current, required: event.target.checked }))} type="checkbox" />
                  <span>Required</span>
                </label>
                <label className={styles.checkboxField}>
                  <input checked={fieldDraft.unique} onChange={(event) => setFieldDraft((current) => ({ ...current, unique: event.target.checked }))} type="checkbox" />
                  <span>Unique</span>
                </label>
              </div>
              <div className={styles.actionsRow}>
                <button className={styles.secondaryButton} disabled={!selectedObject} onClick={() => void handleFieldSave()} type="button">
                  Add field
                </button>
              </div>
            </>
          ) : null}

          {activeTab === 2 ? (
            <div className={styles.emptyState}>Validation rules will be available in a future release. Define field-level and cross-field validation constraints here.</div>
          ) : null}
        </section>
      </div>
    );
  }

  function renderPagesWorkspace() {
    const activePageMenus = sortedMenus.filter((menu) => menu.pageKey === (pageDraft.key || selectedPage?.key));

    if (activeTab === 1) {
      return (
        <div className={styles.workspaceGrid}>
          <section className={styles.panel}>
            <div className={styles.sectionHeader}>
              <div>
                <p className={styles.cardEyebrow}>Pages</p>
                <h2>Route inventory</h2>
              </div>
            </div>
            <div className={styles.listStack}>
              {manifest.pages.map((page) => (
                <button
                  className={page.id === selectedPage?.id ? styles.activeListItem : styles.listItem}
                  key={page.id}
                  onClick={() => applySelectedPage(page)}
                  type="button"
                >
                  <span>{page.title}</span>
                  <small>/{page.route}</small>
                </button>
              ))}
            </div>
          </section>
          <section className={styles.panelWide}>
            <div className={styles.tableWrap}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th>Page</th>
                    <th>Route</th>
                    <th>Layout</th>
                    <th>Menu</th>
                  </tr>
                </thead>
                <tbody>
                  {manifest.pages.map((page) => (
                    <tr key={page.id}>
                      <td>{page.title}</td>
                      <td>/{page.route}</td>
                      <td>{page.layoutKey}</td>
                      <td>{sortedMenus.find((menu) => menu.pageKey === page.key)?.label ?? "Not in nav"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        </div>
      );
    }

    if (activeTab === 2) {
      return (
        <div className={styles.workspaceGrid}>
          <section className={styles.panel}>
            <div className={styles.sectionHeader}>
              <div>
                <p className={styles.cardEyebrow}>Page templates</p>
                <h2>Tenant-aware starters</h2>
              </div>
            </div>
            <div className={styles.listStack}>
              {bootstrap.designerCatalog.pageTemplates.map((template) => (
                <article className={styles.workflowCard} key={template.key}>
                  <strong>{template.label}</strong>
                  <p>{template.description}</p>
                  <div className={styles.inlineList}>
                    <span className={styles.inlineTag}>{template.source}</span>
                    <span className={styles.inlineTag}>{template.layout.sections.length} sections</span>
                  </div>
                  <button className={styles.secondaryButton} onClick={() => applyPageTemplate(template.key)} type="button">
                    Use template
                  </button>
                </article>
              ))}
            </div>
          </section>
          <section className={styles.panelWide}>
            <div className={styles.sectionHeader}>
              <div>
                <p className={styles.cardEyebrow}>Section templates</p>
                <h2>Reusable building blocks</h2>
              </div>
            </div>
            <div className={styles.tileGrid}>
              {bootstrap.designerCatalog.sectionTemplates.map((template) => (
                <article className={styles.metricCard} key={template.key}>
                  <span>{template.label}</span>
                  <strong>{template.section.components.length} components</strong>
                  <p className={styles.metricMeta}>{template.description}</p>
                  <button className={styles.secondaryButton} onClick={() => addLayoutSection(template.key)} type="button">
                    Add to page
                  </button>
                </article>
              ))}
            </div>
          </section>
        </div>
      );
    }

    return (
      <div className={styles.designerGrid}>
        <section className={styles.designerRail}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>Pages</p>
              <h2>Page tree</h2>
            </div>
            <button
              className={styles.secondaryButton}
              onClick={() => {
                pushDesignerHistory();
                setSelectedPageId("");
                setPageDraft(createBlankPageDraft());
                const blankLayout = createBlankLayout(createClientId("page"), "New page");
                setLayoutDraft(blankLayout);
                setSelectedSectionId(blankLayout.sections[0]?.id ?? "");
                setSelectedComponentId(blankLayout.sections[0]?.components[0]?.id ?? "");
                markDesignerDirty();
              }}
              type="button"
            >
              New page
            </button>
          </div>

          <div className={styles.listStack}>
            {manifest.pages.map((page) => (
              <button
                className={page.id === selectedPage?.id ? styles.activeListItem : styles.listItem}
                key={page.id}
                onClick={() => applySelectedPage(page)}
                type="button"
              >
                <span>{page.title}</span>
                <small>/{page.route}</small>
              </button>
            ))}
          </div>

          <div className={styles.sidebarSection}>
            <p className={styles.sidebarLabel}>Route and navigation</p>
            <div className={styles.sidebarMeta}>
              <span>Route</span>
              <strong>{pageDraft.route ? `/${pageDraft.route}` : "Unassigned"}</strong>
            </div>
            <div className={styles.sidebarMeta}>
              <span>Menu</span>
              <strong>{activePageMenus[0]?.label ?? "Hidden"}</strong>
            </div>
            <div className={styles.sidebarMeta}>
              <span>Home</span>
              <strong>{pageDraft.isHome ? "Yes" : "No"}</strong>
            </div>
          </div>

          <div className={styles.sidebarSection}>
            <p className={styles.sidebarLabel}>Templates</p>
            <div className={styles.listStack}>
              {bootstrap.designerCatalog.pageTemplates.slice(0, 4).map((template) => (
                <button className={styles.listItem} key={template.key} onClick={() => applyPageTemplate(template.key)} type="button">
                  <span>{template.label}</span>
                  <small>{template.source}</small>
                </button>
              ))}
            </div>
          </div>
        </section>

        <section className={styles.designerCanvasPanel}>
          <div className={styles.designerCanvasToolbar}>
            <div>
              <p className={styles.cardEyebrow}>Hybrid designer</p>
              <h2>{pageDraft.title || "Untitled page"}</h2>
              <p className={styles.helperCopy}>
                Governed layout primitives with responsive spans, rails, tabs, and drawer-ready sections. Autosave is{" "}
                <strong>{autosaveStatus}</strong>.
              </p>
            </div>
            <div className={styles.inlineList}>
              <button className={previewDevice === "desktop" ? styles.primaryButton : styles.secondaryButton} onClick={() => setPreviewDevice("desktop")} type="button">
                Desktop
              </button>
              <button className={previewDevice === "tablet" ? styles.primaryButton : styles.secondaryButton} onClick={() => setPreviewDevice("tablet")} type="button">
                Tablet
              </button>
              <button className={previewDevice === "mobile" ? styles.primaryButton : styles.secondaryButton} onClick={() => setPreviewDevice("mobile")} type="button">
                Mobile
              </button>
              <button className={styles.secondaryButton} disabled={designerHistory.length === 0} onClick={handleUndoDesignerChange} type="button">
                Undo
              </button>
              <button className={styles.primaryButton} onClick={() => void handleLayoutSave()} type="button">
                Save now
              </button>
              <Link className={styles.secondaryLink} href={`/platform/preview/${bootstrap.tenant.slug}/${pageDraft.route || ""}`} target="_blank">
                Open draft preview
              </Link>
            </div>
          </div>

          <div className={`${styles.designerCanvasViewport} ${styles[`designerCanvas${previewDevice.charAt(0).toUpperCase()}${previewDevice.slice(1)}`]}`}>
            {layoutDraft ? (
              layoutDraft.sections.map((section) => (
                <article className={styles.designerSectionCard} key={section.id}>
                  <div className={styles.designerSectionHeader}>
                    <div>
                      <p className={styles.cardEyebrow}>{section.placement.zone}</p>
                      <h3>{section.title}</h3>
                      <p className={styles.helperCopy}>
                        {section.kind} · {section.components.length} components · span {section.placement.span}
                      </p>
                    </div>
                    <div className={styles.inlineList}>
                      <button className={styles.ghostButton} onClick={() => { setSelectedSectionId(section.id); setSelectedComponentId(""); }} type="button">
                        Inspect
                      </button>
                      <button className={styles.ghostButton} onClick={() => duplicateSection(section.id)} type="button">
                        Duplicate
                      </button>
                      <button className={styles.ghostButton} onClick={() => addComponentToSection(section.id)} type="button">
                        Add component
                      </button>
                      <button className={styles.ghostButton} onClick={() => removeSection(section.id)} type="button">
                        Remove
                      </button>
                    </div>
                  </div>
                  <div className={styles.designerCanvasGrid}>
                    {section.components.map((component) => {
                      const span =
                        previewDevice === "mobile"
                          ? component.placement.responsive.mobileSpan ?? component.placement.span
                          : previewDevice === "tablet"
                            ? component.placement.responsive.tabletSpan ?? component.placement.span
                            : component.placement.responsive.desktopSpan ?? component.placement.span;

                      return (
                        <button
                          className={component.id === selectedComponent?.id ? `${styles.designerComponentCard} ${styles.designerComponentActive}` : styles.designerComponentCard}
                          draggable
                          key={component.id}
                          onClick={() => {
                            setSelectedSectionId(section.id);
                            setSelectedComponentId(component.id);
                          }}
                          onDragOver={(event) => event.preventDefault()}
                          onDragStart={() => setDraggedComponentId(component.id)}
                          onDrop={() => reorderDraggedComponent(section.id, component.id)}
                          style={{ gridColumn: `span ${Math.max(1, Math.min(12, span))}` }}
                          type="button"
                        >
                          <div className={styles.designerComponentHeader}>
                            <span className={styles.inlineTag}>{component.kind}</span>
                            <span className={styles.inlineTag}>{component.placement.zone}</span>
                          </div>
                          <strong>{component.title}</strong>
                          <p>{component.description ?? String(component.props.body ?? "Configure content, bindings, and responsive behavior.")}</p>
                          <div className={styles.inlineList}>
                            {component.binding?.objectKey || component.objectKey ? <span className={styles.inlineTag}>object</span> : null}
                            {component.binding?.workflowKey || component.workflowKey ? <span className={styles.inlineTag}>workflow</span> : null}
                            {component.binding?.agentId || component.agentId ? <span className={styles.inlineTag}>agent</span> : null}
                            <span className={styles.inlineTag}>span {span}</span>
                          </div>
                        </button>
                      );
                    })}
                  </div>
                </article>
              ))
            ) : (
              <div className={styles.emptyState}>Select a page or create a new one to start designing.</div>
            )}
          </div>
        </section>

        <section className={styles.designerInspector}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>Inspector</p>
              <h2>{selectedComponent ? selectedComponent.title : selectedSection ? selectedSection.title : pageDraft.title || "Page"}</h2>
            </div>
          </div>

          <div className={styles.designerInspectorGroup}>
            <label className={styles.formField}>
              <span>Page title</span>
              <input className={styles.input} onChange={(event) => updatePageDraftState((current) => ({ ...current, title: event.target.value }))} value={pageDraft.title} />
            </label>
            <label className={styles.formField}>
              <span>Page key</span>
              <input className={styles.input} onChange={(event) => updatePageDraftState((current) => ({ ...current, key: event.target.value }))} value={pageDraft.key} />
            </label>
            <label className={styles.formField}>
              <span>Route</span>
              <input className={styles.input} onChange={(event) => updatePageDraftState((current) => ({ ...current, route: event.target.value }))} value={pageDraft.route} />
            </label>
            <label className={styles.formField}>
              <span>Object binding</span>
              <select className={styles.select} onChange={(event) => updatePageDraftState((current) => ({ ...current, objectKey: event.target.value }))} value={pageDraft.objectKey}>
                <option value="">No object</option>
                {manifest.objects.map((objectDefinition) => (
                  <option key={objectDefinition.id} value={objectDefinition.key}>
                    {objectDefinition.label}
                  </option>
                ))}
              </select>
            </label>
            <label className={styles.checkboxField}>
              <input checked={pageDraft.isHome} onChange={(event) => updatePageDraftState((current) => ({ ...current, isHome: event.target.checked }))} type="checkbox" />
              <span>Use as home page</span>
            </label>
            <label className={styles.formFieldSpan}>
              <span>Description</span>
              <textarea className={styles.textarea} onChange={(event) => updatePageDraftState((current) => ({ ...current, description: event.target.value }))} value={pageDraft.description} />
            </label>
          </div>

          {selectedSection ? (
            <div className={styles.designerInspectorGroup}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Section</p>
                  <h3>{selectedSection.title}</h3>
                </div>
              </div>
              <label className={styles.formField}>
                <span>Title</span>
                <input className={styles.input} onChange={(event) => updateLayoutSection(selectedSection.id, (current) => ({ ...current, title: event.target.value }))} value={selectedSection.title} />
              </label>
              <label className={styles.formField}>
                <span>Kind</span>
                <select className={styles.select} onChange={(event) => updateLayoutSection(selectedSection.id, (current) => ({ ...current, kind: event.target.value as LayoutSectionDefinition["kind"] }))} value={selectedSection.kind}>
                  <option value="grid">Grid</option>
                  <option value="tabs">Tabs</option>
                  <option value="drawer">Drawer</option>
                </select>
              </label>
              <label className={styles.formField}>
                <span>Zone</span>
                <select className={styles.select} onChange={(event) => updateLayoutSection(selectedSection.id, (current) => ({ ...current, placement: { ...current.placement, zone: event.target.value as LayoutSectionDefinition["placement"]["zone"] } }))} value={selectedSection.placement.zone}>
                  <option value="header">Header</option>
                  <option value="main">Main</option>
                  <option value="rail">Rail</option>
                  <option value="footer">Footer</option>
                  <option value="drawer">Drawer</option>
                </select>
              </label>
              <label className={styles.formField}>
                <span>Min height</span>
                <input className={styles.input} min={0} onChange={(event) => updateLayoutSection(selectedSection.id, (current) => ({ ...current, placement: { ...current.placement, minHeight: Number(event.target.value) || 0 } }))} type="number" value={selectedSection.placement.minHeight} />
              </label>
            </div>
          ) : null}

          {selectedComponent ? (
            <div className={styles.designerInspectorGroup}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Component</p>
                  <h3>{selectedComponent.title}</h3>
                </div>
                <div className={styles.inlineList}>
                  <button className={styles.ghostButton} onClick={() => duplicateComponent(selectedSection!.id, selectedComponent.id)} type="button">
                    Duplicate
                  </button>
                  <button className={styles.ghostButton} onClick={() => removeComponent(selectedSection!.id, selectedComponent.id)} type="button">
                    Remove
                  </button>
                </div>
              </div>
              <label className={styles.formField}>
                <span>Title</span>
                <input className={styles.input} onChange={(event) => updateComponent(selectedSection!.id, selectedComponent.id, (current) => ({ ...current, title: event.target.value }))} value={selectedComponent.title} />
              </label>
              <label className={styles.formField}>
                <span>Kind</span>
                <select className={styles.select} onChange={(event) => updateComponent(selectedSection!.id, selectedComponent.id, (current) => ({ ...current, kind: event.target.value as LayoutComponentDefinition["kind"] }))} value={selectedComponent.kind}>
                  {bootstrap.designerCatalog.componentPresets.map((preset) => (
                    <option key={preset.key} value={preset.kind}>
                      {preset.label}
                    </option>
                  ))}
                </select>
              </label>
              <label className={styles.formField}>
                <span>Zone</span>
                <select className={styles.select} onChange={(event) => updateComponent(selectedSection!.id, selectedComponent.id, (current) => ({ ...current, placement: { ...current.placement, zone: event.target.value as LayoutComponentDefinition["placement"]["zone"] } }))} value={selectedComponent.placement.zone}>
                  <option value="header">Header</option>
                  <option value="main">Main</option>
                  <option value="rail">Rail</option>
                  <option value="footer">Footer</option>
                  <option value="drawer">Drawer</option>
                </select>
              </label>
              <label className={styles.formField}>
                <span>Desktop span</span>
                <input className={styles.input} max={12} min={1} onChange={(event) => updateComponent(selectedSection!.id, selectedComponent.id, (current) => ({ ...current, width: Number(event.target.value) || 12, placement: { ...current.placement, span: Number(event.target.value) || 12, responsive: { ...current.placement.responsive, desktopSpan: Number(event.target.value) || 12 } } }))} type="number" value={selectedComponent.placement.responsive.desktopSpan ?? selectedComponent.placement.span} />
              </label>
              <label className={styles.formField}>
                <span>Tablet span</span>
                <input className={styles.input} max={12} min={1} onChange={(event) => updateComponent(selectedSection!.id, selectedComponent.id, (current) => ({ ...current, placement: { ...current.placement, responsive: { ...current.placement.responsive, tabletSpan: Number(event.target.value) || 6 } } }))} type="number" value={selectedComponent.placement.responsive.tabletSpan ?? 6} />
              </label>
              <label className={styles.formField}>
                <span>Mobile span</span>
                <input className={styles.input} max={12} min={1} onChange={(event) => updateComponent(selectedSection!.id, selectedComponent.id, (current) => ({ ...current, placement: { ...current.placement, responsive: { ...current.placement.responsive, mobileSpan: Number(event.target.value) || 12 } } }))} type="number" value={selectedComponent.placement.responsive.mobileSpan ?? 12} />
              </label>
              <label className={styles.formField}>
                <span>Style preset</span>
                <input className={styles.input} onChange={(event) => updateComponent(selectedSection!.id, selectedComponent.id, (current) => ({ ...current, stylePreset: event.target.value }))} value={selectedComponent.stylePreset ?? ""} />
              </label>
              <label className={styles.formField}>
                <span>Object binding</span>
                <select className={styles.select} onChange={(event) => updateComponent(selectedSection!.id, selectedComponent.id, (current) => ({ ...current, objectKey: event.target.value || undefined, binding: { ...current.binding, objectKey: event.target.value || undefined } }))} value={selectedComponent.binding?.objectKey ?? selectedComponent.objectKey ?? ""}>
                  <option value="">None</option>
                  {manifest.objects.map((objectDefinition) => (
                    <option key={objectDefinition.id} value={objectDefinition.key}>
                      {objectDefinition.label}
                    </option>
                  ))}
                </select>
              </label>
              <label className={styles.formField}>
                <span>Workflow binding</span>
                <select className={styles.select} onChange={(event) => updateComponent(selectedSection!.id, selectedComponent.id, (current) => ({ ...current, workflowKey: event.target.value || undefined, binding: { ...current.binding, workflowKey: event.target.value || undefined } }))} value={selectedComponent.binding?.workflowKey ?? selectedComponent.workflowKey ?? ""}>
                  <option value="">None</option>
                  {manifest.workflows.map((workflow) => (
                    <option key={workflow.id} value={workflow.key}>
                      {workflow.name}
                    </option>
                  ))}
                </select>
              </label>
              <label className={styles.formField}>
                <span>Agent binding</span>
                <select className={styles.select} onChange={(event) => updateComponent(selectedSection!.id, selectedComponent.id, (current) => ({ ...current, agentId: event.target.value || undefined, binding: { ...current.binding, agentId: event.target.value || undefined } }))} value={selectedComponent.binding?.agentId ?? selectedComponent.agentId ?? ""}>
                  <option value="">None</option>
                  {manifest.agents.map((agent) => (
                    <option key={agent.id} value={agent.id}>
                      {agent.name}
                    </option>
                  ))}
                </select>
              </label>
              <label className={styles.formFieldSpan}>
                <span>Visibility rule</span>
                <input className={styles.input} onChange={(event) => updateComponent(selectedSection!.id, selectedComponent.id, (current) => ({ ...current, visibilityRule: event.target.value ? { expression: event.target.value } : undefined }))} value={selectedComponent.visibilityRule?.expression ?? ""} />
              </label>
              <label className={styles.formFieldSpan}>
                <span>Body / note</span>
                <textarea className={styles.textarea} onChange={(event) => updateComponent(selectedSection!.id, selectedComponent.id, (current) => ({ ...current, props: { ...current.props, body: event.target.value } }))} value={String(selectedComponent.props.body ?? "")} />
              </label>
            </div>
          ) : null}

          <div className={styles.designerInspectorGroup}>
            <div className={styles.sectionHeader}>
              <div>
                <p className={styles.cardEyebrow}>Preset library</p>
                <h3>Component presets</h3>
              </div>
            </div>
            <div className={styles.listStack}>
              {bootstrap.designerCatalog.componentPresets.map((preset) => (
                <button
                  className={styles.listItem}
                  key={preset.key}
                  onClick={() => {
                    if (selectedSection) {
                      addComponentToSection(selectedSection.id, preset.key);
                    }
                  }}
                  type="button"
                >
                  <span>{preset.label}</span>
                  <small>{preset.description}</small>
                </button>
              ))}
            </div>
          </div>
        </section>
      </div>
    );
  }

  function renderNavigationWorkspace() {
    return (
      <div className={styles.workspaceGrid}>
        {activeTab === 0 ? (
          <>
            <section className={styles.panel}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Published navigation</p>
                  <h2>Menu items</h2>
                </div>
              </div>
              <div className={styles.listStack}>
                {sortedMenus.map((menu) => (
                  <div className={styles.fieldCard} key={menu.id}>
                    <div>
                      <strong>{menu.label}</strong>
                      <p>
                        {menu.group} · {menu.pageKey}
                      </p>
                    </div>
                    <div className={styles.inlineList}>
                      <button className={styles.ghostButton} onClick={() => moveMenu(menu, -1)} type="button">
                        Up
                      </button>
                      <button className={styles.ghostButton} onClick={() => moveMenu(menu, 1)} type="button">
                        Down
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </section>
            <section className={styles.panelWide}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Add menu item</p>
                  <h2>Navigation builder</h2>
                </div>
              </div>
              <div className={styles.formGrid}>
                <label className={styles.formField}>
                  <span>Label</span>
                  <input className={styles.input} onChange={(event) => setMenuDraft((current) => ({ ...current, label: event.target.value }))} value={menuDraft.label} />
                </label>
                <label className={styles.formField}>
                  <span>Key</span>
                  <input className={styles.input} onChange={(event) => setMenuDraft((current) => ({ ...current, key: event.target.value }))} value={menuDraft.key} />
                </label>
                <label className={styles.formField}>
                  <span>Page</span>
                  <select className={styles.select} onChange={(event) => setMenuDraft((current) => ({ ...current, pageKey: event.target.value }))} value={menuDraft.pageKey}>
                    <option value="">Choose page</option>
                    {manifest.pages.map((page) => (
                      <option key={page.id} value={page.key}>
                        {page.title}
                      </option>
                    ))}
                  </select>
                </label>
                <label className={styles.formField}>
                  <span>Group</span>
                  <input className={styles.input} onChange={(event) => setMenuDraft((current) => ({ ...current, group: event.target.value }))} value={menuDraft.group} />
                </label>
              </div>
              <div className={styles.actionsRow}>
                <button className={styles.primaryButton} onClick={() => void handleMenuSave()} type="button">
                  Save menu item
                </button>
              </div>
            </section>
          </>
        ) : null}

        {activeTab === 1 ? (
          <section className={styles.panelWide}>
            <div className={styles.emptyState}>Route configuration and URL management will be available in a future release.</div>
          </section>
        ) : null}
      </div>
    );
  }

  function renderWorkflowsWorkspace() {
    return (
      <div className={styles.workspaceGrid}>
        <section className={styles.panel}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>Workflow graph</p>
              <h2>Defined workflows</h2>
            </div>
            <button className={styles.secondaryButton} onClick={() => selectWorkflow(undefined)} type="button">
              New workflow
            </button>
          </div>
          <div className={styles.listStack}>
            {manifest.workflows.map((workflow) => (
              <button
                className={workflow.id === selectedWorkflow?.id ? styles.activeListItem : styles.listItem}
                key={workflow.id}
                onClick={() => selectWorkflow(workflow)}
                type="button"
              >
                <span>{workflow.name}</span>
                <small>
                  {workflow.nodes.length} nodes · {workflow.edges.length} edges · {workflow.status}
                </small>
              </button>
            ))}
          </div>
        </section>
        <section className={styles.panelWide}>
          {activeTab === 0 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Workflow foundation</p>
                  <h2>{selectedWorkflow ? selectedWorkflow.name : "Create workflow"}</h2>
                </div>
                {selectedWorkflow ? (
                  <button className={styles.secondaryButton} onClick={() => void handleQueueWorkflowRun(selectedWorkflow.id)} type="button">
                    Queue run
                  </button>
                ) : null}
              </div>
              <div className={styles.formGrid}>
                <label className={styles.formField}>
                  <span>Name</span>
                  <input className={styles.input} onChange={(event) => setWorkflowDraft((current) => ({ ...current, name: event.target.value }))} value={workflowDraft.name} />
                </label>
                <label className={styles.formField}>
                  <span>Key</span>
                  <input className={styles.input} onChange={(event) => setWorkflowDraft((current) => ({ ...current, key: event.target.value }))} value={workflowDraft.key} />
                </label>
                <label className={styles.formField}>
                  <span>Object</span>
                  <select className={styles.select} onChange={(event) => setWorkflowDraft((current) => ({ ...current, objectKey: event.target.value || undefined }))} value={workflowDraft.objectKey ?? ""}>
                    <option value="">No object</option>
                    {manifest.objects.map((objectDefinition) => (
                      <option key={objectDefinition.id} value={objectDefinition.key}>
                        {objectDefinition.label}
                      </option>
                    ))}
                  </select>
                </label>
                <label className={styles.formFieldSpan}>
                  <span>Description</span>
                  <textarea className={styles.textarea} onChange={(event) => setWorkflowDraft((current) => ({ ...current, description: event.target.value }))} value={workflowDraft.description ?? ""} />
                </label>
              </div>
              <div className={styles.workflowEditorGrid}>
                <article className={styles.workflowCanvasCard}>
                  <div className={styles.sectionHeader}>
                    <div>
                      <p className={styles.cardEyebrow}>Node palette</p>
                      <h3>Build the graph</h3>
                    </div>
                  </div>
                  <div className={styles.nodePalette}>
                    {WORKFLOW_NODE_LIBRARY.map((nodeType) => (
                      <button className={styles.paletteButton} key={nodeType.type} onClick={() => addWorkflowNode(nodeType.type)} type="button">
                        <strong>{nodeType.title}</strong>
                        <span>{nodeType.note}</span>
                      </button>
                    ))}
                  </div>
                  <div className={styles.canvasSurface}>
                    {workflowDraft.nodes.map((node) => (
                      <button
                        className={node.id === selectedWorkflowNode?.id ? styles.activeCanvasNode : styles.canvasNode}
                        key={node.id}
                        onClick={() => setSelectedWorkflowNodeId(node.id)}
                        style={{ left: node.position.x, top: node.position.y }}
                        type="button"
                      >
                        <small>{node.type.replace(/_/g, " ")}</small>
                        <strong>{node.label}</strong>
                      </button>
                    ))}
                    {workflowDraft.nodes.length === 0 ? (
                      <div className={styles.emptyCanvasState}>Add nodes from the palette to start wiring the workflow.</div>
                    ) : null}
                  </div>
                  <div className={styles.workflowEdges}>
                    <div className={styles.sectionHeader}>
                      <div>
                        <p className={styles.cardEyebrow}>Edges</p>
                        <h3>Branch wiring</h3>
                      </div>
                    </div>
                    <div className={styles.formGridTight}>
                      <label className={styles.formField}>
                        <span>Source</span>
                        <select
                          className={styles.select}
                          onChange={(event) => setWorkflowEdgeDraft((current) => ({ ...current, sourceId: event.target.value }))}
                          value={workflowEdgeDraft.sourceId}
                        >
                          <option value="">Choose node</option>
                          {workflowDraft.nodes.map((node) => (
                            <option key={node.id} value={node.id}>
                              {node.label}
                            </option>
                          ))}
                        </select>
                      </label>
                      <label className={styles.formField}>
                        <span>Target</span>
                        <select
                          className={styles.select}
                          onChange={(event) => setWorkflowEdgeDraft((current) => ({ ...current, targetId: event.target.value }))}
                          value={workflowEdgeDraft.targetId}
                        >
                          <option value="">Choose node</option>
                          {workflowDraft.nodes.map((node) => (
                            <option key={node.id} value={node.id}>
                              {node.label}
                            </option>
                          ))}
                        </select>
                      </label>
                      <label className={styles.formField}>
                        <span>Label</span>
                        <input
                          className={styles.input}
                          onChange={(event) => setWorkflowEdgeDraft((current) => ({ ...current, label: event.target.value }))}
                          value={workflowEdgeDraft.label}
                        />
                      </label>
                    </div>
                    <div className={styles.actionsRow}>
                      <button className={styles.secondaryButton} onClick={addWorkflowEdge} type="button">
                        Add edge
                      </button>
                    </div>
                    <div className={styles.listStack}>
                      {workflowDraft.edges.map((edge) => (
                        <div className={styles.edgeCard} key={edge.id}>
                          <div>
                            <strong>
                              {workflowDraft.nodes.find((node) => node.id === edge.sourceId)?.label ?? edge.sourceId}
                              {" -> "}
                              {workflowDraft.nodes.find((node) => node.id === edge.targetId)?.label ?? edge.targetId}
                            </strong>
                            <p>{edge.label || "Unlabelled path"}</p>
                          </div>
                          <button className={styles.ghostButtonDanger} onClick={() => removeWorkflowEdge(edge.id)} type="button">
                            Remove
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>
                </article>
                <article className={styles.workflowInspector}>
                  <div className={styles.sectionHeader}>
                    <div>
                      <p className={styles.cardEyebrow}>Node config</p>
                      <h3>{selectedWorkflowNode ? selectedWorkflowNode.label : "Select a node"}</h3>
                    </div>
                    {selectedWorkflowNode ? (
                      <button className={styles.ghostButtonDanger} onClick={() => removeWorkflowNode(selectedWorkflowNode.id)} type="button">
                        Delete
                      </button>
                    ) : null}
                  </div>
                  {selectedWorkflowNode ? (
                    <>
                      <div className={styles.formGridTight}>
                        <label className={styles.formField}>
                          <span>Label</span>
                          <input
                            className={styles.input}
                            onChange={(event) => updateWorkflowNode(selectedWorkflowNode.id, (current) => ({ ...current, label: event.target.value }))}
                            value={selectedWorkflowNode.label}
                          />
                        </label>
                        <label className={styles.formField}>
                          <span>X</span>
                          <input
                            className={styles.input}
                            onChange={(event) =>
                              updateWorkflowNode(selectedWorkflowNode.id, (current) => ({
                                ...current,
                                position: { ...current.position, x: Number(event.target.value) || 0 },
                              }))
                            }
                            type="number"
                            value={selectedWorkflowNode.position.x}
                          />
                        </label>
                        <label className={styles.formField}>
                          <span>Y</span>
                          <input
                            className={styles.input}
                            onChange={(event) =>
                              updateWorkflowNode(selectedWorkflowNode.id, (current) => ({
                                ...current,
                                position: { ...current.position, y: Number(event.target.value) || 0 },
                              }))
                            }
                            type="number"
                            value={selectedWorkflowNode.position.y}
                          />
                        </label>
                      </div>
                      <div className={styles.formGridTight}>
                        {"expression" in selectedWorkflowNode.config ? (
                          <label className={styles.formFieldSpan}>
                            <span>Expression</span>
                            <textarea
                              className={styles.textarea}
                              onChange={(event) =>
                                updateWorkflowNode(selectedWorkflowNode.id, (current) => ({
                                  ...current,
                                  config: { ...current.config, expression: event.target.value },
                                }))
                              }
                              value={selectedWorkflowNode.config.expression}
                            />
                          </label>
                        ) : null}
                        {"operation" in selectedWorkflowNode.config ? (
                          <>
                            <label className={styles.formField}>
                              <span>Operation</span>
                              <select
                                className={styles.select}
                                onChange={(event) =>
                                  updateWorkflowNode(selectedWorkflowNode.id, (current) => ({
                                    ...current,
                                    config: { ...current.config, operation: event.target.value as "create" | "update" | "delete" },
                                  }))
                                }
                                value={selectedWorkflowNode.config.operation}
                              >
                                <option value="create">Create</option>
                                <option value="update">Update</option>
                                <option value="delete">Delete</option>
                              </select>
                            </label>
                            <label className={styles.formField}>
                              <span>Object</span>
                              <select
                                className={styles.select}
                                onChange={(event) =>
                                  updateWorkflowNode(selectedWorkflowNode.id, (current) => ({
                                    ...current,
                                    config: { ...current.config, objectKey: event.target.value },
                                  }))
                                }
                                value={selectedWorkflowNode.config.objectKey}
                              >
                                {manifest.objects.map((objectDefinition) => (
                                  <option key={objectDefinition.id} value={objectDefinition.key}>
                                    {objectDefinition.label}
                                  </option>
                                ))}
                              </select>
                            </label>
                            <label className={styles.formFieldSpan}>
                              <span>Value expression</span>
                              <input
                                className={styles.input}
                                onChange={(event) =>
                                  updateWorkflowNode(selectedWorkflowNode.id, (current) => ({
                                    ...current,
                                    config: { ...current.config, valueExpression: event.target.value },
                                  }))
                                }
                                value={selectedWorkflowNode.config.valueExpression ?? ""}
                              />
                            </label>
                          </>
                        ) : null}
                        {"method" in selectedWorkflowNode.config ? (
                          <>
                            <label className={styles.formField}>
                              <span>Method</span>
                              <select
                                className={styles.select}
                                onChange={(event) =>
                                  updateWorkflowNode(selectedWorkflowNode.id, (current) => ({
                                    ...current,
                                    config: { ...current.config, method: event.target.value as "GET" | "POST" | "PUT" | "PATCH" },
                                  }))
                                }
                                value={selectedWorkflowNode.config.method}
                              >
                                <option value="GET">GET</option>
                                <option value="POST">POST</option>
                                <option value="PUT">PUT</option>
                                <option value="PATCH">PATCH</option>
                              </select>
                            </label>
                            <label className={styles.formFieldSpan}>
                              <span>URL</span>
                              <input
                                className={styles.input}
                                onChange={(event) =>
                                  updateWorkflowNode(selectedWorkflowNode.id, (current) => ({
                                    ...current,
                                    config: { ...current.config, url: event.target.value },
                                  }))
                                }
                                value={selectedWorkflowNode.config.url}
                              />
                            </label>
                          </>
                        ) : null}
                        {"channel" in selectedWorkflowNode.config ? (
                          <>
                            <label className={styles.formField}>
                              <span>Channel</span>
                              <select
                                className={styles.select}
                                onChange={(event) =>
                                  updateWorkflowNode(selectedWorkflowNode.id, (current) => ({
                                    ...current,
                                    config: { ...current.config, channel: event.target.value as "email" | "slack" | "task" },
                                  }))
                                }
                                value={selectedWorkflowNode.config.channel}
                              >
                                <option value="email">Email</option>
                                <option value="slack">Slack</option>
                                <option value="task">Task</option>
                              </select>
                            </label>
                            <label className={styles.formFieldSpan}>
                              <span>Recipient</span>
                              <input
                                className={styles.input}
                                onChange={(event) =>
                                  updateWorkflowNode(selectedWorkflowNode.id, (current) => ({
                                    ...current,
                                    config: { ...current.config, recipient: event.target.value },
                                  }))
                                }
                                value={selectedWorkflowNode.config.recipient ?? ""}
                              />
                            </label>
                          </>
                        ) : null}
                        {"durationMinutes" in selectedWorkflowNode.config ? (
                          <label className={styles.formField}>
                            <span>Duration (minutes)</span>
                            <input
                              className={styles.input}
                              onChange={(event) =>
                                updateWorkflowNode(selectedWorkflowNode.id, (current) => ({
                                  ...current,
                                  config: { ...current.config, durationMinutes: Number(event.target.value) || 1 },
                                }))
                              }
                              type="number"
                              value={selectedWorkflowNode.config.durationMinutes}
                            />
                          </label>
                        ) : null}
                        {"approverRole" in selectedWorkflowNode.config ? (
                          <label className={styles.formField}>
                            <span>Approver role</span>
                            <select
                              className={styles.select}
                              onChange={(event) =>
                                updateWorkflowNode(selectedWorkflowNode.id, (current) => ({
                                  ...current,
                                  config: { ...current.config, approverRole: event.target.value as "SUPER_ADMIN" | "BUILDER_ADMIN" | "USER" },
                                }))
                              }
                              value={selectedWorkflowNode.config.approverRole}
                            >
                              <option value="SUPER_ADMIN">Super admin</option>
                              <option value="BUILDER_ADMIN">Builder admin</option>
                              <option value="USER">User</option>
                            </select>
                          </label>
                        ) : null}
                        {"agentId" in selectedWorkflowNode.config ? (
                          <label className={styles.formField}>
                            <span>Agent</span>
                            <select
                              className={styles.select}
                              onChange={(event) =>
                                updateWorkflowNode(selectedWorkflowNode.id, (current) => ({
                                  ...current,
                                  config: { ...current.config, agentId: event.target.value },
                                }))
                              }
                              value={selectedWorkflowNode.config.agentId}
                            >
                              <option value="">Choose agent</option>
                              {manifest.agents.map((agent) => (
                                <option key={agent.id} value={agent.id}>
                                  {agent.name}
                                </option>
                              ))}
                            </select>
                          </label>
                        ) : null}
                      </div>
                    </>
                  ) : (
                    <div className={styles.emptyState}>Select a node from the canvas to edit its configuration.</div>
                  )}
                </article>
              </div>
              <p className={styles.helperCopy}>Durable worker execution now reads only the active published version. Draft-only workflows stay private until publish.</p>
              <div className={styles.actionsRow}>
                <button className={styles.primaryButton} onClick={() => void handleWorkflowSave()} type="button">
                  Save workflow
                </button>
              </div>
            </>
          ) : null}

          {activeTab === 1 ? (
            <div className={styles.listStack}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Run history</p>
                  <h3>{selectedWorkflow ? selectedWorkflow.name : "No workflow selected"}</h3>
                </div>
                {selectedWorkflow ? (
                  <button className={styles.primaryButton} onClick={() => void handleQueueWorkflowRun(selectedWorkflow.id)} type="button">
                    Queue run
                  </button>
                ) : null}
              </div>
              {workflowRuns.length === 0 ? (
                <div className={styles.emptyState}>No runs yet. Publish the workflow and queue a run to inspect worker output.</div>
              ) : (
                workflowRuns.map((run) => (
                  <article className={styles.runCard} key={run.id}>
                    <div className={styles.runCardHeader}>
                      <div>
                        <strong>{run.workflowKey}</strong>
                        <p>
                          {run.status} · {formatPlatformDateTime(run.createdAt)}
                        </p>
                      </div>
                      <span className={styles.badge}>{run.logs.length} log events</span>
                    </div>
                    {run.logs.length > 0 ? (
                      <div className={styles.logStack}>
                        {run.logs.map((log, index) => (
                          <div className={styles.logRow} key={`${run.id}-${index}`}>
                            <strong>{String(log.level ?? "info").toUpperCase()}</strong>
                            <span>{String(log.message ?? "")}</span>
                          </div>
                        ))}
                      </div>
                    ) : null}
                  </article>
                ))
              )}
            </div>
          ) : null}
        </section>
      </div>
    );
  }

  function renderAgentsWorkspace() {
    const selectedProvider =
      manifest.modelProviders.find((provider) => provider.id === agentDraft.modelProviderId || provider.key === agentDraft.modelProviderId) ??
      manifest.modelProviders[0];

    return (
      <div className={styles.workspaceGrid}>
        <section className={styles.panel}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>Agents</p>
              <h2>Tenant agent registry</h2>
            </div>
            <button className={styles.secondaryButton} onClick={() => selectAgent(undefined)} type="button">
              New agent
            </button>
          </div>
          <div className={styles.listStack}>
            {manifest.agents.map((agent) => (
              <button
                className={agent.id === selectedAgent?.id ? styles.activeListItem : styles.listItem}
                key={agent.id}
                onClick={() => selectAgent(agent)}
                type="button"
              >
                <span>{agent.name}</span>
                <small>
                  {agent.scope} · {agent.zeroRetentionRequired ? "zero retention" : "standard retention"}
                </small>
              </button>
            ))}
          </div>
        </section>
        <section className={styles.panelWide}>
          {activeTab === 0 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Create agent</p>
                  <h2>{selectedAgent ? selectedAgent.name : "AI control plane"}</h2>
                </div>
              </div>
              <div className={styles.formGrid}>
                <label className={styles.formField}>
                  <span>Name</span>
                  <input className={styles.input} onChange={(event) => setAgentDraft((current) => ({ ...current, name: event.target.value }))} value={agentDraft.name} />
                </label>
                <label className={styles.formField}>
                  <span>Key</span>
                  <input className={styles.input} onChange={(event) => setAgentDraft((current) => ({ ...current, key: event.target.value }))} value={agentDraft.key} />
                </label>
                <label className={styles.formField}>
                  <span>Scope</span>
                  <select className={styles.select} onChange={(event) => setAgentDraft((current) => ({ ...current, scope: event.target.value as AgentDefinition["scope"] }))} value={agentDraft.scope}>
                    <option value="workspace">Workspace</option>
                    <option value="node">Node</option>
                  </select>
                </label>
                <label className={styles.formField}>
                  <span>Model provider</span>
                  <select className={styles.select} onChange={(event) => setAgentDraft((current) => ({ ...current, modelProviderId: event.target.value }))} value={agentDraft.modelProviderId}>
                    <option value="">Choose provider</option>
                    {manifest.modelProviders.map((provider) => (
                      <option key={provider.id} value={provider.id}>
                        {provider.name}
                      </option>
                    ))}
                  </select>
                </label>
                <label className={styles.formFieldSpan}>
                  <span>Description</span>
                  <textarea className={styles.textarea} onChange={(event) => setAgentDraft((current) => ({ ...current, description: event.target.value }))} value={agentDraft.description ?? ""} />
                </label>
                <label className={styles.formFieldSpan}>
                  <span>Prompt</span>
                  <textarea className={styles.textarea} onChange={(event) => setAgentDraft((current) => ({ ...current, prompt: event.target.value }))} value={agentDraft.prompt} />
                </label>
              </div>
              <div className={styles.inlineList}>
                <label className={styles.checkboxField}>
                  <input checked={agentDraft.zeroRetentionRequired} onChange={(event) => setAgentDraft((current) => ({ ...current, zeroRetentionRequired: event.target.checked }))} type="checkbox" />
                  <span>Zero retention required</span>
                </label>
              </div>
              {selectedProvider ? (
                <div className={styles.agentPolicyCard}>
                  <div>
                    <p className={styles.cardEyebrow}>Security posture</p>
                    <h3>{selectedProvider.name}</h3>
                  </div>
                  <div className={styles.inlineList}>
                    <span className={styles.inlineTag}>{selectedProvider.provider}</span>
                    <span className={styles.inlineTag}>{selectedProvider.supportsZeroRetention ? "zero retention" : "retention enabled"}</span>
                    <span className={styles.inlineTag}>{selectedProvider.allowedForSensitiveData ? "sensitive ok" : "masked only"}</span>
                  </div>
                </div>
              ) : null}
              <div className={styles.actionsRow}>
                <button className={styles.primaryButton} onClick={() => void handleAgentSave()} type="button">
                  Save agent
                </button>
              </div>
            </>
          ) : null}

          {activeTab === 1 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Prompt and tools</p>
                  <h3>Guarded builder</h3>
                </div>
                <button className={styles.secondaryButton} disabled={!agentDraft.id} onClick={() => void handlePreviewAgent(agentDraft.id)} type="button">
                  Preview masked invocation
                </button>
              </div>
              <div className={styles.formGrid}>
                <label className={styles.formFieldSpan}>
                  <span>System prompt</span>
                  <textarea className={styles.textareaTall} onChange={(event) => setAgentDraft((current) => ({ ...current, prompt: event.target.value }))} value={agentDraft.prompt} />
                </label>
              </div>
              <div className={styles.scopeGrid}>
                <article className={styles.scopeCard}>
                  <p className={styles.cardEyebrow}>Allowed tools</p>
                  <div className={styles.listStack}>
                    {manifest.tools.map((tool) => (
                      <label className={styles.checkboxField} key={tool.id}>
                        <input checked={agentDraft.allowedToolIds.includes(tool.id)} onChange={() => toggleAgentTool(tool.id)} type="checkbox" />
                        <span>
                          {tool.name} · {tool.type}
                        </span>
                      </label>
                    ))}
                  </div>
                </article>
                <article className={styles.scopeCard}>
                  <p className={styles.cardEyebrow}>Preview output</p>
                  {agentPreview ? (
                    <div className={styles.previewStack}>
                      <div className={styles.sidebarMeta}>
                        <span>Provider</span>
                        <strong>{agentPreview.provider.name}</strong>
                      </div>
                      <div className={styles.sidebarMeta}>
                        <span>Sample size</span>
                        <strong>{agentPreview.sampleSize}</strong>
                      </div>
                      <div className={styles.sidebarPanel}>
                        Masked: {String(agentPreview.metadata.masked)} · Zero retention: {String(agentPreview.metadata.zeroRetentionRequired)}
                      </div>
                      <pre className={styles.previewCode}>{JSON.stringify(agentPreview.maskedRecords, null, 2)}</pre>
                    </div>
                  ) : (
                    <div className={styles.emptyState}>Run a preview to inspect masked records before wiring this agent into runtime flows.</div>
                  )}
                </article>
              </div>
            </>
          ) : null}

          {activeTab === 2 ? (
            <div className={styles.scopeGrid}>
              <article className={styles.scopeCard}>
                <div className={styles.sectionHeader}>
                  <div>
                    <p className={styles.cardEyebrow}>Object scope</p>
                    <h3>Access envelope</h3>
                  </div>
                </div>
                <div className={styles.listStack}>
                  {manifest.objects.map((objectDefinition) => (
                    <label className={styles.checkboxField} key={objectDefinition.id}>
                      <input checked={agentDraft.objectKeys.includes(objectDefinition.key)} onChange={() => toggleAgentObject(objectDefinition.key)} type="checkbox" />
                      <span>
                        {objectDefinition.label} · {objectDefinition.fields.length} fields
                      </span>
                    </label>
                  ))}
                </div>
              </article>
              <article className={styles.scopeCard}>
                <div className={styles.sectionHeader}>
                  <div>
                    <p className={styles.cardEyebrow}>Recent usage</p>
                    <h3>Operator visibility</h3>
                  </div>
                </div>
                {agentActivity.length === 0 ? (
                  <div className={styles.emptyState}>No activity yet. Agent previews and future runtime calls will appear here.</div>
                ) : (
                  <div className={styles.listStack}>
                    {agentActivity.map((event) => (
                      <article className={styles.auditRow} key={event.id}>
                        <div>
                          <strong>{event.summary}</strong>
                          <p>{event.actorEmail ?? "system"}</p>
                        </div>
                        <span>{formatPlatformDateTime(event.createdAt)}</span>
                      </article>
                    ))}
                  </div>
                )}
              </article>
            </div>
          ) : null}
        </section>
      </div>
    );
  }

  function renderModelsWorkspace() {
    return (
      <div className={styles.workspaceGrid}>
        <section className={styles.panel}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>Provider registry</p>
              <h2>Models</h2>
            </div>
          </div>
          <div className={styles.listStack}>
            {manifest.modelProviders.map((provider) => (
              <article className={styles.workflowCard} key={provider.id}>
                <strong>{provider.name}</strong>
                <p>
                  {provider.provider} · {provider.model}
                </p>
                <div className={styles.inlineList}>
                  <span className={styles.inlineTag}>{provider.status}</span>
                  <span className={styles.inlineTag}>{provider.supportsZeroRetention ? "zero retention" : "standard"}</span>
                </div>
              </article>
            ))}
          </div>
        </section>
        <section className={styles.panelWide}>
          {activeTab === 0 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>God-tier admin</p>
                  <h2>Register model provider</h2>
                </div>
                <span className={styles.badge}>{actor.role}</span>
              </div>
              <div className={styles.formGrid}>
                <label className={styles.formField}>
                  <span>Name</span>
                  <input className={styles.input} onChange={(event) => setProviderDraft((current) => ({ ...current, name: event.target.value }))} value={providerDraft.name} />
                </label>
                <label className={styles.formField}>
                  <span>Key</span>
                  <input className={styles.input} onChange={(event) => setProviderDraft((current) => ({ ...current, key: event.target.value }))} value={providerDraft.key} />
                </label>
                <label className={styles.formField}>
                  <span>Provider</span>
                  <select className={styles.select} onChange={(event) => setProviderDraft((current) => ({ ...current, provider: event.target.value as ModelProviderDefinition["provider"] }))} value={providerDraft.provider}>
                    <option value="openai">OpenAI</option>
                    <option value="azure_openai">Azure OpenAI</option>
                    <option value="anthropic">Anthropic</option>
                    <option value="google">Google</option>
                    <option value="custom">Custom</option>
                  </select>
                </label>
                <label className={styles.formField}>
                  <span>Model</span>
                  <input className={styles.input} onChange={(event) => setProviderDraft((current) => ({ ...current, model: event.target.value }))} value={providerDraft.model} />
                </label>
                <label className={styles.formFieldSpan}>
                  <span>Secret reference</span>
                  <input className={styles.input} onChange={(event) => setProviderDraft((current) => ({ ...current, apiKeySecretRef: event.target.value }))} value={providerDraft.apiKeySecretRef} />
                </label>
              </div>
              <div className={styles.inlineList}>
                <label className={styles.checkboxField}>
                  <input checked={providerDraft.supportsZeroRetention} onChange={(event) => setProviderDraft((current) => ({ ...current, supportsZeroRetention: event.target.checked }))} type="checkbox" />
                  <span>Supports zero retention</span>
                </label>
                <label className={styles.checkboxField}>
                  <input checked={providerDraft.allowedForSensitiveData} onChange={(event) => setProviderDraft((current) => ({ ...current, allowedForSensitiveData: event.target.checked }))} type="checkbox" />
                  <span>Allowed for sensitive data</span>
                </label>
              </div>
              <div className={styles.actionsRow}>
                <button className={styles.primaryButton} disabled={actor.role !== "SUPER_ADMIN"} onClick={() => void handleProviderSave()} type="button">
                  Save provider
                </button>
              </div>
            </>
          ) : null}

          {activeTab === 1 ? (
            <div className={styles.emptyState}>Model configuration, rate limits, and usage monitoring will be available in a future release.</div>
          ) : null}
        </section>
      </div>
    );
  }

  function renderSecurityWorkspace() {
    return (
      <div className={styles.workspaceGrid}>
        <section className={styles.panelWide}>
          {activeTab === 0 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Security policy</p>
                  <h2>Masked-by-default AI controls</h2>
                </div>
              </div>
              <div className={styles.inlineList}>
                <label className={styles.checkboxField}>
                  <input
                    checked={securityDraft.zeroRetentionRequiredForSensitiveData}
                    onChange={(event) =>
                      setSecurityDraft((current) => ({
                        ...current,
                        zeroRetentionRequiredForSensitiveData: event.target.checked,
                      }))
                    }
                    type="checkbox"
                  />
                  <span>Require zero retention for sensitive data</span>
                </label>
              </div>
              <div className={styles.inlineList}>
                {manifest.modelProviders.map((provider) => {
                  const checked = securityDraft.allowedModelProviderKeys.includes(provider.key);
                  return (
                    <label className={styles.checkboxField} key={provider.id}>
                      <input
                        checked={checked}
                        onChange={(event) =>
                          setSecurityDraft((current) => ({
                            ...current,
                            allowedModelProviderKeys: event.target.checked
                              ? [...current.allowedModelProviderKeys, provider.key]
                              : current.allowedModelProviderKeys.filter((candidate) => candidate !== provider.key),
                          }))
                        }
                        type="checkbox"
                      />
                      <span>{provider.name}</span>
                    </label>
                  );
                })}
              </div>
              <div className={styles.actionsRow}>
                <button className={styles.primaryButton} disabled={actor.role !== "SUPER_ADMIN"} onClick={() => void handleSecuritySave()} type="button">
                  Save security policy
                </button>
              </div>
            </>
          ) : null}

          {activeTab === 1 ? (
            <div className={styles.scopeGrid}>
              <article className={styles.scopeCard}>
                <div className={styles.sectionHeader}>
                  <div>
                    <p className={styles.cardEyebrow}>Tenant access</p>
                    <h2>Invite builder access</h2>
                  </div>
                  <span className={styles.badge}>{actor.role}</span>
                </div>
                <div className={styles.formGrid}>
                  <label className={styles.formField}>
                    <span>Email</span>
                    <input className={styles.input} onChange={(event) => setInviteDraft((current) => ({ ...current, email: event.target.value }))} value={inviteDraft.email} />
                  </label>
                  <label className={styles.formField}>
                    <span>Role</span>
                    <select className={styles.select} onChange={(event) => setInviteDraft((current) => ({ ...current, role: event.target.value as PlatformRole }))} value={inviteDraft.role}>
                      <option value="SUPER_ADMIN">Super admin</option>
                      <option value="BUILDER_ADMIN">Builder admin</option>
                      <option value="USER">User</option>
                    </select>
                  </label>
                  <label className={styles.formField}>
                    <span>Expires (days)</span>
                    <input className={styles.input} max={30} min={1} onChange={(event) => setInviteDraft((current) => ({ ...current, expiresInDays: Number(event.target.value) || 7 }))} type="number" value={inviteDraft.expiresInDays} />
                  </label>
                </div>
                <div className={styles.actionsRow}>
                  <button className={styles.primaryButton} disabled={actor.role !== "SUPER_ADMIN"} onClick={() => void handleInviteCreate()} type="button">
                    Create invite
                  </button>
                </div>
              </article>
              <article className={styles.scopeCard}>
                <div className={styles.sectionHeader}>
                  <div>
                    <p className={styles.cardEyebrow}>Pending invites</p>
                    <h2>Manual tenant onboarding</h2>
                  </div>
                </div>
                {bootstrap.invites.length === 0 ? (
                  <div className={styles.emptyState}>No invites have been issued for this tenant yet.</div>
                ) : (
                  <div className={styles.listStack}>
                    {bootstrap.invites.map((invite) => (
                      <article className={styles.auditRow} key={invite.id}>
                        <div>
                          <strong>{invite.email}</strong>
                          <p>
                            {invite.role} · {invite.status}
                          </p>
                          <p className={styles.helperCopy}>{invite.inviteUrl}</p>
                        </div>
                        <span>{formatPlatformDateTime(invite.createdAt)}</span>
                      </article>
                    ))}
                  </div>
                )}
              </article>
            </div>
          ) : null}

          {activeTab === 2 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Masking rules</p>
                  <h2>Default masking policy</h2>
                </div>
              </div>
              <div className={styles.formGrid}>
                <label className={styles.formField}>
                  <span>Default masking policy</span>
                  <select
                    className={styles.select}
                    onChange={(event) => setSecurityDraft((current) => ({ ...current, defaultMaskingPolicyKey: event.target.value }))}
                    value={securityDraft.defaultMaskingPolicyKey}
                  >
                    {manifest.maskingPolicies.map((policy) => (
                      <option key={policy.id} value={policy.key}>
                        {policy.name}
                      </option>
                    ))}
                  </select>
                </label>
              </div>
              <div className={styles.actionsRow}>
                <button className={styles.primaryButton} disabled={actor.role !== "SUPER_ADMIN"} onClick={() => void handleSecuritySave()} type="button">
                  Save masking rules
                </button>
              </div>
            </>
          ) : null}
        </section>
      </div>
    );
  }

  function renderAuditWorkspace() {
    return (
      <div className={styles.workspaceGrid}>
        <section className={styles.panelWide}>
          {activeTab === 0 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Version history</p>
                  <h2>Publish trail</h2>
                </div>
              </div>
              <div className={styles.listStack}>
                {bootstrap.versions.map((version) => (
                  <article className={styles.auditRow} key={version.id}>
                    <div>
                      <strong>v{version.versionNumber}</strong>
                      <p>
                        {version.status} · {version.manifestPath ?? "No manifest path"} · {version.gitCommitSha ?? "No git commit"}
                      </p>
                    </div>
                    <button
                      className={styles.ghostButton}
                      onClick={() =>
                        startTransition(() => {
                          void executeAction(
                            async () => {
                              await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/versions/${version.id}/rollback`, {
                                method: "POST",
                              });
                            },
                            `Rolled back to v${version.versionNumber}.`,
                          );
                        })
                      }
                      type="button"
                    >
                      Roll back
                    </button>
                  </article>
                ))}
              </div>
            </>
          ) : null}

          {activeTab === 1 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Audit events</p>
                  <h3>Administrative activity</h3>
                </div>
              </div>
              <div className={styles.tableWrap}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>Action</th>
                      <th>Resource</th>
                      <th>Actor</th>
                      <th>Time</th>
                    </tr>
                  </thead>
                  <tbody>
                    {bootstrap.auditEvents.map((event) => (
                      <tr key={event.id}>
                        <td>{event.summary}</td>
                        <td>
                          {event.resourceType}:{event.resourceId}
                        </td>
                        <td>{event.actorEmail ?? "system"}</td>
                        <td>{formatPlatformDateTime(event.createdAt)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          ) : null}
        </section>
      </div>
    );
  }

  return (
    <div className={styles.studioShell}>
      <aside className={sidebarCollapsed ? `${styles.studioSidebar} ${styles.sidebarCollapsed}` : styles.studioSidebar}>
        <div className={styles.sidebarToggleSlot}>
          <button
            aria-label={sidebarCollapsed ? "Expand sidebar" : "Collapse sidebar"}
            className={styles.sidebarToggleButton}
            onClick={() => setSidebarCollapsed((current) => !current)}
            type="button"
          >
            <svg className={styles.sidebarToggleIcon} fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.75" viewBox="0 0 12 12">
              <path d="M8 1L3 6l5 5" />
            </svg>
          </button>
        </div>
        <div className={styles.brandBlock}>
          <div className={styles.brandRow}>
            <div className={styles.brandMark}>TA</div>
            <div className={styles.brandCopy}>
              <p className={styles.sidebarTitle}>{bootstrap.tenant.name}</p>
              <p className={styles.sidebarSub}>Adaptive platform studio</p>
            </div>
          </div>
          <div className={styles.sidebarModePill}>Control plane · draft authoring</div>
        </div>
        <div className={styles.sidebarSection}>
          <p className={styles.sidebarLabel}>Tenant</p>
          <div className={styles.sidebarMeta}>
            <span>Environment</span>
            <strong>{bootstrap.environment.name}</strong>
          </div>
          <div className={styles.sidebarMeta}>
            <span>Role</span>
            <strong>{actor.role}</strong>
          </div>
          <div className={styles.sidebarMeta}>
            <span>Active version</span>
            <strong>{publishedLabel}</strong>
          </div>
          {bootstrap.session.memberships.length > 1 ? (
            <label className={styles.formField}>
              <span>Switch tenant</span>
              <select
                className={styles.select}
                onChange={(event) => {
                  window.location.href = `/platform/${event.target.value}`;
                }}
                value={bootstrap.tenant.slug}
              >
                {bootstrap.session.memberships.map((membership) => (
                  <option key={membership.tenantId} value={membership.tenantSlug}>
                    {membership.tenantName}
                  </option>
                ))}
              </select>
            </label>
          ) : null}
        </div>
        <div className={`${styles.sidebarSection} ${styles.sidebarNavSection}`}>
          <p className={styles.sidebarLabel}>Workspaces</p>
          <nav className={styles.navStack}>
            {WORKSPACES.map((entry) => (
              <button
                className={entry.key === workspace ? styles.activeNavItem : styles.navItem}
                key={entry.key}
                onClick={() => { setWorkspace(entry.key); setActiveTab(0); }}
                type="button"
              >
                <span className={styles.navIcon}>{entry.code}</span>
                <span className={styles.navCopy}>
                  <span>{entry.label}</span>
                  <small>{entry.note}</small>
                </span>
                <span className={styles.navTooltip}>{entry.label}</span>
              </button>
            ))}
          </nav>
        </div>
        <div className={styles.sidebarSection}>
          <p className={styles.sidebarLabel}>Publish</p>
          <div className={styles.sidebarPanel}>
            Draft definitions stay private until publish. Publishing freezes a runtime version and writes a manifest back into code.
          </div>
        </div>
        <div className={styles.sidebarSection}>
          <p className={styles.sidebarLabel}>Actions</p>
          <div className={styles.sidebarActions}>
            <button className={styles.primaryButton} disabled={isPending} onClick={() => startTransition(() => void handlePublish())} type="button">
              {isPending ? "Working..." : "Publish draft"}
            </button>
            <Link className={styles.secondaryLink} href={`/platform/preview/${bootstrap.tenant.slug}/${pageDraft.route || ""}`} target="_blank">
              Draft preview
            </Link>
            <Link className={styles.secondaryLink} href={`/platform/runtime/${bootstrap.tenant.slug}`} target="_blank">
              Open runtime
            </Link>
          </div>
        </div>
        <div className={styles.sidebarFooter}>
          Internal use only.
          <br />
          Objects, pages, workflows, agents, and models publish through one versioned manifest.
        </div>
      </aside>

      <main className={styles.studioMain}>
        <nav className={styles.breadcrumb}>
          <span>Studio</span>
          <span className={styles.breadcrumbSep}>/</span>
          <span>{activeWorkspace.label}</span>
          <span className={styles.breadcrumbSep}>/</span>
          <span className={styles.breadcrumbActive}>{WORKSPACE_TABS[workspace][activeTab]}</span>
        </nav>

        <header className={styles.studioHeader}>
          <div className={styles.headerLead}>
            <p className={styles.eyebrow}>TeacherActive adaptive platform</p>
            <h1>{activeWorkspace.label}</h1>
            <p className={styles.headerCopy}>{activeWorkspace.note}. Publish turns these draft definitions into immutable runtime manifests written back into code.</p>
          </div>
          <div className={styles.headerRail}>
            <div className={styles.headerStats}>
              <article className={styles.metricCard}>
                <span>Objects</span>
                <strong>{manifest.objects.length}</strong>
              </article>
              <article className={styles.metricCard}>
                <span>Pages</span>
                <strong>{manifest.pages.length}</strong>
              </article>
              <article className={styles.metricCard}>
                <span>Version</span>
                <strong>{bootstrap.activeVersion ? `v${bootstrap.activeVersion.versionNumber}` : "draft"}</strong>
              </article>
            </div>
            <div className={styles.headerNote}>
              <span>{bootstrap.environment.name}</span>
              <strong>{bootstrap.tenant.slug}</strong>
              <p>Reuse the existing shell, but author everything here as metadata first and runtime surfaces second.</p>
            </div>
          </div>
        </header>

        {message ? <div className={styles.successBanner}>{message}</div> : null}
        {error ? <div className={styles.errorBanner}>{error}</div> : null}

        {publishPreview ? (
          <section className={styles.publishPreviewStrip}>
            <div>
              <p className={styles.cardEyebrow}>Publish preview</p>
              <h3>Next activation will create v{publishPreview.nextVersionNumber}</h3>
            </div>
            <div className={styles.previewMetrics}>
              <div className={styles.previewMetric}>
                <span>Workflows</span>
                <strong>
                  +{publishPreview.summary.workflows.added.length} / ~{publishPreview.summary.workflows.updated.length} / -
                  {publishPreview.summary.workflows.removed.length}
                </strong>
              </div>
              <div className={styles.previewMetric}>
                <span>Agents</span>
                <strong>
                  +{publishPreview.summary.agents.added.length} / ~{publishPreview.summary.agents.updated.length} / -
                  {publishPreview.summary.agents.removed.length}
                </strong>
              </div>
              <div className={styles.previewMetric}>
                <span>Pages</span>
                <strong>
                  +{publishPreview.summary.pages.added.length} / ~{publishPreview.summary.pages.updated.length} / -
                  {publishPreview.summary.pages.removed.length}
                </strong>
              </div>
              <div className={styles.previewMetric}>
                <span>Layouts</span>
                <strong>
                  +{publishPreview.summary.layouts.added.length} / ~{publishPreview.summary.layouts.updated.length} / -
                  {publishPreview.summary.layouts.removed.length}
                </strong>
              </div>
            </div>
            <div className={styles.previewImpactStack}>
              {publishPreview.routeImpacts.slice(0, 4).map((impact) => (
                <div className={styles.previewImpactRow} key={`${impact.pageKey}-${impact.route}`}>
                  <strong>{impact.status.toUpperCase()}</strong>
                  <span>
                    {impact.pageKey} {"->"} /{impact.route}
                  </span>
                </div>
              ))}
              {publishPreview.pageImpacts.slice(0, 4).map((impact) => (
                <div className={styles.previewImpactRow} key={`${impact.pageKey}-${impact.status}`}>
                  <strong>{impact.status.toUpperCase()}</strong>
                  <span>
                    {impact.title} · {impact.sectionCount} sections · {impact.componentCount} components
                  </span>
                </div>
              ))}
            </div>
          </section>
        ) : null}

        {renderTabBar()}

        {workspace === "data-model" ? renderDataModelWorkspace() : null}
        {workspace === "pages" ? renderPagesWorkspace() : null}
        {workspace === "navigation" ? renderNavigationWorkspace() : null}
        {workspace === "workflows" ? renderWorkflowsWorkspace() : null}
        {workspace === "agents" ? renderAgentsWorkspace() : null}
        {workspace === "models" ? renderModelsWorkspace() : null}
        {workspace === "security" ? renderSecurityWorkspace() : null}
        {workspace === "audit" ? renderAuditWorkspace() : null}
      </main>
    </div>
  );
}
