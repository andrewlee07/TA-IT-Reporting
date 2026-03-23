"use client";

import Link from "next/link";
import { useCallback, useEffect, useEffectEvent, useMemo, useState, useTransition } from "react";

import { createThemeAccessibilityReport, createThemeSuggestions } from "@/lib/platform/branding-review";
import {
  createComponentFromPreset,
  createDefaultPlacement,
  createPageTemplateInstance,
  createSectionFromTemplate,
  normalizeLayoutComponentDefinition,
  normalizeLayoutSectionDefinition,
} from "@/lib/platform/designer";
import { formatPlatformDateTime } from "@/lib/platform/format";
import { getThemeCssVariables } from "@/lib/platform/theme";
import type {
  AgentDefinition,
  AgentRunRecord,
  AppShellDefinition,
  CostLedgerRecord,
  FormDefinition,
  LayoutComponentDefinition,
  LayoutDefinition,
  LayoutSectionDefinition,
  MenuItemDefinition,
  ModelProviderDefinition,
  NotificationCenterDefinition,
  NotificationDeliveryRecord,
  ObjectDefinition,
  PageDefinition,
  PlatformAlertRecord,
  PlatformApprovalTaskRecord,
  PlatformFormSubmissionRecord,
  PlatformAgentPreview,
  PlatformBootstrap,
  PlatformPublishPreview,
  PlatformRole,
  ThemeAccessibilityReport,
  ThemeTokenSuggestion,
  PlatformWorkflowRunRecord,
  SecurityPolicyDefinition,
  TenantBrandingDefinition,
  WorkflowDefinition,
  WorkflowEdgeDefinition,
  WorkflowNodeDefinition,
  WorkflowNodeType,
} from "@/lib/platform/types";

import styles from "./platform-shell.module.css";

type WorkspaceKey =
  | "data-model"
  | "pages"
  | "forms"
  | "branding"
  | "profiles"
  | "navigation"
  | "workflows"
  | "agents"
  | "control-tower"
  | "models"
  | "security"
  | "audit";

const WORKSPACES: Array<{ key: WorkspaceKey; label: string; note: string; code: string }> = [
  { key: "data-model", label: "Data Model", note: "Objects, fields, rules, formulas", code: "DM" },
  { key: "pages", label: "Pages", note: "Page definitions, layout sections, runtime components", code: "PG" },
  { key: "forms", label: "Forms", note: "Public forms, embedded intake flows, submissions", code: "FM" },
  { key: "branding", label: "Branding", note: "Tenant theme, logos, brand assets, shell identity", code: "BR" },
  { key: "profiles", label: "Profiles", note: "Profile pages, settings, admin view-as-user lens", code: "PF" },
  { key: "navigation", label: "Shell", note: "Menus, shell chrome, quick actions, notification routing", code: "SH" },
  { key: "workflows", label: "Workflows", note: "Visual graph metadata and execution scaffolding", code: "WF" },
  { key: "agents", label: "Agent Studio", note: "Prompt blocks, scope, tools, policy, handoffs", code: "AG" },
  { key: "control-tower", label: "Control Tower", note: "Agent runs, costs, alerts, deliveries, operator visibility", code: "CT" },
  { key: "models", label: "Models", note: "Provider registry and zero-retention controls", code: "ML" },
  { key: "security", label: "Security", note: "Masking defaults and protected-model policy", code: "SC" },
  { key: "audit", label: "Audit", note: "Publish history and admin activity", code: "AU" },
];

const WORKSPACE_TABS: Record<WorkspaceKey, string[]> = {
  "data-model": ["Objects", "Fields", "Validation"],
  pages: ["Designer", "Pages", "Templates"],
  forms: ["Builder", "Submissions"],
  branding: ["Theme", "Assets"],
  profiles: ["Experience", "View As"],
  navigation: ["Menu Items", "Shell", "Notifications", "Routes"],
  workflows: ["Definitions", "Runs"],
  agents: ["Definitions", "Prompts", "Scope", "Simulate"],
  "control-tower": ["Runs", "Costs", "Alerts", "Deliveries", "Approvals", "Dead Letters"],
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

function parseJsonPayloadSafely(value: string): Record<string, unknown> {
  const trimmed = value.trim();
  if (!trimmed) {
    return {};
  }

  const parsed = JSON.parse(trimmed) as unknown;
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new Error("Workflow payload must be a JSON object.");
  }

  return parsed as Record<string, unknown>;
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
    promptBlocks: [
      {
        id: createClientId("prompt-block"),
        label: "System role",
        kind: "system",
        content: "",
      },
    ],
    allowedToolIds: [],
    objectKeys: [],
    handoffWorkflowKeys: [],
    outputSchema: "",
    evalPolicy: {
      rubric: "Evaluate safety, masking, and operational usefulness.",
      samplePrompt: "Summarise the current workload.",
      passingScore: 0.8,
    },
    costBudgetUsd: 10,
    approvalPolicy: {
      required: false,
      approverRole: "BUILDER_ADMIN",
      notes: "",
    },
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

function createBlankForm(): FormDefinition {
  return {
    id: createClientId("form"),
    key: "",
    title: "",
    description: "",
    route: "",
    objectKey: "booking_request",
    deliveryMode: "public",
    submitLabel: "Submit",
    successMessage: "Submitted successfully.",
    saveAndResume: true,
    requireAuthentication: false,
    analyticsEnabled: true,
    fields: [
      {
        id: createClientId("form-field"),
        key: "full_name",
        label: "Full Name",
        type: "text",
        required: true,
        placeholder: "Jane Smith",
        validations: [
          {
            id: createClientId("val"),
            type: "required",
            message: "Full Name is required.",
          },
        ],
      },
    ],
    steps: [
      {
        id: createClientId("form-step"),
        key: "details",
        title: "Details",
        fieldKeys: ["full_name"],
      },
    ],
  };
}

function createFormDraftFromDefinition(formDefinition?: FormDefinition) {
  return formDefinition ? structuredClone(formDefinition) : createBlankForm();
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
  const initialForm = initialBootstrap.draftManifest.forms[0];
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
  const [selectedFormId, setSelectedFormId] = useState(initialForm?.id ?? "");
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
  const [formDraft, setFormDraft] = useState<FormDefinition>(createFormDraftFromDefinition(initialForm));
  const [formSubmissions, setFormSubmissions] = useState<PlatformFormSubmissionRecord[]>([]);
  const [brandingDraft, setBrandingDraft] = useState<TenantBrandingDefinition>(initialBootstrap.draftManifest.branding);
  const [brandAssetKind, setBrandAssetKind] = useState<"logo" | "icon" | "brand_book" | "reference">("logo");
  const [brandAssetLabel, setBrandAssetLabel] = useState("Tenant logo");
  const [brandAssetFile, setBrandAssetFile] = useState<File | null>(null);
  const [profileDraft, setProfileDraft] = useState({
    pageTitle: initialBootstrap.draftManifest.profiles.pageTitle,
    visibleFieldKeys: initialBootstrap.draftManifest.profiles.visibleFieldKeys.join(", "),
    profilePageKey: initialBootstrap.draftManifest.profiles.profilePageKey ?? "",
    settingsPageKey: initialBootstrap.draftManifest.profiles.settingsPageKey ?? "",
  });
  const [appShellDraft, setAppShellDraft] = useState<AppShellDefinition>(initialBootstrap.draftManifest.appShell);
  const [notificationDraft, setNotificationDraft] = useState<NotificationCenterDefinition>(initialBootstrap.draftManifest.notifications);
  const [viewAsDraft, setViewAsDraft] = useState({
    role: initialBootstrap.viewAs?.role ?? ("USER" as PlatformRole),
    personaLabel: initialBootstrap.viewAs?.personaLabel ?? "Sample teacher",
    active: initialBootstrap.viewAs?.active ?? false,
  });
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
    description: "",
    icon: "dot",
    pageKey: "",
    group: "Workspace",
    groupKey: "",
    order: 0,
    highlight: false,
    visibleToRoles: ["SUPER_ADMIN", "BUILDER_ADMIN", "USER"] as PlatformRole[],
    badgeBindingKey: "",
  });
  const [workflowDraft, setWorkflowDraft] = useState<WorkflowDefinition>(createWorkflowDraftFromDefinition(initialWorkflow));
  const [workflowRuns, setWorkflowRuns] = useState<PlatformWorkflowRunRecord[]>([]);
  const [allWorkflowRuns, setAllWorkflowRuns] = useState<PlatformWorkflowRunRecord[]>([]);
  const [workflowEdgeDraft, setWorkflowEdgeDraft] = useState({ sourceId: "", targetId: "", label: "" });
  const [workflowTestPayload, setWorkflowTestPayload] = useState('{\n  "previewMode": true\n}');
  const [agentDraft, setAgentDraft] = useState<AgentDefinition>(createAgentDraftFromDefinition(initialAgent));
  const [agentPreview, setAgentPreview] = useState<PlatformAgentPreview | null>(null);
  const [agentEvalSummary, setAgentEvalSummary] = useState<string | null>(null);
  const [agentRuns, setAgentRuns] = useState<AgentRunRecord[]>([]);
  const [costLedger, setCostLedger] = useState<CostLedgerRecord[]>([]);
  const [platformAlerts, setPlatformAlerts] = useState<PlatformAlertRecord[]>([]);
  const [notificationDeliveries, setNotificationDeliveries] = useState<NotificationDeliveryRecord[]>([]);
  const [approvalTasks, setApprovalTasks] = useState<PlatformApprovalTaskRecord[]>([]);
  const [deadLetters, setDeadLetters] = useState<Array<{ id: string; type: string; reason: string; createdAt: string }>>([]);
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
  const selectedForm = manifest.forms.find((formDefinition) => formDefinition.id === selectedFormId) ?? manifest.forms[0];
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
  const brandingSuggestions = useMemo<ThemeTokenSuggestion[]>(() => createThemeSuggestions(brandingDraft), [brandingDraft]);
  const brandingAccessibility = useMemo<ThemeAccessibilityReport>(() => createThemeAccessibilityReport(brandingDraft), [brandingDraft]);
  const formSubmissionSummary = useMemo(() => {
    const total = formSubmissions.length;
    const submitted = formSubmissions.filter((submission) => submission.status === "submitted").length;
    const drafts = total - submitted;
    return {
      total,
      submitted,
      drafts,
    };
  }, [formSubmissions]);
  const formDiagnostics = useMemo(() => {
    const diagnostics: string[] = [];
    if (!formDraft.title.trim()) {
      diagnostics.push("Form title is missing.");
    }
    if (!formDraft.route.trim()) {
      diagnostics.push("Form route is missing.");
    }
    if (formDraft.fields.length === 0) {
      diagnostics.push("Add at least one field.");
    }
    if (formDraft.steps.length === 0) {
      diagnostics.push("Add at least one step.");
    }
    const fieldKeys = new Set(formDraft.fields.map((field) => field.key).filter(Boolean));
    const danglingStep = formDraft.steps.find((step) => step.fieldKeys.some((fieldKey) => !fieldKeys.has(fieldKey)));
    if (danglingStep) {
      diagnostics.push(`Step "${danglingStep.title}" references a field key that is not defined.`);
    }
    if (formDraft.requireAuthentication && formDraft.deliveryMode === "public") {
      diagnostics.push("Authenticated forms should not remain in public delivery mode.");
    }
    return diagnostics;
  }, [formDraft]);
  const pageDiagnostics = useMemo(() => {
    const diagnostics: string[] = [];
    if (!pageDraft.title.trim()) {
      diagnostics.push("Page title is missing.");
    }
    if (!pageDraft.route.trim()) {
      diagnostics.push("Route is missing.");
    }
    if (!layoutDraft?.sections.length) {
      diagnostics.push("Add at least one section.");
    }
    if (layoutDraft?.sections.some((section) => section.components.length === 0)) {
      diagnostics.push("Every section should contain at least one component.");
    }
    const hasMenu = sortedMenus.some((menu) => menu.pageKey === (pageDraft.key || selectedPage?.key));
    if (!hasMenu) {
      diagnostics.push("This page is not exposed in the runtime menu.");
    }
    return diagnostics;
  }, [layoutDraft, pageDraft.key, pageDraft.route, pageDraft.title, selectedPage?.key, sortedMenus]);

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

  const refreshFormSubmissions = useCallback(async (formKey: string): Promise<void> => {
    try {
      const payload = await fetchJson<{ submissions: PlatformFormSubmissionRecord[] }>(
        `/api/platform/tenants/${bootstrap.tenant.slug}/forms/${formKey}/submissions`,
      );
      setFormSubmissions(payload.submissions);
    } catch {
      setFormSubmissions([]);
    }
  }, [bootstrap.tenant.slug]);

  const refreshControlTower = useCallback(async (): Promise<void> => {
    try {
      const [workflowRunsPayload, runsPayload, costsPayload, alertsPayload, deliveriesPayload, approvalsPayload, deadLettersPayload] = await Promise.all([
        fetchJson<{ runs: PlatformWorkflowRunRecord[] }>(`/api/platform/tenants/${bootstrap.tenant.slug}/workflow-runs`),
        fetchJson<{ runs: AgentRunRecord[] }>(`/api/platform/tenants/${bootstrap.tenant.slug}/agent-runs`),
        fetchJson<{ entries: CostLedgerRecord[] }>(`/api/platform/tenants/${bootstrap.tenant.slug}/cost-ledger`),
        fetchJson<{ alerts: PlatformAlertRecord[] }>(`/api/platform/tenants/${bootstrap.tenant.slug}/alerts`),
        fetchJson<{ deliveries: NotificationDeliveryRecord[] }>(`/api/platform/tenants/${bootstrap.tenant.slug}/deliveries`),
        fetchJson<{ tasks: PlatformApprovalTaskRecord[] }>(`/api/platform/tenants/${bootstrap.tenant.slug}/approval-tasks`),
        fetchJson<{ deadLetters: Array<{ id: string; type: string; reason: string; createdAt: string }> }>(
          `/api/platform/tenants/${bootstrap.tenant.slug}/dead-letters`,
        ),
      ]);
      setAllWorkflowRuns(workflowRunsPayload.runs);
      setAgentRuns(runsPayload.runs);
      setCostLedger(costsPayload.entries);
      setPlatformAlerts(alertsPayload.alerts);
      setNotificationDeliveries(deliveriesPayload.deliveries);
      setApprovalTasks(approvalsPayload.tasks);
      setDeadLetters(deadLettersPayload.deadLetters);
    } catch {
      setAllWorkflowRuns([]);
      setAgentRuns([]);
      setCostLedger([]);
      setPlatformAlerts([]);
      setNotificationDeliveries([]);
      setApprovalTasks([]);
      setDeadLetters([]);
    }
  }, [bootstrap.tenant.slug]);

  useEffect(() => {
    if (workspace === "control-tower" || workspace === "agents") {
      void refreshControlTower();
    }
  }, [refreshControlTower, workspace]);

  async function refreshBootstrap(): Promise<void> {
    const payload = await fetchJson<PlatformBootstrap>(`/api/platform/tenants/${bootstrap.tenant.slug}/bootstrap`);
    setBootstrap(payload);
    const nextObject =
      payload.draftManifest.objects.find((objectDefinition) => objectDefinition.id === selectedObjectId) ?? payload.draftManifest.objects[0];
    const nextPage =
      payload.draftManifest.pages.find((pageDefinition) => pageDefinition.id === selectedPageId) ?? payload.draftManifest.pages[0];
    const nextForm =
      payload.draftManifest.forms.find((formDefinition) => formDefinition.id === selectedFormId) ?? payload.draftManifest.forms[0];
    const nextWorkflow =
      payload.draftManifest.workflows.find((workflowDefinition) => workflowDefinition.id === selectedWorkflowId) ?? payload.draftManifest.workflows[0];
    const nextAgent =
      payload.draftManifest.agents.find((agentDefinition) => agentDefinition.id === selectedAgentId) ?? payload.draftManifest.agents[0];

    setSelectedObjectId(nextObject?.id ?? "");
    setSelectedPageId(nextPage?.id ?? "");
    setSelectedFormId(nextForm?.id ?? "");
    setSelectedWorkflowId(nextWorkflow?.id ?? "");
    setSelectedWorkflowNodeId(nextWorkflow?.nodes[0]?.id ?? "");
    setSelectedAgentId(nextAgent?.id ?? "");
    setObjectDraft(createObjectDraftFromDefinition(nextObject));
    setPageDraft(createPageDraftFromDefinition(nextPage));
    setFormDraft(createFormDraftFromDefinition(nextForm));
    setWorkflowDraft(createWorkflowDraftFromDefinition(nextWorkflow));
    setWorkflowEdgeDraft({
      sourceId: nextWorkflow?.nodes[0]?.id ?? "",
      targetId: nextWorkflow?.nodes[1]?.id ?? nextWorkflow?.nodes[0]?.id ?? "",
      label: "",
    });
    setAgentDraft(createAgentDraftFromDefinition(nextAgent));
    setBrandingDraft(payload.draftManifest.branding);
    setAppShellDraft(payload.draftManifest.appShell);
    setNotificationDraft(payload.draftManifest.notifications);
    setProfileDraft({
      pageTitle: payload.draftManifest.profiles.pageTitle,
      visibleFieldKeys: payload.draftManifest.profiles.visibleFieldKeys.join(", "),
      profilePageKey: payload.draftManifest.profiles.profilePageKey ?? "",
      settingsPageKey: payload.draftManifest.profiles.settingsPageKey ?? "",
    });
    setViewAsDraft({
      role: payload.viewAs?.role ?? "USER",
      personaLabel: payload.viewAs?.personaLabel ?? "Sample teacher",
      active: payload.viewAs?.active ?? false,
    });
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
    if (nextForm) {
      const formPayload = await fetchJson<{ submissions: PlatformFormSubmissionRecord[] }>(
        `/api/platform/tenants/${bootstrap.tenant.slug}/forms/${nextForm.key}/submissions`,
      ).catch(() => ({ submissions: [] }));
      setFormSubmissions(formPayload.submissions);
    } else {
      setFormSubmissions([]);
    }
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
          description: "",
          icon: "dot",
          pageKey: "",
          group: "Workspace",
          groupKey: "",
          order: manifest.menus.length,
          highlight: false,
          visibleToRoles: ["SUPER_ADMIN", "BUILDER_ADMIN", "USER"],
          badgeBindingKey: "",
        });
      },
      `Saved menu item ${menuDraft.label}.`,
    );
  }

  async function handleAppShellSave(): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/app-shell`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify(appShellDraft),
        });
      },
      `Saved shell config for ${appShellDraft.productName}.`,
    );
  }

  async function handleNotificationsSave(): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/notifications`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify(notificationDraft),
        });
        await refreshControlTower();
      },
      "Saved notification center configuration.",
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

  async function handleWorkflowTest(workflowId: string): Promise<void> {
    try {
      setError(null);
      setMessage(null);
      const payloadDraft = parseJsonPayloadSafely(workflowTestPayload);
      const payload = await fetchJson<{ run: PlatformWorkflowRunRecord }>(
        `/api/platform/tenants/${bootstrap.tenant.slug}/workflows/${workflowId}/test`,
        {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            payload: payloadDraft,
          }),
        },
      );
      setWorkflowRuns((current) => [payload.run, ...current]);
      setMessage(`Ran draft test for ${workflowDraft.name}.`);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Workflow test failed.");
    }
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

  async function handleReplayWorkflowRun(runId: string): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/workflow-runs/${runId}/replay`, {
          method: "POST",
        });
        await refreshControlTower();
      },
      "Workflow run replayed.",
    );
  }

  async function handleResolveApprovalTask(taskId: string, resolution: "approved" | "rejected"): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/approval-tasks/${taskId}/resolve`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({ resolution }),
        });
        await refreshControlTower();
      },
      `Approval task ${resolution}.`,
    );
  }

  async function handleBrandingSave(): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/branding`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify(brandingDraft),
        });
      },
      `Saved branding theme ${brandingDraft.themeName}.`,
    );
  }

  function applyBrandingSuggestion(suggestion: ThemeTokenSuggestion): void {
    setBrandingDraft((current) => ({
      ...current,
      ...suggestion.tokens,
      mode: "review",
    }));
  }

  async function handleBrandAssetUpload(): Promise<void> {
    if (!brandAssetFile) {
      setError("Choose a file before uploading a brand asset.");
      return;
    }

    const formData = new FormData();
    formData.append("file", brandAssetFile);
    formData.append("kind", brandAssetKind);
    formData.append("label", brandAssetLabel);

    await executeAction(
      async () => {
        const response = await fetch(`/api/platform/tenants/${bootstrap.tenant.slug}/branding/assets`, {
          method: "POST",
          body: formData,
        });

        const payload = (await response.json()) as { error?: string };
        if (!response.ok) {
          throw new Error(payload.error ?? "Failed to upload brand asset.");
        }
        setBrandAssetFile(null);
      },
      `Uploaded ${brandAssetKind.replace(/_/g, " ")} asset.`,
    );
  }

  async function handleFormSave(): Promise<void> {
    await executeAction(
      async () => {
        const payload = await fetchJson<{ form: FormDefinition }>(`/api/platform/tenants/${bootstrap.tenant.slug}/forms`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify(formDraft),
        });
        setSelectedFormId(payload.form.id);
      },
      `Saved form ${formDraft.title}.`,
    );
  }

  async function handleProfileSave(): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/profiles`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            pageTitle: profileDraft.pageTitle,
            visibleFieldKeys: profileDraft.visibleFieldKeys
              .split(",")
              .map((value) => value.trim())
              .filter(Boolean),
            profilePageKey: profileDraft.profilePageKey || undefined,
            settingsPageKey: profileDraft.settingsPageKey || undefined,
          }),
        });
      },
      "Saved profile and settings configuration.",
    );
  }

  async function handleViewAsSave(active: boolean): Promise<void> {
    try {
      setError(null);
      setMessage(null);
      await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/view-as`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
        },
        body: JSON.stringify({
          active,
          role: viewAsDraft.role,
          personaLabel: viewAsDraft.personaLabel,
          actorEmail: actor.email,
        }),
      });
      await refreshBootstrap();
      setMessage(active ? `Viewing runtime as ${viewAsDraft.personaLabel}.` : "Exited view-as mode.");
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Failed to update view-as state.");
    }
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
      setAgentEvalSummary(null);
      await refreshBootstrap();
      setMessage(`Prepared masked preview for ${agentDraft.name}.`);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Failed to preview agent.");
    }
  }

  async function handleEvaluateAgent(agentId: string): Promise<void> {
    try {
      setError(null);
      setMessage(null);
      const payload = await fetchJson<{ evaluation: { score: number; summary: string } }>(
        `/api/platform/tenants/${bootstrap.tenant.slug}/agents/${agentId}/evals`,
        {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            objectKey: agentDraft.objectKeys[0] ?? manifest.objects[0]?.key,
            sampleSize: 3,
          }),
        },
      );
      setAgentEvalSummary(`${payload.evaluation.score}/100 — ${payload.evaluation.summary}`);
      setMessage(`Evaluated ${agentDraft.name}.`);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Failed to evaluate agent.");
    }
  }

  async function handleSimulateAgent(agentId: string): Promise<void> {
    try {
      setError(null);
      setMessage(null);
      const payload = await fetchJson<{
        preview: PlatformAgentPreview;
        run: AgentRunRecord;
        cost: CostLedgerRecord;
      }>(`/api/platform/tenants/${bootstrap.tenant.slug}/agents/${agentId}/simulate`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
        },
        body: JSON.stringify({
          prompt: agentDraft.evalPolicy?.samplePrompt ?? agentDraft.prompt,
          objectKey: agentDraft.objectKeys[0] ?? manifest.objects[0]?.key,
          sampleSize: 3,
        }),
      });
      setAgentPreview(payload.preview);
      setAgentRuns((current) => [payload.run, ...current]);
      setCostLedger((current) => [payload.cost, ...current]);
      await refreshControlTower();
      setMessage(`Simulated ${agentDraft.name} in Control Tower.`);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Failed to simulate agent.");
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

  function updateFormField(
    fieldId: string,
    updater: (field: FormDefinition["fields"][number]) => FormDefinition["fields"][number],
  ): void {
    setFormDraft((current) => ({
      ...current,
      fields: current.fields.map((field) => (field.id === fieldId ? updater(field) : field)),
    }));
  }

  function removeFormField(fieldId: string): void {
    setFormDraft((current) => ({
      ...current,
      fields: current.fields.filter((field) => field.id !== fieldId),
      steps: current.steps.map((step) => ({
        ...step,
        fieldKeys: step.fieldKeys.filter((fieldKey) => current.fields.find((field) => field.id === fieldId)?.key !== fieldKey),
      })),
    }));
  }

  function addFormValidationRule(fieldId: string): void {
    updateFormField(fieldId, (field) => ({
      ...field,
      validations: [
        ...field.validations,
        {
          id: createClientId("form-validation"),
          type: "required",
          message: `${field.label} is required.`,
        },
      ],
    }));
  }

  function updateAgentPromptBlock(
    blockId: string,
    updater: (block: AgentDefinition["promptBlocks"][number]) => AgentDefinition["promptBlocks"][number],
  ): void {
    setAgentDraft((current) => ({
      ...current,
      promptBlocks: current.promptBlocks.map((block) => (block.id === blockId ? updater(block) : block)),
    }));
  }

  function addAgentPromptBlock(): void {
    setAgentDraft((current) => ({
      ...current,
      promptBlocks: [
        ...current.promptBlocks,
        {
          id: createClientId("prompt-block"),
          label: `Block ${current.promptBlocks.length + 1}`,
          kind: "instruction",
          content: "",
        },
      ],
    }));
  }

  function removeAgentPromptBlock(blockId: string): void {
    setAgentDraft((current) => ({
      ...current,
      promptBlocks: current.promptBlocks.filter((block) => block.id !== blockId),
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
            <div className={styles.sidebarMeta}>
              <span>Diagnostics</span>
              <strong>{pageDiagnostics.length === 0 ? "Clean" : `${pageDiagnostics.length} issues`}</strong>
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

          <div className={styles.designerInspectorGroup}>
            <div className={styles.sectionHeader}>
              <div>
                <p className={styles.cardEyebrow}>Diagnostics</p>
                <h3>Page readiness</h3>
              </div>
            </div>
            {pageDiagnostics.length === 0 ? (
              <div className={styles.sidebarPanel}>This page is structurally ready for preview and publish.</div>
            ) : (
              <div className={styles.listStack}>
                {pageDiagnostics.map((diagnostic) => (
                  <div className={styles.errorBanner} key={diagnostic}>
                    {diagnostic}
                  </div>
                ))}
              </div>
            )}
          </div>
        </section>
      </div>
    );
  }

  function renderFormsWorkspace() {
    return (
      <div className={styles.workspaceGrid}>
        <section className={styles.panel}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>Forms</p>
              <h2>Form registry</h2>
            </div>
            <button className={styles.secondaryButton} onClick={() => { setSelectedFormId(""); setFormDraft(createBlankForm()); setFormSubmissions([]); }} type="button">
              New form
            </button>
          </div>
          <div className={styles.listStack}>
            {manifest.forms.map((form) => (
              <button
                className={form.id === selectedForm?.id ? styles.activeListItem : styles.listItem}
                key={form.id}
                onClick={() => {
                  setSelectedFormId(form.id);
                  setFormDraft(createFormDraftFromDefinition(form));
                  void refreshFormSubmissions(form.key);
                }}
                type="button"
              >
                <span>{form.title}</span>
                <small>/{form.route} · {form.deliveryMode}</small>
              </button>
            ))}
          </div>
        </section>

        <section className={styles.panelWide}>
          {activeTab === 0 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Form builder</p>
                  <h2>{formDraft.title || "Create form"}</h2>
                </div>
                <div className={styles.inlineList}>
                  {formDraft.route ? (
                    <Link className={styles.secondaryLink} href={`/platform/preview/${bootstrap.tenant.slug}/forms/${formDraft.route}`} target="_blank">
                      Open draft preview
                    </Link>
                  ) : null}
                  {formDraft.route ? (
                    <Link className={styles.secondaryLink} href={`/platform/forms/${bootstrap.tenant.slug}/${formDraft.route}`} target="_blank">
                      Open live form
                    </Link>
                  ) : null}
                  <button className={styles.primaryButton} onClick={() => void handleFormSave()} type="button">
                    Save form
                  </button>
                </div>
              </div>
              <div className={styles.formGrid}>
                <label className={styles.formField}>
                  <span>Title</span>
                  <input className={styles.input} onChange={(event) => setFormDraft((current) => ({ ...current, title: event.target.value }))} value={formDraft.title} />
                </label>
                <label className={styles.formField}>
                  <span>Key</span>
                  <input className={styles.input} onChange={(event) => setFormDraft((current) => ({ ...current, key: event.target.value }))} value={formDraft.key} />
                </label>
                <label className={styles.formField}>
                  <span>Route</span>
                  <input className={styles.input} onChange={(event) => setFormDraft((current) => ({ ...current, route: event.target.value }))} value={formDraft.route} />
                </label>
                <label className={styles.formField}>
                  <span>Delivery mode</span>
                  <select className={styles.select} onChange={(event) => setFormDraft((current) => ({ ...current, deliveryMode: event.target.value as FormDefinition["deliveryMode"] }))} value={formDraft.deliveryMode}>
                    <option value="public">Public</option>
                    <option value="embedded">Embedded</option>
                    <option value="authenticated">Authenticated</option>
                  </select>
                </label>
                <label className={styles.formField}>
                  <span>Object</span>
                  <select className={styles.select} onChange={(event) => setFormDraft((current) => ({ ...current, objectKey: event.target.value || undefined }))} value={formDraft.objectKey ?? ""}>
                    <option value="">No object</option>
                    {manifest.objects.map((objectDefinition) => (
                      <option key={objectDefinition.id} value={objectDefinition.key}>
                        {objectDefinition.label}
                      </option>
                    ))}
                  </select>
                </label>
                <label className={styles.formField}>
                  <span>Submit label</span>
                  <input className={styles.input} onChange={(event) => setFormDraft((current) => ({ ...current, submitLabel: event.target.value }))} value={formDraft.submitLabel} />
                </label>
                <label className={styles.formFieldSpan}>
                  <span>Success message</span>
                  <textarea className={styles.textarea} onChange={(event) => setFormDraft((current) => ({ ...current, successMessage: event.target.value }))} value={formDraft.successMessage} />
                </label>
                <label className={styles.formFieldSpan}>
                  <span>Description</span>
                  <textarea className={styles.textarea} onChange={(event) => setFormDraft((current) => ({ ...current, description: event.target.value }))} value={formDraft.description ?? ""} />
                </label>
              </div>
              <div className={styles.inlineList}>
                <label className={styles.checkboxField}>
                  <input checked={formDraft.saveAndResume} onChange={(event) => setFormDraft((current) => ({ ...current, saveAndResume: event.target.checked }))} type="checkbox" />
                  <span>Save and resume</span>
                </label>
                <label className={styles.checkboxField}>
                  <input checked={formDraft.requireAuthentication} onChange={(event) => setFormDraft((current) => ({ ...current, requireAuthentication: event.target.checked }))} type="checkbox" />
                  <span>Require authentication</span>
                </label>
                <label className={styles.checkboxField}>
                  <input checked={formDraft.analyticsEnabled} onChange={(event) => setFormDraft((current) => ({ ...current, analyticsEnabled: event.target.checked }))} type="checkbox" />
                  <span>Analytics enabled</span>
                </label>
              </div>
              <div className={formDiagnostics.length === 0 ? styles.sidebarPanel : styles.errorBanner}>
                {formDiagnostics.length === 0
                  ? "Form structure is ready for preview. Add branching, calculations, and validations to deepen the experience."
                  : formDiagnostics.join(" ")}
              </div>
              <div className={styles.subSection}>
                <div className={styles.sectionHeader}>
                  <div>
                    <p className={styles.cardEyebrow}>Fields</p>
                    <h3>Question model</h3>
                  </div>
                  <button
                    className={styles.secondaryButton}
                    onClick={() =>
                      setFormDraft((current) => ({
                        ...current,
                        fields: [
                          ...current.fields,
                          {
                            id: createClientId("form-field"),
                            key: `field_${current.fields.length + 1}`,
                            label: `Question ${current.fields.length + 1}`,
                            type: "text",
                            required: false,
                            validations: [],
                          },
                        ],
                      }))
                    }
                    type="button"
                  >
                    Add field
                  </button>
                </div>
                <div className={styles.listStack}>
                  {formDraft.fields.map((field) => (
                    <div className={styles.fieldCard} key={field.id}>
                      <div className={styles.sectionHeader}>
                        <div>
                          <p className={styles.cardEyebrow}>Field</p>
                          <h3>{field.label}</h3>
                        </div>
                        <button className={styles.ghostButtonDanger} onClick={() => removeFormField(field.id)} type="button">
                          Remove
                        </button>
                      </div>
                      <div className={styles.formGridTight}>
                        <label className={styles.formField}>
                          <span>Label</span>
                          <input className={styles.input} onChange={(event) => updateFormField(field.id, (current) => ({ ...current, label: event.target.value }))} value={field.label} />
                        </label>
                        <label className={styles.formField}>
                          <span>Key</span>
                          <input className={styles.input} onChange={(event) => updateFormField(field.id, (current) => ({ ...current, key: event.target.value }))} value={field.key} />
                        </label>
                        <label className={styles.formField}>
                          <span>Type</span>
                          <select className={styles.select} onChange={(event) => updateFormField(field.id, (current) => ({ ...current, type: event.target.value as FormDefinition["fields"][number]["type"] }))} value={field.type}>
                            <option value="text">Text</option>
                            <option value="long_text">Long text</option>
                            <option value="number">Number</option>
                            <option value="currency">Currency</option>
                            <option value="boolean">Boolean</option>
                            <option value="date">Date</option>
                            <option value="datetime">DateTime</option>
                            <option value="select">Select</option>
                          </select>
                        </label>
                        <label className={styles.formField}>
                          <span>Tooltip</span>
                          <input className={styles.input} onChange={(event) => updateFormField(field.id, (current) => ({ ...current, tooltip: event.target.value }))} value={field.tooltip ?? ""} />
                        </label>
                        <label className={styles.formField}>
                          <span>Placeholder</span>
                          <input className={styles.input} onChange={(event) => updateFormField(field.id, (current) => ({ ...current, placeholder: event.target.value }))} value={field.placeholder ?? ""} />
                        </label>
                        <label className={styles.formField}>
                          <span>Help text</span>
                          <input className={styles.input} onChange={(event) => updateFormField(field.id, (current) => ({ ...current, helpText: event.target.value }))} value={field.helpText ?? ""} />
                        </label>
                        <label className={styles.formField}>
                          <span>Options</span>
                          <input className={styles.input} onChange={(event) => updateFormField(field.id, (current) => ({ ...current, options: event.target.value.split(",").map((entry) => entry.trim()).filter(Boolean) }))} value={field.options?.join(", ") ?? ""} />
                        </label>
                        <label className={styles.formField}>
                          <span>Default value</span>
                          <input className={styles.input} onChange={(event) => updateFormField(field.id, (current) => ({ ...current, defaultValue: event.target.value }))} value={String(field.defaultValue ?? "")} />
                        </label>
                        <label className={styles.formFieldSpan}>
                          <span>Mandatory rule</span>
                          <input className={styles.input} onChange={(event) => updateFormField(field.id, (current) => ({ ...current, mandatoryRule: event.target.value ? { mode: "text", expression: event.target.value } : undefined }))} value={field.mandatoryRule?.expression ?? ""} />
                        </label>
                        <label className={styles.formFieldSpan}>
                          <span>Calculation expression</span>
                          <input className={styles.input} onChange={(event) => updateFormField(field.id, (current) => ({ ...current, calculation: event.target.value ? { id: current.calculation?.id ?? createClientId("form-calculation"), expression: event.target.value, outputType: current.type === "currency" ? "currency" : current.type === "number" ? "number" : "text" } : null }))} value={field.calculation?.expression ?? ""} />
                        </label>
                      </div>
                      <div className={styles.inlineList}>
                        <label className={styles.checkboxField}>
                          <input checked={field.required} onChange={(event) => updateFormField(field.id, (current) => ({ ...current, required: event.target.checked }))} type="checkbox" />
                          <span>Required</span>
                        </label>
                      </div>
                      <div className={styles.subSection}>
                        <div className={styles.sectionHeader}>
                          <div>
                            <p className={styles.cardEyebrow}>Validation</p>
                            <h3>{field.validations.length} rules</h3>
                          </div>
                          <button className={styles.secondaryButton} onClick={() => addFormValidationRule(field.id)} type="button">
                            Add rule
                          </button>
                        </div>
                        <div className={styles.listStack}>
                          {field.validations.length === 0 ? (
                            <div className={styles.sidebarPanel}>No explicit validation rules yet. Required, regex, and numeric bounds can all be configured here.</div>
                          ) : (
                            field.validations.map((validation) => (
                              <article className={styles.workflowCard} key={validation.id}>
                                <div className={styles.formGridTight}>
                                  <label className={styles.formField}>
                                    <span>Type</span>
                                    <select className={styles.select} onChange={(event) => updateFormField(field.id, (current) => ({ ...current, validations: current.validations.map((candidate) => candidate.id === validation.id ? { ...candidate, type: event.target.value as typeof candidate.type } : candidate) }))} value={validation.type}>
                                      <option value="required">Required</option>
                                      <option value="min">Min</option>
                                      <option value="max">Max</option>
                                      <option value="regex">Regex</option>
                                      <option value="unique">Unique</option>
                                    </select>
                                  </label>
                                  <label className={styles.formField}>
                                    <span>Value</span>
                                    <input className={styles.input} onChange={(event) => updateFormField(field.id, (current) => ({ ...current, validations: current.validations.map((candidate) => candidate.id === validation.id ? { ...candidate, value: event.target.value } : candidate) }))} value={String(validation.value ?? "")} />
                                  </label>
                                  <label className={styles.formFieldSpan}>
                                    <span>Message</span>
                                    <input className={styles.input} onChange={(event) => updateFormField(field.id, (current) => ({ ...current, validations: current.validations.map((candidate) => candidate.id === validation.id ? { ...candidate, message: event.target.value } : candidate) }))} value={validation.message} />
                                  </label>
                                </div>
                                <div className={styles.actionsRow}>
                                  <button className={styles.ghostButtonDanger} onClick={() => updateFormField(field.id, (current) => ({ ...current, validations: current.validations.filter((candidate) => candidate.id !== validation.id) }))} type="button">
                                    Remove rule
                                  </button>
                                </div>
                              </article>
                            ))
                          )}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
              <div className={styles.subSection}>
                <div className={styles.sectionHeader}>
                  <div>
                    <p className={styles.cardEyebrow}>Steps</p>
                    <h3>Journey and branching</h3>
                  </div>
                  <button
                    className={styles.secondaryButton}
                    onClick={() =>
                      setFormDraft((current) => ({
                        ...current,
                        steps: [
                          ...current.steps,
                          {
                            id: createClientId("form-step"),
                            key: `step_${current.steps.length + 1}`,
                            title: `Step ${current.steps.length + 1}`,
                            fieldKeys: [],
                          },
                        ],
                      }))
                    }
                    type="button"
                  >
                    Add step
                  </button>
                </div>
                <div className={styles.listStack}>
                  {formDraft.steps.map((step) => (
                    <article className={styles.workflowCard} key={step.id}>
                      <div className={styles.sectionHeader}>
                        <div>
                          <p className={styles.cardEyebrow}>Step</p>
                          <h3>{step.title}</h3>
                        </div>
                        <button className={styles.ghostButtonDanger} onClick={() => setFormDraft((current) => ({ ...current, steps: current.steps.filter((candidate) => candidate.id !== step.id) }))} type="button">
                          Remove
                        </button>
                      </div>
                      <div className={styles.formGrid}>
                        <label className={styles.formField}>
                          <span>Title</span>
                          <input className={styles.input} onChange={(event) => setFormDraft((current) => ({ ...current, steps: current.steps.map((candidate) => candidate.id === step.id ? { ...candidate, title: event.target.value } : candidate) }))} value={step.title} />
                        </label>
                        <label className={styles.formField}>
                          <span>Key</span>
                          <input className={styles.input} onChange={(event) => setFormDraft((current) => ({ ...current, steps: current.steps.map((candidate) => candidate.id === step.id ? { ...candidate, key: event.target.value } : candidate) }))} value={step.key} />
                        </label>
                        <label className={styles.formFieldSpan}>
                          <span>Field keys</span>
                          <input className={styles.input} onChange={(event) => setFormDraft((current) => ({ ...current, steps: current.steps.map((candidate) => candidate.id === step.id ? { ...candidate, fieldKeys: event.target.value.split(",").map((entry) => entry.trim()).filter(Boolean) } : candidate) }))} value={step.fieldKeys.join(", ")} />
                        </label>
                        <label className={styles.formFieldSpan}>
                          <span>Visibility rule</span>
                          <input className={styles.input} onChange={(event) => setFormDraft((current) => ({ ...current, steps: current.steps.map((candidate) => candidate.id === step.id ? { ...candidate, visibilityRule: event.target.value ? { expression: event.target.value } : undefined } : candidate) }))} value={step.visibilityRule?.expression ?? ""} />
                        </label>
                      </div>
                    </article>
                  ))}
                </div>
              </div>
            </>
          ) : null}

          {activeTab === 1 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Submissions</p>
                  <h2>{selectedForm ? selectedForm.title : "No form selected"}</h2>
                </div>
                {selectedForm ? (
                  <div className={styles.inlineList}>
                    <Link className={styles.secondaryLink} href={`/platform/preview/${bootstrap.tenant.slug}/forms/${selectedForm.route}`} target="_blank">
                      Open draft preview
                    </Link>
                    <Link className={styles.secondaryLink} href={`/platform/forms/${bootstrap.tenant.slug}/${selectedForm.route}`} target="_blank">
                      Open live form
                    </Link>
                  </div>
                ) : null}
              </div>
              <div className={styles.metricGrid}>
                <article className={styles.metricCard}>
                  <span>Total</span>
                  <strong>{formSubmissionSummary.total}</strong>
                </article>
                <article className={styles.metricCard}>
                  <span>Submitted</span>
                  <strong>{formSubmissionSummary.submitted}</strong>
                </article>
                <article className={styles.metricCard}>
                  <span>Drafts</span>
                  <strong>{formSubmissionSummary.drafts}</strong>
                </article>
              </div>
              {formSubmissions.length === 0 ? (
                <div className={styles.emptyState}>No submissions yet. Publish the tenant runtime and submit the form to inspect captured entries.</div>
              ) : (
                <div className={styles.listStack}>
                  {formSubmissions.map((submission) => (
                    <article className={styles.workflowCard} key={submission.id}>
                      <strong>{submission.formKey}</strong>
                      <p>{submission.status} · {formatPlatformDateTime(submission.createdAt)}</p>
                      <pre className={styles.previewCode}>{JSON.stringify(submission.data, null, 2)}</pre>
                    </article>
                  ))}
                </div>
              )}
            </>
          ) : null}
        </section>
      </div>
    );
  }

  function renderBrandingWorkspace() {
    return (
      <div className={styles.workspaceGrid}>
        <section className={styles.panel}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>Brand system</p>
              <h2>Tenant identity</h2>
            </div>
          </div>
          <div className={styles.sidebarPanel}>
            Upload logos and brand books, then tune the tenant theme tokens that flow into runtime pages, forms, and previews.
          </div>
          <div className={styles.sidebarMeta}>
            <span>Theme</span>
            <strong>{brandingDraft.themeName}</strong>
          </div>
          <div className={styles.sidebarMeta}>
            <span>Assets</span>
            <strong>{manifest.branding.assets.length}</strong>
          </div>
        </section>
        <section className={styles.panelWide}>
          {activeTab === 0 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Theme</p>
                  <h2>{brandingDraft.themeName}</h2>
                </div>
                <button className={styles.primaryButton} onClick={() => void handleBrandingSave()} type="button">
                  Save theme
                </button>
              </div>
              <div className={styles.formGrid}>
                <label className={styles.formField}>
                  <span>Theme name</span>
                  <input className={styles.input} onChange={(event) => setBrandingDraft((current) => ({ ...current, themeName: event.target.value }))} value={brandingDraft.themeName} />
                </label>
                <label className={styles.formField}>
                  <span>Font family</span>
                  <input className={styles.input} onChange={(event) => setBrandingDraft((current) => ({ ...current, fontFamily: event.target.value }))} value={brandingDraft.fontFamily} />
                </label>
                {[
                  ["primaryColor", "Primary"],
                  ["secondaryColor", "Secondary"],
                  ["accentColor", "Accent"],
                  ["surfaceColor", "Surface"],
                  ["textColor", "Text"],
                  ["pageBackground", "Page background"],
                ].map(([key, label]) => (
                  <label className={styles.formField} key={key}>
                    <span>{label}</span>
                    <input className={styles.input} onChange={(event) => setBrandingDraft((current) => ({ ...current, [key]: event.target.value }))} value={String(brandingDraft[key as keyof TenantBrandingDefinition] ?? "")} />
                  </label>
                ))}
                <label className={styles.formFieldSpan}>
                  <span>Notes</span>
                  <textarea className={styles.textarea} onChange={(event) => setBrandingDraft((current) => ({ ...current, notes: event.target.value }))} value={brandingDraft.notes ?? ""} />
                </label>
              </div>
              <div className={styles.scopeGrid}>
                <article className={styles.scopeCard}>
                  <div className={styles.sectionHeader}>
                    <div>
                      <p className={styles.cardEyebrow}>Review queue</p>
                      <h3>Suggested theme directions</h3>
                    </div>
                  </div>
                  <div className={styles.listStack}>
                    {brandingSuggestions.map((suggestion) => (
                      <article className={styles.workflowCard} key={suggestion.key}>
                        <strong>{suggestion.label}</strong>
                        <p>{suggestion.description}</p>
                        <div className={styles.inlineList}>
                          <span className={styles.colorChip} style={{ background: suggestion.tokens.primaryColor }} />
                          <span className={styles.colorChip} style={{ background: suggestion.tokens.secondaryColor }} />
                          <span className={styles.colorChip} style={{ background: suggestion.tokens.accentColor }} />
                        </div>
                        <button className={styles.secondaryButton} onClick={() => applyBrandingSuggestion(suggestion)} type="button">
                          Apply suggestion
                        </button>
                      </article>
                    ))}
                  </div>
                </article>
                <article className={styles.scopeCard}>
                  <div className={styles.sectionHeader}>
                    <div>
                      <p className={styles.cardEyebrow}>Accessibility</p>
                      <h3>{brandingAccessibility.score}% pass rate</h3>
                    </div>
                  </div>
                  <div className={styles.listStack}>
                    {brandingAccessibility.checks.map((check) => (
                      <article className={styles.auditRow} key={check.key}>
                        <div>
                          <strong>{check.label}</strong>
                          <p>
                            {check.ratio}:1 against {check.requiredRatio}:1
                          </p>
                        </div>
                        <span className={check.passed ? styles.successPill : styles.warningPill}>{check.passed ? "Pass" : "Fix"}</span>
                      </article>
                    ))}
                  </div>
                  {brandingAccessibility.recommendations.length ? (
                    <div className={styles.sidebarPanel}>
                      {brandingAccessibility.recommendations.join(" ")}
                    </div>
                  ) : (
                    <div className={styles.sidebarPanel}>Theme tokens are currently passing the core contrast checks used by shell, forms, and runtime pages.</div>
                  )}
                </article>
                <article className={styles.scopeCard}>
                  <div className={styles.sectionHeader}>
                    <div>
                      <p className={styles.cardEyebrow}>Live preview</p>
                      <h3>Shell and form posture</h3>
                    </div>
                  </div>
                  <div className={styles.previewShell} style={getThemeCssVariables({ ...manifest, branding: brandingDraft })}>
                    <div className={styles.previewBanner}>
                      <div>
                        <p className={styles.cardEyebrow}>{brandingDraft.themeName}</p>
                        <h2>{bootstrap.tenant.name}</h2>
                        <p className={styles.helperCopy}>Brand tokens now preview shell chrome, runtime pages, and public forms before publish.</p>
                      </div>
                      <div className={styles.inlineList}>
                        <span className={styles.inlineTag}>{brandingDraft.fontFamily}</span>
                        <span className={styles.inlineTag}>{manifest.menus.length} nav items</span>
                      </div>
                    </div>
                    <div className={styles.tileGrid}>
                      {sortedMenus.slice(0, 4).map((menu) => (
                        <article className={styles.metricCard} key={menu.id}>
                          <span>{menu.group}</span>
                          <strong>{menu.label}</strong>
                          <p className={styles.metricMeta}>/{manifest.pages.find((page) => page.key === menu.pageKey)?.route ?? menu.pageKey}</p>
                        </article>
                      ))}
                    </div>
                  </div>
                </article>
              </div>
            </>
          ) : null}

          {activeTab === 1 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Assets</p>
                  <h2>Brand uploads</h2>
                </div>
                <button className={styles.primaryButton} onClick={() => void handleBrandAssetUpload()} type="button">
                  Upload asset
                </button>
              </div>
              <div className={styles.formGrid}>
                <label className={styles.formField}>
                  <span>Kind</span>
                  <select className={styles.select} onChange={(event) => setBrandAssetKind(event.target.value as typeof brandAssetKind)} value={brandAssetKind}>
                    <option value="logo">Logo</option>
                    <option value="icon">Icon</option>
                    <option value="brand_book">Brand book</option>
                    <option value="reference">Reference</option>
                  </select>
                </label>
                <label className={styles.formField}>
                  <span>Label</span>
                  <input className={styles.input} onChange={(event) => setBrandAssetLabel(event.target.value)} value={brandAssetLabel} />
                </label>
                <label className={styles.formFieldSpan}>
                  <span>File</span>
                  <input className={styles.input} onChange={(event) => setBrandAssetFile(event.target.files?.[0] ?? null)} type="file" />
                </label>
              </div>
              <div className={styles.listStack}>
                {manifest.branding.assets.map((asset) => (
                  <article className={styles.auditRow} key={asset.id}>
                    <div>
                      <strong>{asset.label}</strong>
                      <p>{asset.kind} · {asset.fileName}</p>
                      <small>
                        {brandingDraft.logoAssetId === asset.id ? "Logo" : null}
                        {brandingDraft.iconAssetId === asset.id ? `${brandingDraft.logoAssetId === asset.id ? " · " : ""}Icon` : null}
                        {brandingDraft.brandBookAssetId === asset.id ? `${brandingDraft.logoAssetId === asset.id || brandingDraft.iconAssetId === asset.id ? " · " : ""}Brand book` : null}
                      </small>
                    </div>
                    <div className={styles.inlineList}>
                      <button className={styles.ghostButton} onClick={() => setBrandingDraft((current) => ({ ...current, logoAssetId: asset.id }))} type="button">
                        Logo
                      </button>
                      <button className={styles.ghostButton} onClick={() => setBrandingDraft((current) => ({ ...current, iconAssetId: asset.id }))} type="button">
                        Icon
                      </button>
                      <button className={styles.ghostButton} onClick={() => setBrandingDraft((current) => ({ ...current, brandBookAssetId: asset.id }))} type="button">
                        Brand book
                      </button>
                      <Link className={styles.secondaryLink} href={asset.url} target="_blank">
                        Open
                      </Link>
                    </div>
                  </article>
                ))}
              </div>
            </>
          ) : null}
        </section>
      </div>
    );
  }

  function renderProfilesWorkspace() {
    return (
      <div className={styles.workspaceGrid}>
        <section className={styles.panel}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>Profiles</p>
              <h2>Identity surfaces</h2>
            </div>
          </div>
          <div className={styles.sidebarMeta}>
            <span>Profile page</span>
            <strong>{profileDraft.profilePageKey || "Unassigned"}</strong>
          </div>
          <div className={styles.sidebarMeta}>
            <span>Settings page</span>
            <strong>{profileDraft.settingsPageKey || "Unassigned"}</strong>
          </div>
          <div className={styles.sidebarMeta}>
            <span>View-as</span>
            <strong>{viewAsDraft.active ? `Active · ${viewAsDraft.personaLabel}` : "Inactive"}</strong>
          </div>
        </section>
        <section className={styles.panelWide}>
          {activeTab === 0 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Profile configuration</p>
                  <h2>{profileDraft.pageTitle}</h2>
                </div>
                <button className={styles.primaryButton} onClick={() => void handleProfileSave()} type="button">
                  Save profile config
                </button>
              </div>
              <div className={styles.formGrid}>
                <label className={styles.formField}>
                  <span>Page title</span>
                  <input className={styles.input} onChange={(event) => setProfileDraft((current) => ({ ...current, pageTitle: event.target.value }))} value={profileDraft.pageTitle} />
                </label>
                <label className={styles.formField}>
                  <span>Profile page key</span>
                  <input className={styles.input} onChange={(event) => setProfileDraft((current) => ({ ...current, profilePageKey: event.target.value }))} value={profileDraft.profilePageKey} />
                </label>
                <label className={styles.formField}>
                  <span>Settings page key</span>
                  <input className={styles.input} onChange={(event) => setProfileDraft((current) => ({ ...current, settingsPageKey: event.target.value }))} value={profileDraft.settingsPageKey} />
                </label>
                <label className={styles.formFieldSpan}>
                  <span>Visible field keys</span>
                  <input className={styles.input} onChange={(event) => setProfileDraft((current) => ({ ...current, visibleFieldKeys: event.target.value }))} value={profileDraft.visibleFieldKeys} />
                </label>
              </div>
              <div className={styles.inlineList}>
                <Link className={styles.secondaryLink} href={`/platform/admin-preview/${bootstrap.tenant.slug}/${profileDraft.profilePageKey || "profile"}`} target="_blank">
                  Open admin preview
                </Link>
                <Link className={styles.secondaryLink} href={`/platform/runtime/${bootstrap.tenant.slug}/${profileDraft.profilePageKey || "profile"}`} target="_blank">
                  Open runtime profile
                </Link>
              </div>
            </>
          ) : null}

          {activeTab === 1 ? (
            <>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>View as user</p>
                  <h2>Admin impersonation lens</h2>
                </div>
              </div>
              <div className={styles.formGrid}>
                <label className={styles.formField}>
                  <span>Role</span>
                  <select className={styles.select} onChange={(event) => setViewAsDraft((current) => ({ ...current, role: event.target.value as PlatformRole }))} value={viewAsDraft.role}>
                    <option value="USER">User</option>
                    <option value="BUILDER_ADMIN">Builder admin</option>
                    <option value="SUPER_ADMIN">Super admin</option>
                  </select>
                </label>
                <label className={styles.formField}>
                  <span>Persona label</span>
                  <input className={styles.input} onChange={(event) => setViewAsDraft((current) => ({ ...current, personaLabel: event.target.value }))} value={viewAsDraft.personaLabel} />
                </label>
              </div>
              <div className={styles.actionsRow}>
                <div className={styles.inlineList}>
                  <button className={styles.primaryButton} onClick={() => void handleViewAsSave(true)} type="button">
                    Enable view-as
                  </button>
                  <button className={styles.secondaryButton} onClick={() => void handleViewAsSave(false)} type="button">
                    Clear view-as
                  </button>
                </div>
                <div className={styles.inlineList}>
                  <Link className={styles.secondaryLink} href={`/platform/admin-preview/${bootstrap.tenant.slug}`} target="_blank">
                    Open admin preview
                  </Link>
                  <Link className={styles.secondaryLink} href={`/platform/runtime/${bootstrap.tenant.slug}`} target="_blank">
                    Open runtime
                  </Link>
                </div>
              </div>
            </>
          ) : null}
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
                      {menu.visibleToRoles?.length ? <small>{menu.visibleToRoles.join(" · ")}</small> : null}
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
                <label className={styles.formFieldSpan}>
                  <span>Description</span>
                  <input className={styles.input} onChange={(event) => setMenuDraft((current) => ({ ...current, description: event.target.value }))} value={menuDraft.description} />
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
                <label className={styles.formField}>
                  <span>Group key</span>
                  <input className={styles.input} onChange={(event) => setMenuDraft((current) => ({ ...current, groupKey: event.target.value }))} value={menuDraft.groupKey} />
                </label>
                <label className={styles.formField}>
                  <span>Badge binding</span>
                  <select className={styles.select} onChange={(event) => setMenuDraft((current) => ({ ...current, badgeBindingKey: event.target.value }))} value={menuDraft.badgeBindingKey}>
                    <option value="">No badge</option>
                    {appShellDraft.badgeBindings.map((binding) => (
                      <option key={binding.key} value={binding.key}>
                        {binding.label}
                      </option>
                    ))}
                  </select>
                </label>
              </div>
              <div className={styles.inlineList}>
                <label className={styles.checkboxField}>
                  <input checked={menuDraft.highlight} onChange={(event) => setMenuDraft((current) => ({ ...current, highlight: event.target.checked }))} type="checkbox" />
                  <span>Highlight in runtime nav</span>
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
          <>
            <section className={styles.panel}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>App shell</p>
                  <h2>Runtime chrome</h2>
                </div>
              </div>
              <div className={styles.formGrid}>
                <label className={styles.formField}>
                  <span>Product name</span>
                  <input className={styles.input} onChange={(event) => setAppShellDraft((current) => ({ ...current, productName: event.target.value }))} value={appShellDraft.productName} />
                </label>
                <label className={styles.formField}>
                  <span>Navigation mode</span>
                  <select className={styles.select} onChange={(event) => setAppShellDraft((current) => ({ ...current, menuStyle: event.target.value as AppShellDefinition["menuStyle"], navigationMode: event.target.value as AppShellDefinition["navigationMode"] }))} value={appShellDraft.navigationMode}>
                    <option value="sidebar">Sidebar</option>
                    <option value="topbar">Topbar</option>
                  </select>
                </label>
                <label className={styles.formFieldSpan}>
                  <span>Tag line</span>
                  <input className={styles.input} onChange={(event) => setAppShellDraft((current) => ({ ...current, tagLine: event.target.value }))} value={appShellDraft.tagLine ?? ""} />
                </label>
                <label className={styles.formField}>
                  <span>Support email</span>
                  <input className={styles.input} onChange={(event) => setAppShellDraft((current) => ({ ...current, supportEmail: event.target.value }))} value={appShellDraft.supportEmail ?? ""} />
                </label>
                <label className={styles.formField}>
                  <span>Default landing page</span>
                  <select className={styles.select} onChange={(event) => setAppShellDraft((current) => ({ ...current, defaultLandingPageKey: event.target.value }))} value={appShellDraft.defaultLandingPageKey ?? ""}>
                    <option value="">Choose page</option>
                    {manifest.pages.map((page) => (
                      <option key={page.id} value={page.key}>
                        {page.title}
                      </option>
                    ))}
                  </select>
                </label>
              </div>
              <div className={styles.actionsRow}>
                <button className={styles.primaryButton} onClick={() => void handleAppShellSave()} type="button">
                  Save shell
                </button>
              </div>
            </section>
            <section className={styles.panelWide}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Quick actions and groups</p>
                  <h2>Shell controls</h2>
                </div>
              </div>
              <div className={styles.scopeGrid}>
                <article className={styles.scopeCard}>
                  <p className={styles.cardEyebrow}>Menu groups</p>
                  <div className={styles.listStack}>
                    {appShellDraft.menuGroups.map((group, index) => (
                      <div className={styles.fieldCard} key={group.key}>
                        <div className={styles.formGrid}>
                          <label className={styles.formField}>
                            <span>Label</span>
                            <input className={styles.input} onChange={(event) => setAppShellDraft((current) => ({ ...current, menuGroups: current.menuGroups.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, label: event.target.value } : candidate) }))} value={group.label} />
                          </label>
                          <label className={styles.formField}>
                            <span>Order</span>
                            <input className={styles.input} onChange={(event) => setAppShellDraft((current) => ({ ...current, menuGroups: current.menuGroups.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, order: Number(event.target.value) || 0 } : candidate) }))} type="number" value={group.order} />
                          </label>
                        </div>
                      </div>
                    ))}
                  </div>
                </article>
                <article className={styles.scopeCard}>
                  <p className={styles.cardEyebrow}>Quick actions</p>
                  <div className={styles.listStack}>
                    {appShellDraft.quickActions.map((action, index) => (
                      <div className={styles.fieldCard} key={action.key}>
                        <div className={styles.formGrid}>
                          <label className={styles.formField}>
                            <span>Label</span>
                            <input className={styles.input} onChange={(event) => setAppShellDraft((current) => ({ ...current, quickActions: current.quickActions.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, label: event.target.value } : candidate) }))} value={action.label} />
                          </label>
                          <label className={styles.formField}>
                            <span>Page</span>
                            <select className={styles.select} onChange={(event) => setAppShellDraft((current) => ({ ...current, quickActions: current.quickActions.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, pageKey: event.target.value || undefined } : candidate) }))} value={action.pageKey ?? ""}>
                              <option value="">No page</option>
                              {manifest.pages.map((page) => (
                                <option key={page.id} value={page.key}>
                                  {page.title}
                                </option>
                              ))}
                            </select>
                          </label>
                        </div>
                      </div>
                    ))}
                  </div>
                </article>
              </div>
            </section>
          </>
        ) : null}

        {activeTab === 2 ? (
          <>
            <section className={styles.panel}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Notification channels</p>
                  <h2>Delivery routing</h2>
                </div>
              </div>
              <div className={styles.listStack}>
                {notificationDraft.channels.map((channel, index) => (
                  <div className={styles.fieldCard} key={channel.key}>
                    <div className={styles.formGrid}>
                      <label className={styles.formField}>
                        <span>Name</span>
                        <input className={styles.input} onChange={(event) => setNotificationDraft((current) => ({ ...current, channels: current.channels.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, name: event.target.value } : candidate) }))} value={channel.name} />
                      </label>
                      <label className={styles.formField}>
                        <span>Destination</span>
                        <input className={styles.input} onChange={(event) => setNotificationDraft((current) => ({ ...current, channels: current.channels.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, destination: event.target.value } : candidate) }))} value={channel.destination ?? ""} />
                      </label>
                    </div>
                    <label className={styles.checkboxField}>
                      <input checked={channel.enabled} onChange={(event) => setNotificationDraft((current) => ({ ...current, channels: current.channels.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, enabled: event.target.checked } : candidate) }))} type="checkbox" />
                      <span>{channel.kind}</span>
                    </label>
                  </div>
                ))}
              </div>
            </section>
            <section className={styles.panelWide}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Notification rules</p>
                  <h2>Routing policies</h2>
                </div>
              </div>
              <div className={styles.listStack}>
                {notificationDraft.rules.map((rule, index) => (
                  <div className={styles.workflowCard} key={rule.key}>
                    <div className={styles.formGrid}>
                      <label className={styles.formField}>
                        <span>Name</span>
                        <input className={styles.input} onChange={(event) => setNotificationDraft((current) => ({ ...current, rules: current.rules.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, name: event.target.value } : candidate) }))} value={rule.name} />
                      </label>
                      <label className={styles.formField}>
                        <span>Event type</span>
                        <input className={styles.input} onChange={(event) => setNotificationDraft((current) => ({ ...current, rules: current.rules.map((candidate, candidateIndex) => candidateIndex === index ? { ...candidate, eventType: event.target.value } : candidate) }))} value={rule.eventType} />
                      </label>
                    </div>
                    <div className={styles.inlineList}>
                      <span className={styles.inlineTag}>{rule.severity}</span>
                      <span className={styles.inlineTag}>{rule.channelKeys.join(", ")}</span>
                    </div>
                  </div>
                ))}
              </div>
              <div className={styles.actionsRow}>
                <button className={styles.primaryButton} onClick={() => void handleNotificationsSave()} type="button">
                  Save notifications
                </button>
              </div>
            </section>
          </>
        ) : null}

        {activeTab === 3 ? (
          <section className={styles.panelWide}>
            <div className={styles.emptyState}>Route ownership stays metadata-driven through pages. This tab is reserved for route diagnostics and impact review.</div>
          </section>
        ) : null}
      </div>
    );
  }

  function renderWorkflowsWorkspace() {
    const workflowDiagnostics: string[] = [];
    if (!workflowDraft.name.trim()) {
      workflowDiagnostics.push("Workflow name is required.");
    }
    if (!workflowDraft.triggers.length) {
      workflowDiagnostics.push("Add at least one trigger.");
    }
    if (!workflowDraft.nodes.length) {
      workflowDiagnostics.push("Add at least one node.");
    }
    if (workflowDraft.nodes.length > 1 && workflowDraft.edges.length === 0) {
      workflowDiagnostics.push("Connect nodes with at least one edge.");
    }

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
                  <div className={styles.inlineList}>
                    <button className={styles.secondaryButton} onClick={() => void handleWorkflowTest(selectedWorkflow.id)} type="button">
                      Run draft test
                    </button>
                    <button className={styles.secondaryButton} onClick={() => void handleQueueWorkflowRun(selectedWorkflow.id)} type="button">
                      Queue run
                    </button>
                  </div>
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
              <div className={styles.scopeGrid}>
                <article className={styles.scopeCard}>
                  <div className={styles.sectionHeader}>
                    <div>
                      <p className={styles.cardEyebrow}>Triggers</p>
                      <h3>Launch conditions</h3>
                    </div>
                    <button
                      className={styles.secondaryButton}
                      onClick={() =>
                        setWorkflowDraft((current) => ({
                          ...current,
                          triggers: [
                            ...current.triggers,
                            {
                              id: createClientId("trigger"),
                              type: "record_created",
                              label: `Trigger ${current.triggers.length + 1}`,
                              config: {
                                objectKey: current.objectKey || manifest.objects[0]?.key || "booking_request",
                              },
                            },
                          ],
                        }))
                      }
                      type="button"
                    >
                      Add trigger
                    </button>
                  </div>
                  <div className={styles.listStack}>
                    {workflowDraft.triggers.map((trigger) => (
                      <article className={styles.workflowCard} key={trigger.id}>
                        <div className={styles.formGrid}>
                          <label className={styles.formField}>
                            <span>Label</span>
                            <input className={styles.input} onChange={(event) => setWorkflowDraft((current) => ({ ...current, triggers: current.triggers.map((candidate) => candidate.id === trigger.id ? { ...candidate, label: event.target.value } : candidate) }))} value={trigger.label} />
                          </label>
                          <label className={styles.formField}>
                            <span>Type</span>
                            <select className={styles.select} onChange={(event) => setWorkflowDraft((current) => ({ ...current, triggers: current.triggers.map((candidate) => candidate.id === trigger.id ? { ...candidate, type: event.target.value as WorkflowDefinition["triggers"][number]["type"] } : candidate) }))} value={trigger.type}>
                              <option value="manual">Manual</option>
                              <option value="record_created">Record created</option>
                              <option value="record_updated">Record updated</option>
                            </select>
                          </label>
                          {"objectKey" in trigger.config ? (
                            <label className={styles.formField}>
                              <span>Object</span>
                              <select className={styles.select} onChange={(event) => setWorkflowDraft((current) => ({ ...current, triggers: current.triggers.map((candidate) => candidate.id === trigger.id ? { ...candidate, config: { objectKey: event.target.value } } : candidate) }))} value={trigger.config.objectKey}>
                                {manifest.objects.map((objectDefinition) => (
                                  <option key={objectDefinition.id} value={objectDefinition.key}>
                                    {objectDefinition.label}
                                  </option>
                                ))}
                              </select>
                            </label>
                          ) : (
                            <label className={styles.formFieldSpan}>
                              <span>Notes</span>
                              <input className={styles.input} onChange={(event) => setWorkflowDraft((current) => ({ ...current, triggers: current.triggers.map((candidate) => candidate.id === trigger.id ? { ...candidate, config: { notes: event.target.value } } : candidate) }))} value={"notes" in trigger.config ? trigger.config.notes ?? "" : ""} />
                            </label>
                          )}
                        </div>
                      </article>
                    ))}
                  </div>
                </article>
                <article className={styles.scopeCard}>
                  <div className={styles.sectionHeader}>
                    <div>
                      <p className={styles.cardEyebrow}>Draft test</p>
                      <h3>Sample payload</h3>
                    </div>
                  </div>
                  <textarea className={styles.textareaTall} onChange={(event) => setWorkflowTestPayload(event.target.value)} value={workflowTestPayload} />
                  <div className={styles.sidebarPanel}>
                    {workflowDiagnostics.length === 0 ? "Graph passes the current structural checks." : workflowDiagnostics.join(" ")}
                  </div>
                </article>
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
                  <div className={styles.inlineList}>
                    <button className={styles.secondaryButton} onClick={() => void handleWorkflowTest(selectedWorkflow.id)} type="button">
                      Run draft test
                    </button>
                    <button className={styles.primaryButton} onClick={() => void handleQueueWorkflowRun(selectedWorkflow.id)} type="button">
                      Queue run
                    </button>
                  </div>
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
                <label className={styles.formField}>
                  <span>Output schema</span>
                  <textarea className={styles.textarea} onChange={(event) => setAgentDraft((current) => ({ ...current, outputSchema: event.target.value }))} value={agentDraft.outputSchema ?? ""} />
                </label>
                <label className={styles.formField}>
                  <span>Budget (USD)</span>
                  <input className={styles.input} min="0" onChange={(event) => setAgentDraft((current) => ({ ...current, costBudgetUsd: Number(event.target.value) || 0 }))} step="0.01" type="number" value={agentDraft.costBudgetUsd ?? 0} />
                </label>
              </div>
              <div className={styles.inlineList}>
                <label className={styles.checkboxField}>
                  <input checked={agentDraft.zeroRetentionRequired} onChange={(event) => setAgentDraft((current) => ({ ...current, zeroRetentionRequired: event.target.checked }))} type="checkbox" />
                  <span>Zero retention required</span>
                </label>
                <label className={styles.checkboxField}>
                  <input checked={agentDraft.approvalPolicy?.required ?? false} onChange={(event) => setAgentDraft((current) => ({ ...current, approvalPolicy: { ...(current.approvalPolicy ?? { approverRole: "BUILDER_ADMIN", notes: "" }), required: event.target.checked } }))} type="checkbox" />
                  <span>Approval required</span>
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
                <div className={styles.inlineList}>
                  <button className={styles.secondaryButton} onClick={addAgentPromptBlock} type="button">
                    Add prompt block
                  </button>
                  <button className={styles.secondaryButton} disabled={!agentDraft.id} onClick={() => void handlePreviewAgent(agentDraft.id)} type="button">
                    Preview masked invocation
                  </button>
                  <button className={styles.secondaryButton} disabled={!agentDraft.id} onClick={() => void handleEvaluateAgent(agentDraft.id)} type="button">
                    Run eval
                  </button>
                </div>
              </div>
              <div className={styles.formGrid}>
                <label className={styles.formFieldSpan}>
                  <span>System prompt</span>
                  <textarea className={styles.textareaTall} onChange={(event) => setAgentDraft((current) => ({ ...current, prompt: event.target.value }))} value={agentDraft.prompt} />
                </label>
                <label className={styles.formField}>
                  <span>Eval sample prompt</span>
                  <input className={styles.input} onChange={(event) => setAgentDraft((current) => ({ ...current, evalPolicy: { ...(current.evalPolicy ?? { rubric: "", samplePrompt: "", passingScore: 0.8 }), samplePrompt: event.target.value } }))} value={agentDraft.evalPolicy?.samplePrompt ?? ""} />
                </label>
                <label className={styles.formField}>
                  <span>Passing score</span>
                  <input className={styles.input} max="1" min="0" onChange={(event) => setAgentDraft((current) => ({ ...current, evalPolicy: { ...(current.evalPolicy ?? { rubric: "", samplePrompt: "", passingScore: 0.8 }), passingScore: Number(event.target.value) || 0 } }))} step="0.05" type="number" value={agentDraft.evalPolicy?.passingScore ?? 0.8} />
                </label>
              </div>
              <div className={styles.listStack}>
                {agentDraft.promptBlocks.map((block) => (
                  <article className={styles.workflowCard} key={block.id}>
                    <div className={styles.sectionHeader}>
                      <div>
                        <p className={styles.cardEyebrow}>Prompt block</p>
                        <h3>{block.label}</h3>
                      </div>
                      <button className={styles.ghostButtonDanger} disabled={agentDraft.promptBlocks.length === 1} onClick={() => removeAgentPromptBlock(block.id)} type="button">
                        Remove
                      </button>
                    </div>
                    <div className={styles.formGridTight}>
                      <label className={styles.formField}>
                        <span>Label</span>
                        <input className={styles.input} onChange={(event) => updateAgentPromptBlock(block.id, (current) => ({ ...current, label: event.target.value }))} value={block.label} />
                      </label>
                      <label className={styles.formField}>
                        <span>Kind</span>
                        <select className={styles.select} onChange={(event) => updateAgentPromptBlock(block.id, (current) => ({ ...current, kind: event.target.value as typeof current.kind }))} value={block.kind}>
                          <option value="system">System</option>
                          <option value="policy">Policy</option>
                          <option value="instruction">Instruction</option>
                          <option value="example">Example</option>
                        </select>
                      </label>
                      <label className={styles.formFieldSpan}>
                        <span>Content</span>
                        <textarea className={styles.textarea} onChange={(event) => updateAgentPromptBlock(block.id, (current) => ({ ...current, content: event.target.value }))} value={block.content} />
                      </label>
                    </div>
                  </article>
                ))}
              </div>
              {agentEvalSummary ? <div className={styles.sidebarPanel}>{agentEvalSummary}</div> : null}
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

          {activeTab === 3 ? (
            <div className={styles.scopeGrid}>
              <article className={styles.scopeCard}>
                <div className={styles.sectionHeader}>
                  <div>
                    <p className={styles.cardEyebrow}>Control Tower</p>
                    <h3>Agent simulation</h3>
                  </div>
                  <button className={styles.primaryButton} disabled={!agentDraft.id} onClick={() => void handleSimulateAgent(agentDraft.id)} type="button">
                    Run simulation
                  </button>
                </div>
                <div className={styles.formGrid}>
                  <label className={styles.formFieldSpan}>
                    <span>Simulation prompt</span>
                    <textarea className={styles.textarea} onChange={(event) => setAgentDraft((current) => ({ ...current, evalPolicy: { ...(current.evalPolicy ?? { rubric: "", samplePrompt: "", passingScore: 0.8 }), samplePrompt: event.target.value } }))} value={agentDraft.evalPolicy?.samplePrompt ?? ""} />
                  </label>
                  <label className={styles.formField}>
                    <span>Handoff workflows</span>
                    <select className={styles.select} onChange={(event) => setAgentDraft((current) => ({ ...current, handoffWorkflowKeys: event.target.value ? [event.target.value] : [] }))} value={agentDraft.handoffWorkflowKeys[0] ?? ""}>
                      <option value="">No handoff</option>
                      {manifest.workflows.map((workflow) => (
                        <option key={workflow.id} value={workflow.key}>
                          {workflow.name}
                        </option>
                      ))}
                    </select>
                  </label>
                </div>
                <div className={styles.sidebarPanel}>
                  Latest budget: ${(agentDraft.costBudgetUsd ?? 0).toFixed(2)} · Handoffs: {agentDraft.handoffWorkflowKeys.length || 0}
                </div>
              </article>
              <article className={styles.scopeCard}>
                <div className={styles.sectionHeader}>
                  <div>
                    <p className={styles.cardEyebrow}>Recent runs</p>
                    <h3>Operator visibility</h3>
                  </div>
                </div>
                {agentRuns.filter((run) => run.agentId === agentDraft.id).length === 0 ? (
                  <div className={styles.emptyState}>No agent simulations yet.</div>
                ) : (
                  <div className={styles.listStack}>
                    {agentRuns.filter((run) => run.agentId === agentDraft.id).slice(0, 5).map((run) => (
                      <article className={styles.auditRow} key={run.id}>
                        <div>
                          <strong>{run.status}</strong>
                          <p>
                            ${run.costUsd.toFixed(4)} · {run.tokensIn}/{run.tokensOut} tokens
                          </p>
                        </div>
                        <span>{formatPlatformDateTime(run.createdAt)}</span>
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

  function renderControlTowerWorkspace() {
    return (
      <div className={styles.workspaceGrid}>
        {activeTab === 0 ? (
          <>
            <section className={styles.panel}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Workflow runs</p>
                  <h2>Queue, retries, and replay</h2>
                </div>
                <button className={styles.secondaryButton} onClick={() => void refreshControlTower()} type="button">
                  Refresh
                </button>
              </div>
              <div className={styles.listStack}>
                {allWorkflowRuns.length === 0 ? (
                  <div className={styles.emptyState}>No workflow runs yet.</div>
                ) : (
                  allWorkflowRuns.slice(0, 12).map((run) => (
                    <article className={styles.auditRow} key={run.id}>
                      <div>
                        <strong>{run.workflowKey}</strong>
                        <p>
                          {run.status} · {run.logs.length} log events
                        </p>
                      </div>
                      <div className={styles.inlineList}>
                        <span>{formatPlatformDateTime(run.createdAt)}</span>
                        <button className={styles.ghostButton} onClick={() => void handleReplayWorkflowRun(run.id)} type="button">
                          Replay
                        </button>
                      </div>
                    </article>
                  ))
                )}
              </div>
            </section>
            <section className={styles.panelWide}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Agent runs</p>
                  <h2>Provider-backed simulations</h2>
                </div>
              </div>
              <div className={styles.listStack}>
                {agentRuns.slice(0, 6).map((run) => (
                  <article className={styles.workflowCard} key={`detail-${run.id}`}>
                    <strong>{run.agentKey}</strong>
                    <p>{run.logs[0]?.message ? String(run.logs[0].message) : "Simulation completed."}</p>
                    <div className={styles.inlineList}>
                      <span className={styles.inlineTag}>{run.modelProviderKey}</span>
                      <span className={styles.inlineTag}>{run.tokensIn + run.tokensOut} tokens</span>
                    </div>
                  </article>
                ))}
              </div>
            </section>
          </>
        ) : null}

        {activeTab === 1 ? (
          <section className={styles.panelWide}>
            <div className={styles.sectionHeader}>
              <div>
                <p className={styles.cardEyebrow}>Cost ledger</p>
                <h2>Usage and spend</h2>
              </div>
            </div>
            <div className={styles.tableWrap}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th>Category</th>
                    <th>Reference</th>
                    <th>Provider</th>
                    <th>Amount</th>
                    <th>Created</th>
                  </tr>
                </thead>
                <tbody>
                  {costLedger.length > 0 ? (
                    costLedger.slice(0, 12).map((entry) => (
                      <tr key={entry.id}>
                        <td>{entry.category}</td>
                        <td>{entry.referenceId}</td>
                        <td>{entry.providerKey ?? "—"}</td>
                        <td>${entry.amountUsd.toFixed(4)}</td>
                        <td>{formatPlatformDateTime(entry.createdAt)}</td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={5}>
                        <div className={styles.emptyState}>No cost data yet.</div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>
        ) : null}

        {activeTab === 2 ? (
          <section className={styles.panelWide}>
            <div className={styles.sectionHeader}>
              <div>
                <p className={styles.cardEyebrow}>Alerts</p>
                <h2>Budget, delivery, and runtime warnings</h2>
              </div>
            </div>
            <div className={styles.listStack}>
              {platformAlerts.length === 0 ? (
                <div className={styles.emptyState}>No alerts triggered.</div>
              ) : (
                platformAlerts.slice(0, 12).map((alert) => (
                  <article className={styles.auditRow} key={alert.id}>
                    <div>
                      <strong>{alert.title}</strong>
                      <p>{alert.summary}</p>
                    </div>
                    <div className={styles.inlineList}>
                      <span className={styles.inlineTag}>{alert.category}</span>
                      <span className={styles.inlineTag}>{alert.severity}</span>
                    </div>
                  </article>
                ))
              )}
            </div>
          </section>
        ) : null}

        {activeTab === 3 ? (
          <section className={styles.panelWide}>
            <div className={styles.sectionHeader}>
              <div>
                <p className={styles.cardEyebrow}>Deliveries</p>
                <h2>Notification history</h2>
              </div>
            </div>
            <div className={styles.tableWrap}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th>Rule</th>
                    <th>Channel</th>
                    <th>Status</th>
                    <th>Severity</th>
                    <th>Delivered</th>
                  </tr>
                </thead>
                <tbody>
                  {notificationDeliveries.length > 0 ? (
                    notificationDeliveries.slice(0, 12).map((delivery) => (
                      <tr key={delivery.id}>
                        <td>{delivery.ruleKey}</td>
                        <td>{delivery.channelKey}</td>
                        <td>{delivery.status}</td>
                        <td>{delivery.severity}</td>
                        <td>{formatPlatformDateTime(delivery.deliveredAt ?? delivery.createdAt)}</td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={5}>
                        <div className={styles.emptyState}>No deliveries yet.</div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>
        ) : null}

        {activeTab === 4 ? (
          <section className={styles.panelWide}>
            <div className={styles.sectionHeader}>
              <div>
                <p className={styles.cardEyebrow}>Approvals</p>
                <h2>Pending human checkpoints</h2>
              </div>
            </div>
            <div className={styles.listStack}>
              {approvalTasks.length === 0 ? (
                <div className={styles.emptyState}>No approval tasks are pending.</div>
              ) : (
                approvalTasks.map((task) => (
                  <article className={styles.workflowCard} key={task.id}>
                    <strong>{task.nodeLabel}</strong>
                    <p>
                      {task.workflowKey} · {task.status} · {task.approverRole}
                    </p>
                    {task.instructions ? <div className={styles.sidebarPanel}>{task.instructions}</div> : null}
                    <div className={styles.inlineList}>
                      <button className={styles.secondaryButton} disabled={task.status !== "pending"} onClick={() => void handleResolveApprovalTask(task.id, "approved")} type="button">
                        Approve
                      </button>
                      <button className={styles.ghostButtonDanger} disabled={task.status !== "pending"} onClick={() => void handleResolveApprovalTask(task.id, "rejected")} type="button">
                        Reject
                      </button>
                    </div>
                  </article>
                ))
              )}
            </div>
          </section>
        ) : null}

        {activeTab === 5 ? (
          <section className={styles.panelWide}>
            <div className={styles.sectionHeader}>
              <div>
                <p className={styles.cardEyebrow}>Dead letters</p>
                <h2>Escalated runtime failures</h2>
              </div>
            </div>
            <div className={styles.listStack}>
              {deadLetters.length === 0 ? (
                <div className={styles.emptyState}>No dead letters recorded.</div>
              ) : (
                deadLetters.map((entry) => (
                  <article className={styles.auditRow} key={entry.id}>
                    <div>
                      <strong>{entry.type}</strong>
                      <p>{entry.reason}</p>
                    </div>
                    <span>{formatPlatformDateTime(entry.createdAt)}</span>
                  </article>
                ))
              )}
            </div>
          </section>
        ) : null}
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
            <Link className={styles.secondaryLink} href={`/platform/admin-preview/${bootstrap.tenant.slug}`} target="_blank">
              Admin preview
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
          <button type="button" className={styles.breadcrumbLink} onClick={() => { setActiveTab(0); window.scrollTo({ top: 0, behavior: "smooth" }); }}>Studio</button>
          <span className={styles.breadcrumbSep}>/</span>
          <button type="button" className={styles.breadcrumbLink} onClick={() => setActiveTab(0)}>{activeWorkspace.label}</button>
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

        {renderTabBar()}

        {workspace === "data-model" ? renderDataModelWorkspace() : null}
        {workspace === "pages" ? renderPagesWorkspace() : null}
        {workspace === "forms" ? renderFormsWorkspace() : null}
        {workspace === "branding" ? renderBrandingWorkspace() : null}
        {workspace === "profiles" ? renderProfilesWorkspace() : null}
        {workspace === "navigation" ? renderNavigationWorkspace() : null}
        {workspace === "workflows" ? renderWorkflowsWorkspace() : null}
        {workspace === "agents" ? renderAgentsWorkspace() : null}
        {workspace === "control-tower" ? renderControlTowerWorkspace() : null}
        {workspace === "models" ? renderModelsWorkspace() : null}
        {workspace === "security" ? renderSecurityWorkspace() : null}
        {workspace === "audit" ? renderAuditWorkspace() : null}

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
      </main>
    </div>
  );
}
