"use client";

import Link from "next/link";
import { useCallback, useEffect, useEffectEvent, useMemo, useRef, useState, useTransition } from "react";

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
  AgentRunDetail,
  AgentRunRecord,
  AppShellDefinition,
  CostLedgerRecord,
  DeadLetterRecord,
  FormDefinition,
  LayoutComponentDefinition,
  LayoutDefinition,
  LayoutSectionDefinition,
  MenuItemDefinition,
  ModelProviderDefinition,
  NotificationAttemptRecord,
  NotificationChannelHealth,
  NotificationCenterDefinition,
  NotificationDeliveryRecord,
  ObjectDefinition,
  PageDefinition,
  PageDiagnostic,
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
  WorkflowRunDetail,
  WorkflowDiagnostic,
  WorkflowDefinition,
  WorkflowEdgeDefinition,
  WorkflowNodeDefinition,
  WorkflowNodeType,
  WorkflowTemplateDefinition,
  WorkflowTestCaseDefinition,
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
  { key: "navigation", label: "Navigation & Shell", note: "Menus, shell chrome, quick actions, notification routing", code: "SH" },
  { key: "workflows", label: "Workflows", note: "Visual graph metadata and execution scaffolding", code: "WF" },
  { key: "agents", label: "Agent Studio", note: "Prompt blocks, scope, tools, policy, handoffs", code: "AG" },
  { key: "control-tower", label: "Control Tower", note: "Agent runs, costs, alerts, deliveries, operator visibility", code: "CT" },
  { key: "models", label: "Models", note: "Provider registry and zero-retention controls", code: "ML" },
  { key: "security", label: "Security", note: "Masking defaults and protected-model policy", code: "SC" },
  { key: "audit", label: "Audit", note: "Publish history and admin activity", code: "AU" },
];

const WORKSPACE_GROUPS: Array<{ label: string; key: string; workspaces: WorkspaceKey[] }> = [
  { label: "Build", key: "build", workspaces: ["data-model", "pages", "forms", "branding"] },
  { label: "Configure", key: "configure", workspaces: ["profiles", "navigation", "workflows", "agents"] },
  { label: "Operate", key: "operate", workspaces: ["control-tower", "models", "security", "audit"] },
];

const WORKSPACE_TABS: Record<WorkspaceKey, string[]> = {
  "data-model": ["Objects", "Fields", "Validation"],
  pages: ["Designer", "Pages", "Templates"],
  forms: ["Builder", "Submissions"],
  branding: ["Theme", "Assets"],
  profiles: ["Experience", "View As"],
  navigation: ["Menus", "App Chrome", "Notifications"],
  workflows: ["Definitions", "Runs"],
  agents: ["Agents", "Configure", "Test"],
  "control-tower": ["Execution", "Alerts & Approvals", "Messaging"],
  models: ["Providers", "Configuration"],
  security: ["Policies", "Roles & Access", "Data Masking"],
  audit: ["History", "Activity"],
};

type ToastKind = "success" | "error" | "info";

interface ToastAction {
  label: string;
  onClick: () => void;
  tone?: "default" | "danger";
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

interface InlineUndoState {
  key: string;
  expiresAt: number;
}

function tabKeyFromLabel(label: string): string {
  return label
    .trim()
    .toLowerCase()
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function getTabLabelFromKey(workspace: WorkspaceKey, key: string | null | undefined): number {
  if (!key) {
    return 0;
  }

  const index = WORKSPACE_TABS[workspace].findIndex((label) => tabKeyFromLabel(label) === key);
  return index >= 0 ? index : 0;
}

function isValidCssColor(value: string): boolean {
  const trimmed = value.trim();
  if (!trimmed) {
    return false;
  }

  if (typeof CSS !== "undefined" && typeof CSS.supports === "function") {
    return CSS.supports("color", trimmed);
  }

  return /^#([a-f0-9]{3}|[a-f0-9]{6}|[a-f0-9]{8})$/i.test(trimmed);
}

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

type ControlTowerSelectionType = "workflow-run" | "agent-run" | "delivery" | "alert" | "approval" | "dead-letter";

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
    case "subflow":
      return {
        id: createClientId("node"),
        type,
        label: "Subflow",
        config: {
          workflowKey: "booking_triage",
          description: "Invoke a published reusable workflow.",
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
  { type: "subflow", title: "Subflow", note: "Reuse a published workflow as a node." },
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
  const [activeTabByWorkspace, setActiveTabByWorkspace] = useState<Partial<Record<WorkspaceKey, number>>>({});
  const [toasts, setToasts] = useState<ToastRecord[]>([]);
  const [confirmDialog, setConfirmDialog] = useState<{ title: string; message: string; confirmLabel: string; onConfirm: () => void } | null>(null);
  const [commandPaletteOpen, setCommandPaletteOpen] = useState(false);
  const [commandPaletteQuery, setCommandPaletteQuery] = useState("");
  const [commandPaletteIndex, setCommandPaletteIndex] = useState(0);
  const [commandPaletteRecentIds, setCommandPaletteRecentIds] = useState<string[]>([]);
  const [shortcutsOpen, setShortcutsOpen] = useState(false);
  const [collapsedGroups, setCollapsedGroups] = useState<Set<string>>(new Set());
  const [controlTowerFiltersExpanded, setControlTowerFiltersExpanded] = useState(false);
  const [controlTowerListLimits, setControlTowerListLimits] = useState({ workflowRuns: 12, costs: 12, alerts: 12, deliveries: 12, approvals: 12, deadLetters: 12 });
  const [workspaceTransitionKey, setWorkspaceTransitionKey] = useState(0);
  const [controlTowerSelection, setControlTowerSelection] = useState<{ type: ControlTowerSelectionType; id: string } | null>(null);
  const [controlTowerDetail, setControlTowerDetail] = useState<Record<string, unknown> | null>(null);
  const [controlTowerFilters, setControlTowerFilters] = useState({
    status: "all",
    workflowKey: "",
    agentKey: "",
    severity: "all",
    fromDate: "",
    toDate: "",
  });
  const [scrolledPast, setScrolledPast] = useState(false);
  const [pendingInlineDelete, setPendingInlineDelete] = useState<InlineUndoState | null>(null);
  const [sectionActionMenuId, setSectionActionMenuId] = useState<string | null>(null);
  const [hasHydratedUrlState, setHasHydratedUrlState] = useState(false);
  const [isPublishPreviewLoading, setIsPublishPreviewLoading] = useState(false);
  const [isControlTowerLoading, setIsControlTowerLoading] = useState(false);
  const [isPending, startTransition] = useTransition();
  const mainRef = useRef<HTMLElement | null>(null);
  const previouslyFocusedElementRef = useRef<HTMLElement | null>(null);
  const confirmDialogRef = useRef<HTMLDivElement | null>(null);
  const confirmCancelButtonRef = useRef<HTMLButtonElement | null>(null);
  const toastTimersRef = useRef<Map<string, number>>(new Map());
  const toastsRef = useRef<ToastRecord[]>([]);
  const studioUrlRef = useRef<string | null>(null);
  const inlineDeleteTimerRef = useRef<number | null>(null);
  const inlineDeleteActionRef = useRef<(() => void | Promise<void>) | null>(null);
  const inlineDeleteIntervalRef = useRef<number | null>(null);

  const activeTab = activeTabByWorkspace[workspace] ?? 0;
  const setActiveTab = useCallback((index: number) => setActiveTabByWorkspace((prev) => ({ ...prev, [workspace]: index })), [workspace]);

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
    const duplicateFieldKeys = formDraft.fields
      .map((field) => field.key.trim())
      .filter(Boolean)
      .filter((key, index, values) => values.indexOf(key) !== index);
    if (duplicateFieldKeys.length > 0) {
      diagnostics.push(`Field keys must be unique. Duplicate keys: ${[...new Set(duplicateFieldKeys)].join(", ")}.`);
    }
    const emptyStep = formDraft.steps.find((step) => step.fieldKeys.length === 0);
    if (emptyStep) {
      diagnostics.push(`Step "${emptyStep.title}" has no fields assigned.`);
    }
    const unassignedField = formDraft.fields.find((field) => !formDraft.steps.some((step) => step.fieldKeys.includes(field.key)));
    if (unassignedField) {
      diagnostics.push(`Field "${unassignedField.label}" is not assigned to any step.`);
    }
    if (formDraft.requireAuthentication && formDraft.deliveryMode === "public") {
      diagnostics.push("Authenticated forms should not remain in public delivery mode.");
    }
    if (formDraft.saveAndResume && !formDraft.analyticsEnabled) {
      diagnostics.push("Consider enabling analytics when save-and-resume is active so drop-off can be measured.");
    }
    return diagnostics;
  }, [formDraft]);
  const pageDiagnostics = useMemo<PageDiagnostic[]>(() => {
    const diagnostics: PageDiagnostic[] = [];
    if (!pageDraft.title.trim()) {
      diagnostics.push({ id: "page-title", severity: "blocking", category: "content", message: "Page title is missing." });
    }
    if (!pageDraft.route.trim()) {
      diagnostics.push({ id: "page-route", severity: "blocking", category: "route", message: "Route is missing." });
    }
    if (!layoutDraft?.sections.length) {
      diagnostics.push({ id: "page-sections", severity: "blocking", category: "content", message: "Add at least one section." });
    }
    if (layoutDraft?.sections.some((section) => section.components.length === 0)) {
      diagnostics.push({
        id: "page-empty-section",
        severity: "warning",
        category: "content",
        message: "Every section should contain at least one component.",
      });
    }
    if (
      layoutDraft?.sections.some((section) =>
        section.components.some((component) => {
          const responsive = component.placement.responsive;
          return (
            (responsive.desktopSpan != null && responsive.desktopSpan > 12) ||
            (responsive.tabletSpan != null && responsive.tabletSpan > 12) ||
            (responsive.mobileSpan != null && responsive.mobileSpan > 12) ||
            component.placement.span > 12
          );
        }),
      )
    ) {
      diagnostics.push({
        id: "page-responsive-span",
        severity: "blocking",
        category: "responsive",
        message: "One or more components exceed the supported grid span for mobile, tablet, or desktop.",
      });
    }
    if (
      layoutDraft?.sections.some((section) =>
        section.components.some(
          (component) =>
            component.visibilityRule != null &&
            (!component.visibilityRule.expression || component.visibilityRule.expression.trim().length === 0),
        ),
      )
    ) {
      diagnostics.push({
        id: "page-visibility",
        severity: "warning",
        category: "visibility",
        message: "A component visibility rule is configured but empty.",
      });
    }
    const componentMissingBinding = layoutDraft?.sections
      .flatMap((section) => section.components)
      .find((component) => {
        if (component.kind === "record_table" || component.kind === "record_form" || component.kind === "related_records") {
          return !(component.binding?.objectKey || component.objectKey || component.binding?.relatedObjectKey || component.relatedObjectKey);
        }
        if (component.kind === "workflow_launcher") {
          return !(component.binding?.workflowKey || component.workflowKey);
        }
        if (component.kind === "agent_summary" || component.kind === "agent_panel") {
          return !(component.binding?.agentId || component.agentId);
        }
        return false;
      });
    if (componentMissingBinding) {
      diagnostics.push({
        id: "page-binding",
        severity: "blocking",
        category: "binding",
        message: `Component "${componentMissingBinding.title}" is missing its required runtime binding.`,
      });
    }
    const hasMenu = sortedMenus.some((menu) => menu.pageKey === (pageDraft.key || selectedPage?.key));
    if (!hasMenu && !pageDraft.isHome) {
      diagnostics.push({
        id: "page-menu",
        severity: "warning",
        category: "menu",
        message: "This page is not exposed in the runtime menu.",
      });
    }
    const duplicateRoute = manifest.pages.find(
      (page) => page.id !== (selectedPage?.id ?? pageDraft.id) && page.route.trim().toLowerCase() === pageDraft.route.trim().toLowerCase() && pageDraft.route.trim(),
    );
    if (duplicateRoute) {
      diagnostics.push({
        id: "page-duplicate-route",
        severity: "blocking",
        category: "route",
        message: `Route /${pageDraft.route} is already owned by ${duplicateRoute.title}.`,
      });
    }
    const routeConflict = publishPreview?.routeImpacts.find((impact) => impact.pageKey === pageDraft.key && impact.status === "updated");
    if (routeConflict) {
      diagnostics.push({
        id: "page-route-conflict",
        severity: "info",
        category: "route",
        message: `Publishing will change the live route to /${routeConflict.route}.`,
      });
    }
    return diagnostics;
  }, [layoutDraft, manifest.pages, pageDraft.id, pageDraft.isHome, pageDraft.key, pageDraft.route, pageDraft.title, publishPreview?.routeImpacts, selectedPage?.id, selectedPage?.key, sortedMenus]);
  const pageReadinessScore = useMemo(() => {
    return Math.max(
      0,
      100 -
        pageDiagnostics.filter((diagnostic) => diagnostic.severity === "blocking").length * 30 -
        pageDiagnostics.filter((diagnostic) => diagnostic.severity === "warning").length * 12 -
        pageDiagnostics.filter((diagnostic) => diagnostic.severity === "info").length * 4,
    );
  }, [pageDiagnostics]);
  const filteredWorkflowRuns = useMemo(() => {
    return allWorkflowRuns.filter((run) => {
      if (controlTowerFilters.status !== "all" && run.status !== controlTowerFilters.status) {
        return false;
      }
      if (controlTowerFilters.workflowKey && !run.workflowKey.toLowerCase().includes(controlTowerFilters.workflowKey.toLowerCase())) {
        return false;
      }
      if (controlTowerFilters.fromDate && run.createdAt.slice(0, 10) < controlTowerFilters.fromDate) {
        return false;
      }
      if (controlTowerFilters.toDate && run.createdAt.slice(0, 10) > controlTowerFilters.toDate) {
        return false;
      }
      return true;
    });
  }, [allWorkflowRuns, controlTowerFilters.fromDate, controlTowerFilters.status, controlTowerFilters.toDate, controlTowerFilters.workflowKey]);
  const filteredAgentRuns = useMemo(() => {
    return agentRuns.filter((run) => {
      if (controlTowerFilters.status !== "all" && run.status !== controlTowerFilters.status) {
        return false;
      }
      if (controlTowerFilters.agentKey && !run.agentKey.toLowerCase().includes(controlTowerFilters.agentKey.toLowerCase())) {
        return false;
      }
      if (controlTowerFilters.fromDate && run.createdAt.slice(0, 10) < controlTowerFilters.fromDate) {
        return false;
      }
      if (controlTowerFilters.toDate && run.createdAt.slice(0, 10) > controlTowerFilters.toDate) {
        return false;
      }
      return true;
    });
  }, [agentRuns, controlTowerFilters.agentKey, controlTowerFilters.fromDate, controlTowerFilters.status, controlTowerFilters.toDate]);
  const filteredAlerts = useMemo(() => {
    return platformAlerts.filter((alert) => {
      if (controlTowerFilters.severity !== "all" && alert.severity !== controlTowerFilters.severity) {
        return false;
      }
      if (controlTowerFilters.fromDate && alert.createdAt.slice(0, 10) < controlTowerFilters.fromDate) {
        return false;
      }
      if (controlTowerFilters.toDate && alert.createdAt.slice(0, 10) > controlTowerFilters.toDate) {
        return false;
      }
      return true;
    });
  }, [controlTowerFilters.fromDate, controlTowerFilters.severity, controlTowerFilters.toDate, platformAlerts]);
  const filteredDeliveries = useMemo(() => {
    return notificationDeliveries.filter((delivery) => {
      if (controlTowerFilters.status !== "all" && delivery.status !== controlTowerFilters.status) {
        return false;
      }
      if (controlTowerFilters.severity !== "all" && delivery.severity !== controlTowerFilters.severity) {
        return false;
      }
      if (controlTowerFilters.fromDate && delivery.createdAt.slice(0, 10) < controlTowerFilters.fromDate) {
        return false;
      }
      if (controlTowerFilters.toDate && delivery.createdAt.slice(0, 10) > controlTowerFilters.toDate) {
        return false;
      }
      return true;
    });
  }, [controlTowerFilters.fromDate, controlTowerFilters.severity, controlTowerFilters.status, controlTowerFilters.toDate, notificationDeliveries]);
  const activeFilterCount = [
    controlTowerFilters.status !== "all",
    controlTowerFilters.workflowKey !== "",
    controlTowerFilters.agentKey !== "",
    controlTowerFilters.severity !== "all",
    controlTowerFilters.fromDate !== "",
    controlTowerFilters.toDate !== "",
  ].filter(Boolean).length;
  const visibleControlTowerApprovals = approvalTasks.filter((task) => task.status === "pending");

  const dismissToast = useCallback((toastId: string, immediate = false) => {
    const clearTimer = toastTimersRef.current.get(toastId);
    if (clearTimer) {
      window.clearTimeout(clearTimer);
      toastTimersRef.current.delete(toastId);
    }

    setToasts((current) =>
      current.map((toast) => (toast.id === toastId ? { ...toast, dismissed: true, paused: true } : toast)),
    );

    const remove = () => {
      setToasts((current) => current.filter((toast) => toast.id !== toastId));
    };

    if (immediate) {
      remove();
      return;
    }

    window.setTimeout(remove, 220);
  }, []);

  const expireToast = useCallback(
    async (toastId: string) => {
      const toast = toastsRef.current.find((candidate) => candidate.id === toastId);
      if (!toast) {
        return;
      }

      try {
        await toast.onExpire?.();
      } catch (caughtError) {
        const nextMessage = caughtError instanceof Error ? caughtError.message : "Action failed.";
        const nextToastId = createClientId("toast");
        setToasts((current) => [
          ...current,
          {
            id: nextToastId,
            kind: "error",
            text: nextMessage,
            createdAt: Date.now(),
            dismissed: false,
            paused: false,
            durationMs: 6000,
            remainingMs: 6000,
            lastResumedAt: Date.now(),
            cycle: 0,
          },
        ]);
      } finally {
        dismissToast(toastId);
      }
    },
    [dismissToast],
  );

  const scheduleToastExpiry = useCallback(
    (toastId: string, delay: number) => {
      const existingTimer = toastTimersRef.current.get(toastId);
      if (existingTimer) {
        window.clearTimeout(existingTimer);
      }

      const nextTimer = window.setTimeout(() => {
        void expireToast(toastId);
      }, delay);
      toastTimersRef.current.set(toastId, nextTimer);
    },
    [expireToast],
  );

  const addToast = useCallback(
    (kind: ToastKind, text: string, options?: { durationMs?: number; action?: ToastAction; onExpire?: () => Promise<void> | void }) => {
      const toastId = createClientId("toast");
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
    },
    [scheduleToastExpiry],
  );

  const pauseToast = useCallback((toastId: string) => {
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
      current.map((candidate) =>
        candidate.id === toastId
          ? {
              ...candidate,
              paused: true,
              remainingMs,
            }
          : candidate,
      ),
    );
  }, []);

  const resumeToast = useCallback(
    (toastId: string) => {
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
    },
    [scheduleToastExpiry],
  );

  const setMessage = useCallback(
    (text: string | null) => {
      if (text) {
        addToast("success", text);
      }
    },
    [addToast],
  );

  const setError = useCallback(
    (text: string | null) => {
      if (text) {
        addToast("error", text);
      }
    },
    [addToast],
  );

  const applyStudioUrlState = useEffectEvent((searchParams: URLSearchParams): void => {
    const nextWorkspaceParam = searchParams.get("workspace");
    const nextWorkspace = WORKSPACES.some((entry) => entry.key === nextWorkspaceParam)
      ? (nextWorkspaceParam as WorkspaceKey)
      : workspace;
    setWorkspace(nextWorkspace);

    const nextTabKey = searchParams.get("tab");
    if (nextTabKey) {
      setActiveTabByWorkspace((current) => ({
        ...current,
        [nextWorkspace]: getTabLabelFromKey(nextWorkspace, nextTabKey),
      }));
    }

    const objectId = searchParams.get("objectId");
    const pageId = searchParams.get("pageId");
    const formId = searchParams.get("formId");
    const workflowId = searchParams.get("workflowId");
    const agentId = searchParams.get("agentId");
    const sectionId = searchParams.get("sectionId");
    const componentId = searchParams.get("componentId");
    const device = searchParams.get("device");
    const detailType = searchParams.get("detailType");
    const detailId = searchParams.get("detailId");

    const nextObject = manifest.objects.find((candidate) => candidate.id === objectId);
    if (nextObject) {
      setSelectedObjectId(nextObject.id);
      setObjectDraft(createObjectDraftFromDefinition(nextObject));
    }

    const nextPage = manifest.pages.find((candidate) => candidate.id === pageId);
    if (nextPage) {
      const nextLayout = structuredClone(manifest.layouts.find((candidate) => candidate.key === nextPage.layoutKey) ?? createBlankLayout(nextPage.key, nextPage.title));
      setSelectedPageId(nextPage.id);
      setPageDraft(createPageDraftFromDefinition(nextPage));
      setLayoutDraft(nextLayout);
      setSelectedSectionId(nextLayout.sections[0]?.id ?? sectionId ?? "");
      setSelectedComponentId(nextLayout.sections[0]?.components[0]?.id ?? componentId ?? "");
      setDesignerHistory([]);
      setDesignerDirty(false);
      setAutosaveStatus("idle");
    }

    const nextForm = manifest.forms.find((candidate) => candidate.id === formId);
    if (nextForm) {
      setSelectedFormId(nextForm.id);
      setFormDraft(createFormDraftFromDefinition(nextForm));
      void refreshFormSubmissions(nextForm.key);
    }

    const nextWorkflow = manifest.workflows.find((candidate) => candidate.id === workflowId);
    if (nextWorkflow) {
      setSelectedWorkflowId(nextWorkflow.id);
      setSelectedWorkflowNodeId(nextWorkflow.nodes[0]?.id ?? "");
      setWorkflowDraft(createWorkflowDraftFromDefinition(nextWorkflow));
      setWorkflowEdgeDraft({
        sourceId: nextWorkflow.nodes[0]?.id ?? "",
        targetId: nextWorkflow.nodes[1]?.id ?? nextWorkflow.nodes[0]?.id ?? "",
        label: "",
      });
    }

    const nextAgent = manifest.agents.find((candidate) => candidate.id === agentId);
    if (nextAgent) {
      setSelectedAgentId(nextAgent.id);
      setAgentDraft(createAgentDraftFromDefinition(nextAgent));
      setAgentPreview(null);
    }

    if (sectionId) {
      setSelectedSectionId(sectionId);
    }

    if (componentId) {
      setSelectedComponentId(componentId);
    }

    if (device === "desktop" || device === "tablet" || device === "mobile") {
      setPreviewDevice(device);
    }

    setControlTowerFilters({
      status: searchParams.get("status") ?? "all",
      workflowKey: searchParams.get("workflowKey") ?? "",
      agentKey: searchParams.get("agentKey") ?? "",
      severity: searchParams.get("severity") ?? "all",
      fromDate: searchParams.get("fromDate") ?? "",
      toDate: searchParams.get("toDate") ?? "",
    });

    if (
      detailId &&
      detailType &&
      ["workflow-run", "agent-run", "delivery", "alert", "approval", "dead-letter"].includes(detailType)
    ) {
      setControlTowerSelection({ type: detailType as ControlTowerSelectionType, id: detailId });
      setControlTowerDetail(null);
    }
  });

  const buildStudioUrl = useEffectEvent((): string => {
    const params = new URLSearchParams();
    params.set("workspace", workspace);
    params.set("tab", tabKeyFromLabel(WORKSPACE_TABS[workspace][activeTab] ?? WORKSPACE_TABS[workspace][0]));

    if (selectedObjectId) params.set("objectId", selectedObjectId);
    if (selectedPageId) params.set("pageId", selectedPageId);
    if (selectedFormId) params.set("formId", selectedFormId);
    if (selectedWorkflowId) params.set("workflowId", selectedWorkflowId);
    if (selectedAgentId) params.set("agentId", selectedAgentId);
    if (selectedSectionId) params.set("sectionId", selectedSectionId);
    if (selectedComponentId) params.set("componentId", selectedComponentId);
    if (previewDevice !== "desktop") params.set("device", previewDevice);

    if (controlTowerFilters.status !== "all") params.set("status", controlTowerFilters.status);
    if (controlTowerFilters.workflowKey) params.set("workflowKey", controlTowerFilters.workflowKey);
    if (controlTowerFilters.agentKey) params.set("agentKey", controlTowerFilters.agentKey);
    if (controlTowerFilters.severity !== "all") params.set("severity", controlTowerFilters.severity);
    if (controlTowerFilters.fromDate) params.set("fromDate", controlTowerFilters.fromDate);
    if (controlTowerFilters.toDate) params.set("toDate", controlTowerFilters.toDate);

    if (controlTowerSelection) {
      params.set("detailType", controlTowerSelection.type);
      params.set("detailId", controlTowerSelection.id);
    }

    const query = params.toString();
    return `${window.location.pathname}${query ? `?${query}` : ""}`;
  });

  useEffect(() => {
    if (typeof window === "undefined") {
      return;
    }

    const handlePopState = () => {
      applyStudioUrlState(new URLSearchParams(window.location.search));
      studioUrlRef.current = `${window.location.pathname}${window.location.search}`;
    };

    applyStudioUrlState(new URLSearchParams(window.location.search));
    studioUrlRef.current = `${window.location.pathname}${window.location.search}`;
    setHasHydratedUrlState(true);
    window.addEventListener("popstate", handlePopState);
    return () => window.removeEventListener("popstate", handlePopState);
  }, [manifest.agents, manifest.forms, manifest.objects, manifest.pages, manifest.workflows]);

  useEffect(() => {
    if (typeof window === "undefined" || !hasHydratedUrlState) {
      return;
    }

    const nextUrl = buildStudioUrl();
    if (studioUrlRef.current === nextUrl) {
      return;
    }

    window.history.pushState({}, "", nextUrl);
    studioUrlRef.current = nextUrl;
  }, [
    activeTab,
    controlTowerFilters,
    controlTowerSelection,
    hasHydratedUrlState,
    previewDevice,
    selectedAgentId,
    selectedComponentId,
    selectedFormId,
    selectedObjectId,
    selectedPageId,
    selectedSectionId,
    selectedWorkflowId,
    workspace,
  ]);

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

  useEffect(() => {
    toastsRef.current = toasts;
  }, [toasts]);

  useEffect(() => {
    const timers = toastTimersRef.current;
    return () => {
      for (const timer of timers.values()) {
        window.clearTimeout(timer);
      }
      if (inlineDeleteTimerRef.current) {
        window.clearTimeout(inlineDeleteTimerRef.current);
      }
      if (inlineDeleteIntervalRef.current) {
        window.clearInterval(inlineDeleteIntervalRef.current);
      }
    };
  }, []);

  useEffect(() => {
    if (!confirmDialog) {
      return;
    }

    window.setTimeout(() => {
      confirmCancelButtonRef.current?.focus();
    }, 0);
  }, [confirmDialog]);

  useEffect(() => {
    if (!confirmDialog || !confirmDialogRef.current) {
      return;
    }

    function handleFocusTrap(event: KeyboardEvent) {
      if (event.key !== "Tab" || !confirmDialogRef.current) {
        return;
      }

      const focusable = [...confirmDialogRef.current.querySelectorAll<HTMLElement>('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])')]
        .filter((candidate) => !candidate.hasAttribute("disabled"));

      if (focusable.length === 0) {
        return;
      }

      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      const activeElement = document.activeElement as HTMLElement | null;

      if (event.shiftKey && activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    }

    document.addEventListener("keydown", handleFocusTrap);
    return () => document.removeEventListener("keydown", handleFocusTrap);
  }, [confirmDialog]);

  useEffect(() => {
    if (!sectionActionMenuId) {
      return;
    }

    function handlePointerDown(event: MouseEvent) {
      const target = event.target;
      if (!(target instanceof HTMLElement) || !target.closest("[data-section-menu-root]")) {
        setSectionActionMenuId(null);
      }
    }

    document.addEventListener("mousedown", handlePointerDown);
    return () => document.removeEventListener("mousedown", handlePointerDown);
  }, [sectionActionMenuId]);

  useEffect(() => {
    function handleScroll() {
      setScrolledPast(window.scrollY > 100);
    }

    handleScroll();
    window.addEventListener("scroll", handleScroll, { passive: true });
    return () => window.removeEventListener("scroll", handleScroll);
  }, []);

  const handleGlobalKeyDown = useEffectEvent((event: KeyboardEvent) => {
    const meta = event.metaKey || event.ctrlKey;
    const tag = (document.activeElement?.tagName ?? "").toLowerCase();
    const isInput = tag === "input" || tag === "textarea" || tag === "select";

    if (meta && event.key === "k") {
      event.preventDefault();
      setCommandPaletteOpen((prev) => !prev);
      setCommandPaletteQuery("");
      setCommandPaletteIndex(0);
      return;
    }

    if (meta && event.key === "s") {
      event.preventDefault();
      handleSaveShortcut();
      return;
    }

    if (meta && event.key === "\\") {
      event.preventDefault();
      setSidebarCollapsed((current) => !current);
      return;
    }

    if (meta && event.key === ".") {
      event.preventDefault();
      cycleWorkspaceTab();
      return;
    }

    if (event.key === "Escape" && commandPaletteOpen) {
      event.preventDefault();
      closeCommandPalette();
      return;
    }

    if (event.key === "Escape" && confirmDialog) {
      event.preventDefault();
      closeConfirmDialog();
      return;
    }

    if (event.key === "Escape" && shortcutsOpen) {
      event.preventDefault();
      setShortcutsOpen(false);
      previouslyFocusedElementRef.current?.focus();
      return;
    }

    if (event.key === "Escape" && sectionActionMenuId) {
      event.preventDefault();
      setSectionActionMenuId(null);
      return;
    }

    if (event.key === "Escape" && !isInput) {
      deselectCurrentSelection();
    }

    if (!isInput && !commandPaletteOpen && meta) {
      const digitMatch = event.key.match(/^([1-9])$/);
      if (digitMatch) {
        const index = Number(digitMatch[1]) - 1;
        if (index < WORKSPACES.length) {
          event.preventDefault();
          selectWorkspace(WORKSPACES[index].key);
        }
      }
    }
  });

  // Command palette keyboard listener
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => handleGlobalKeyDown(event);
    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, []);

  // Load collapsed groups from localStorage
  useEffect(() => {
    try {
      const stored = window.localStorage.getItem("ta-platform-collapsed-groups");
      if (stored) setCollapsedGroups(new Set(JSON.parse(stored) as string[]));
      const storedRecents = window.localStorage.getItem("ta-platform-command-recents");
      if (storedRecents) {
        setCommandPaletteRecentIds(JSON.parse(storedRecents) as string[]);
      }
    } catch { /* optional */ }
  }, []);

  // Persist collapsed groups
  useEffect(() => {
    try {
      window.localStorage.setItem("ta-platform-collapsed-groups", JSON.stringify([...collapsedGroups]));
      window.localStorage.setItem("ta-platform-command-recents", JSON.stringify(commandPaletteRecentIds));
    } catch { /* optional */ }
  }, [collapsedGroups, commandPaletteRecentIds]);

  const refreshPublishPreview = useCallback(async (): Promise<void> => {
    setIsPublishPreviewLoading(true);
    try {
      const payload = await fetchJson<{ preview: PlatformPublishPreview }>(`/api/platform/tenants/${bootstrap.tenant.slug}/publish/preview`);
      setPublishPreview(payload.preview);
    } catch {
      setPublishPreview(null);
    } finally {
      setIsPublishPreviewLoading(false);
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
    setIsControlTowerLoading(true);
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
    } finally {
      setIsControlTowerLoading(false);
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

  // Unsaved changes guard — warn on browser close/navigate
  useEffect(() => {
    if (!designerDirty) return;
    const handler = (e: BeforeUnloadEvent) => { e.preventDefault(); };
    window.addEventListener("beforeunload", handler);
    return () => window.removeEventListener("beforeunload", handler);
  }, [designerDirty]);

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

  async function handleSavePageTemplate(): Promise<void> {
    const pageKey = pageDraft.key || selectedPage?.key;
    if (!pageKey) {
      setError("Save the page first, then save it as a template.");
      return;
    }

    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/page-templates`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            pageKey,
            label: pageDraft.title || "Reusable page template",
            description: pageDraft.description || "Saved from the page builder.",
          }),
        });
      },
      `Saved ${pageDraft.title || pageKey} as a tenant page template.`,
    );
  }

  async function handleSaveSectionTemplate(sectionId?: string): Promise<void> {
    const targetSectionId = sectionId ?? selectedSection?.id;
    const targetSection = layoutDraft?.sections.find((section) => section.id === targetSectionId);
    if (!layoutDraft || !targetSectionId || !targetSection) {
      setError("Select a section before saving a template.");
      return;
    }

    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/section-templates`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            layoutKey: layoutDraft.key,
            sectionId: targetSectionId,
            label: targetSection.title,
            description: targetSection.description || "Saved from the page builder.",
          }),
        });
      },
      `Saved ${targetSection.title} as a section template.`,
    );
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

  async function handleTestNotificationChannel(channelKey: string): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/notifications/channels/${channelKey}/test`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            subject: "Platform channel test",
            body: `Control Tower sent a delivery test for ${channelKey}.`,
          }),
        });
        await refreshControlTower();
      },
      `Sent test notification for ${channelKey}.`,
    );
  }

  async function handleToggleNotificationChannel(channelKey: string, enabled: boolean): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/notifications/channels/${channelKey}/toggle`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({ enabled }),
        });
        await refreshControlTower();
      },
      `${enabled ? "Enabled" : "Disabled"} channel ${channelKey}.`,
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

  async function handleSaveWorkflowTemplate(): Promise<void> {
    const workflowId = selectedWorkflow?.id || workflowDraft.id;
    if (!workflowId) {
      setError("Save the workflow first, then save it as a template.");
      return;
    }
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/workflow-templates`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
          },
          body: JSON.stringify({
            workflowId,
            name: workflowDraft.name || "Reusable workflow template",
            description: workflowDraft.description || "Saved from the workflow studio.",
          }),
        });
      },
      `Saved ${workflowDraft.name || workflowId} as a workflow template.`,
    );
  }

  async function handleSaveWorkflowTestCase(): Promise<void> {
    const workflowKey = workflowDraft.key || selectedWorkflow?.key;
    if (!workflowKey) {
      setError("Save the workflow before storing a test case.");
      return;
    }
    try {
      const payload = parseJsonPayloadSafely(workflowTestPayload);
      await executeAction(
        async () => {
          await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/workflow-tests`, {
            method: "POST",
            headers: {
              "content-type": "application/json",
            },
            body: JSON.stringify({
              name: `${workflowDraft.name || workflowKey} draft test`,
              workflowKey,
              payload,
              expectedStatus: "SUCCEEDED",
            }),
          });
        },
        `Saved a named workflow test for ${workflowDraft.name || workflowKey}.`,
      );
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Invalid workflow test payload.");
    }
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

  async function handleAcknowledgeAlert(alertId: string): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/alerts/${alertId}/acknowledge`, {
          method: "POST",
        });
        await refreshControlTower();
      },
      "Alert acknowledged.",
    );
  }

  async function handleRetryDelivery(deliveryId: string): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/deliveries/${deliveryId}/retry`, {
          method: "POST",
        });
        await refreshControlTower();
      },
      "Notification delivery retried.",
    );
  }

  async function handleResumeAgentRun(runId: string): Promise<void> {
    await executeAction(
      async () => {
        await fetchJson(`/api/platform/tenants/${bootstrap.tenant.slug}/agent-runs/${runId}/resume`, {
          method: "POST",
        });
        await refreshControlTower();
      },
      "Agent run resumed.",
    );
  }

  async function handleInspectControlTowerDetail(type: ControlTowerSelectionType, id: string): Promise<void> {
    try {
      setError(null);
      setControlTowerSelection({ type, id });
      const endpoint =
        type === "workflow-run"
          ? `/api/platform/tenants/${bootstrap.tenant.slug}/workflow-runs/${id}`
          : type === "agent-run"
            ? `/api/platform/tenants/${bootstrap.tenant.slug}/agent-runs/${id}`
            : type === "delivery"
              ? `/api/platform/tenants/${bootstrap.tenant.slug}/deliveries/${id}`
              : type === "alert"
                ? `/api/platform/tenants/${bootstrap.tenant.slug}/alerts/${id}`
                : type === "approval"
                  ? `/api/platform/tenants/${bootstrap.tenant.slug}/approval-tasks/${id}`
                  : `/api/platform/tenants/${bootstrap.tenant.slug}/dead-letters/${id}`;
      const payload = await fetchJson<Record<string, unknown>>(endpoint);
      const detail =
        "detail" in payload && payload.detail && typeof payload.detail === "object" && !Array.isArray(payload.detail)
          ? (payload.detail as Record<string, unknown>)
          : payload;
      setControlTowerDetail(detail);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Failed to load operator detail.");
      setControlTowerDetail(null);
    }
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
          availableTemplates: bootstrap.designerCatalog.sectionTemplates,
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

  function clearInlineDelete(showCancelledToast = false): void {
    if (inlineDeleteTimerRef.current) {
      window.clearTimeout(inlineDeleteTimerRef.current);
      inlineDeleteTimerRef.current = null;
    }
    if (inlineDeleteIntervalRef.current) {
      window.clearInterval(inlineDeleteIntervalRef.current);
      inlineDeleteIntervalRef.current = null;
    }
    inlineDeleteActionRef.current = null;
    setPendingInlineDelete(null);
    if (showCancelledToast) {
      addToast("info", "Delete cancelled.");
    }
  }

  function scheduleInlineDelete(key: string, label: string, action: () => void | Promise<void>): void {
    clearInlineDelete(false);
    const expiresAt = Date.now() + 5000;
    inlineDeleteActionRef.current = action;
    setPendingInlineDelete({ key, expiresAt });
    inlineDeleteIntervalRef.current = window.setInterval(() => {
      setPendingInlineDelete((current) => (current ? { ...current } : current));
    }, 250);
    inlineDeleteTimerRef.current = window.setTimeout(() => {
      const nextAction = inlineDeleteActionRef.current;
      clearInlineDelete(false);
      void Promise.resolve(nextAction?.()).catch((caughtError) => {
        setError(caughtError instanceof Error ? caughtError.message : "Delete failed.");
      });
    }, 5000);
    addToast("info", `${label} queued for deletion.`, {
      durationMs: 5000,
      action: {
        label: "Undo",
        onClick: () => clearInlineDelete(true),
      },
    });
  }

  function inlineDeleteCountdown(key: string): number | null {
    if (!pendingInlineDelete || pendingInlineDelete.key !== key) {
      return null;
    }

    return Math.max(1, Math.ceil((pendingInlineDelete.expiresAt - Date.now()) / 1000));
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

  function selectObject(objectDefinition?: ObjectDefinition): void {
    const nextObject = objectDefinition ?? manifest.objects[0];
    setSelectedObjectId(nextObject?.id ?? "");
    setObjectDraft(createObjectDraftFromDefinition(nextObject));
  }

  function selectForm(form?: FormDefinition): void {
    const nextForm = form ?? manifest.forms[0];
    setSelectedFormId(nextForm?.id ?? "");
    setFormDraft(createFormDraftFromDefinition(nextForm));
    if (nextForm?.key) {
      void refreshFormSubmissions(nextForm.key);
    } else {
      setFormSubmissions([]);
    }
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

  function selectWorkspace(nextWorkspace: WorkspaceKey): void {
    setWorkspace(nextWorkspace);
    setWorkspaceTransitionKey((prev) => prev + 1);
  }

  function cycleWorkspaceTab(): void {
    const nextTabIndex = (activeTab + 1) % WORKSPACE_TABS[workspace].length;
    setActiveTab(nextTabIndex);
  }

  function deselectCurrentSelection(): void {
    if (workspace === "data-model") {
      selectObject(undefined);
      return;
    }

    if (workspace === "pages") {
      applySelectedPage(undefined);
      return;
    }

    if (workspace === "forms") {
      selectForm(undefined);
      return;
    }

    if (workspace === "workflows") {
      selectWorkflow(undefined);
      return;
    }

    if (workspace === "agents") {
      selectAgent(undefined);
      return;
    }

    if (workspace === "control-tower") {
      setControlTowerSelection(null);
      setControlTowerDetail(null);
    }
  }

  function handleSaveShortcut(): void {
    if (workspace === "data-model") {
      void handleObjectSave();
      return;
    }

    if (workspace === "pages") {
      void handleLayoutSave();
      return;
    }

    if (workspace === "forms") {
      void handleFormSave();
      return;
    }

    if (workspace === "branding") {
      void handleBrandingSave();
      return;
    }

    if (workspace === "profiles") {
      void handleProfileSave();
      return;
    }

    if (workspace === "navigation") {
      if (activeTab === 0) {
        void handleMenuSave();
      } else if (activeTab === 1) {
        void handleAppShellSave();
      } else {
        void handleNotificationsSave();
      }
      return;
    }

    if (workspace === "workflows") {
      void handleWorkflowSave();
      return;
    }

    if (workspace === "agents") {
      void handleAgentSave();
      return;
    }

    if (workspace === "models") {
      void handleProviderSave();
      return;
    }

    if (workspace === "security") {
      void handleSecuritySave();
    }
  }

  function recordPaletteAction(itemId: string): void {
    setCommandPaletteRecentIds((current) => [itemId, ...current.filter((candidate) => candidate !== itemId)].slice(0, 5));
  }

  function closeCommandPalette(): void {
    setCommandPaletteOpen(false);
    window.setTimeout(() => {
      previouslyFocusedElementRef.current?.focus();
    }, 0);
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

  // ── Confirm dialog helper ──
  function requestConfirmation(title: string, msg: string, confirmLabel: string, onConfirm: () => void) {
    if (typeof document !== "undefined" && document.activeElement instanceof HTMLElement) {
      previouslyFocusedElementRef.current = document.activeElement;
    }
    setConfirmDialog({ title, message: msg, confirmLabel, onConfirm });
  }

  function closeConfirmDialog(): void {
    setConfirmDialog(null);
    window.setTimeout(() => {
      previouslyFocusedElementRef.current?.focus();
    }, 0);
  }

  function renderConfirmDialog() {
    if (!confirmDialog) return null;
    return (
      <div className={styles.confirmOverlay} onClick={closeConfirmDialog} role="presentation">
        <div
          aria-label={confirmDialog.title}
          aria-modal="true"
          className={styles.confirmDialog}
          onClick={(event) => event.stopPropagation()}
          ref={confirmDialogRef}
          role="alertdialog"
        >
          <div className={styles.confirmIconDanger} aria-hidden="true">
            <svg fill="none" viewBox="0 0 20 20">
              <path d="M10 2.2 18 17H2L10 2.2Z" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.6" />
              <path d="M10 7v4.2" stroke="currentColor" strokeLinecap="round" strokeWidth="1.6" />
              <circle cx="10" cy="13.8" fill="currentColor" r="1" />
            </svg>
          </div>
          <h3>{confirmDialog.title}</h3>
          <p>{confirmDialog.message}</p>
          <div className={styles.confirmActions}>
            <button className={styles.confirmCancelButton} onClick={closeConfirmDialog} ref={confirmCancelButtonRef} type="button">
              Cancel
            </button>
            <button
              className={styles.confirmDangerButton}
              onClick={() => {
                confirmDialog.onConfirm();
                closeConfirmDialog();
              }}
              type="button"
            >
              {confirmDialog.confirmLabel}
            </button>
          </div>
        </div>
      </div>
    );
  }

  // ── Command palette ──
  const commandPaletteItems = (() => {
    type PaletteItem = { id: string; label: string; hint: string; code: string; category: string; action: () => void };
    const items: PaletteItem[] = [];

    for (const entry of WORKSPACES) {
      items.push({
        id: `workspace:${entry.key}`,
        label: entry.label,
        hint: `Workspace · ${entry.note}`,
        code: entry.code,
        category: "Workspaces",
        action: () => selectWorkspace(entry.key),
      });
    }

    for (const entry of WORKSPACES) {
      const tabs = WORKSPACE_TABS[entry.key];
      for (let index = 0; index < tabs.length; index += 1) {
        items.push({
          id: `tab:${entry.key}:${tabKeyFromLabel(tabs[index])}`,
          label: `${entry.label} > ${tabs[index]}`,
          hint: `Tab in ${entry.label}`,
          code: entry.code,
          category: "Sections",
          action: () => {
            selectWorkspace(entry.key);
            setActiveTab(index);
          },
        });
      }
    }

    for (const objectDefinition of manifest.objects) {
      items.push({
        id: `object:${objectDefinition.id}`,
        label: objectDefinition.label,
        hint: `Object · ${objectDefinition.fields.length} fields`,
        code: "DM",
        category: "Objects",
        action: () => {
          selectWorkspace("data-model");
          selectObject(objectDefinition);
        },
      });
    }

    for (const page of manifest.pages) {
      items.push({
        id: `page:${page.id}`,
        label: page.title,
        hint: `Page · /${page.route}`,
        code: "PG",
        category: "Pages",
        action: () => {
          selectWorkspace("pages");
          applySelectedPage(page);
        },
      });
    }

    for (const workflow of manifest.workflows) {
      items.push({
        id: `workflow:${workflow.id}`,
        label: workflow.name,
        hint: `Workflow · ${workflow.nodes.length} nodes`,
        code: "WF",
        category: "Workflows",
        action: () => {
          selectWorkspace("workflows");
          selectWorkflow(workflow);
        },
      });
    }

    for (const agent of manifest.agents) {
      items.push({
        id: `agent:${agent.id}`,
        label: agent.name,
        hint: `Agent · ${agent.scope}`,
        code: "AG",
        category: "Agents",
        action: () => {
          selectWorkspace("agents");
          selectAgent(agent);
        },
      });
    }

    return items;
  })();

  const recentPaletteItems = useMemo(() => {
    return commandPaletteRecentIds
      .map((itemId) => commandPaletteItems.find((item) => item.id === itemId))
      .filter((item): item is NonNullable<typeof item> => Boolean(item));
  }, [commandPaletteItems, commandPaletteRecentIds]);

  const filteredPaletteItems = useMemo(() => {
    if (!commandPaletteQuery.trim()) {
      return commandPaletteItems.slice(0, 18);
    }

    const q = commandPaletteQuery.toLowerCase();
    return commandPaletteItems
      .filter((item) => item.label.toLowerCase().includes(q) || item.hint.toLowerCase().includes(q))
      .slice(0, 20);
  }, [commandPaletteItems, commandPaletteQuery]);

  const groupedPaletteItems = useMemo(() => {
    const groups = new Map<string, typeof filteredPaletteItems>();
    for (const item of filteredPaletteItems) {
      groups.set(item.category, [...(groups.get(item.category) ?? []), item]);
    }
    return [...groups.entries()];
  }, [filteredPaletteItems]);

  function renderShortcutsDialog() {
    if (!shortcutsOpen) {
      return null;
    }

    return (
      <div
        className={styles.confirmOverlay}
        onClick={() => {
          setShortcutsOpen(false);
          previouslyFocusedElementRef.current?.focus();
        }}
        role="presentation"
      >
        <div aria-label="Keyboard shortcuts" aria-modal="true" className={styles.confirmDialog} onClick={(event) => event.stopPropagation()} role="dialog">
          <h3>Keyboard Shortcuts</h3>
          <div className={styles.listStack}>
            {[
              ["⌘ K", "Open the command palette"],
              ["⌘ S", "Save the active draft"],
              ["⌘ \\", "Toggle the sidebar"],
              ["⌘ .", "Cycle to the next tab"],
              ["Esc", "Deselect the current item or close overlays"],
            ].map(([key, meaning]) => (
              <div className={styles.sidebarMeta} key={key}>
                <span>{meaning}</span>
                <strong>{key}</strong>
              </div>
            ))}
          </div>
          <div className={styles.confirmActions}>
            <button
              className={styles.confirmCancelButton}
              onClick={() => {
                setShortcutsOpen(false);
                previouslyFocusedElementRef.current?.focus();
              }}
              type="button"
            >
              Close
            </button>
          </div>
        </div>
      </div>
    );
  }

  function renderCommandPalette() {
    if (!commandPaletteOpen) {
      return null;
    }

    return (
      <div className={styles.commandPaletteOverlay} onClick={closeCommandPalette} role="presentation">
        <div className={styles.commandPaletteDialog} onClick={(event) => event.stopPropagation()} role="dialog" aria-label="Command palette">
          <input
            autoFocus
            className={styles.commandPaletteInput}
            onChange={(event) => {
              setCommandPaletteQuery(event.target.value);
              setCommandPaletteIndex(0);
            }}
            onKeyDown={(event) => {
              const flatItems = commandPaletteQuery.trim()
                ? filteredPaletteItems
                : [...recentPaletteItems, ...filteredPaletteItems.filter((item) => !recentPaletteItems.some((recent) => recent.id === item.id))];
              if (event.key === "ArrowDown") {
                event.preventDefault();
                setCommandPaletteIndex((prev) => Math.min(prev + 1, Math.max(0, flatItems.length - 1)));
              } else if (event.key === "ArrowUp") {
                event.preventDefault();
                setCommandPaletteIndex((prev) => Math.max(prev - 1, 0));
              } else if (event.key === "Enter" && flatItems[commandPaletteIndex]) {
                event.preventDefault();
                recordPaletteAction(flatItems[commandPaletteIndex].id);
                flatItems[commandPaletteIndex].action();
                closeCommandPalette();
              }
            }}
            placeholder="Search workspaces, pages, workflows, and agents…"
            value={commandPaletteQuery}
          />
          <div className={styles.commandPaletteResults}>
            {!commandPaletteQuery.trim() && recentPaletteItems.length > 0 ? (
              <>
                <div className={styles.commandPaletteSectionHeader}>Recent</div>
                {recentPaletteItems.map((item, index) => (
                  <button
                    className={index === commandPaletteIndex ? styles.commandPaletteResultActive : styles.commandPaletteResult}
                    key={`recent-${item.id}`}
                    onClick={() => {
                      recordPaletteAction(item.id);
                      item.action();
                      closeCommandPalette();
                    }}
                    onMouseEnter={() => setCommandPaletteIndex(index)}
                    type="button"
                  >
                    <span className={styles.commandPaletteRecent}>Recent</span>
                    <span className={styles.commandPaletteResultIcon}>{item.code}</span>
                    <div className={styles.commandPaletteResultCopy}>
                      <span>{item.label}</span>
                      <small>{item.hint}</small>
                    </div>
                  </button>
                ))}
              </>
            ) : null}

            {groupedPaletteItems.length === 0 ? (
              <div className={styles.commandPaletteEmpty}>No results found.</div>
            ) : (
              groupedPaletteItems.map(([category, items], groupIndex) => (
                <div key={category}>
                  <div className={styles.commandPaletteSectionHeader}>{category}</div>
                  {items.map((item, itemIndex) => {
                    const absoluteIndex =
                      (!commandPaletteQuery.trim() ? recentPaletteItems.length : 0) +
                      groupedPaletteItems
                        .slice(0, groupIndex)
                        .reduce((count, [, previousItems]) => count + previousItems.length, 0) +
                      itemIndex;

                    return (
                      <button
                        className={absoluteIndex === commandPaletteIndex ? styles.commandPaletteResultActive : styles.commandPaletteResult}
                        key={item.id}
                        onClick={() => {
                          recordPaletteAction(item.id);
                          item.action();
                          closeCommandPalette();
                        }}
                        onMouseEnter={() => setCommandPaletteIndex(absoluteIndex)}
                        type="button"
                      >
                        <span className={styles.commandPaletteResultIcon}>{item.code}</span>
                        <div className={styles.commandPaletteResultCopy}>
                          <span>{item.label}</span>
                          <small>{item.hint}</small>
                        </div>
                      </button>
                    );
                  })}
                </div>
              ))
            )}
          </div>
          <div className={styles.commandPaletteFooter}>
            <span><kbd>↑↓</kbd> navigate</span>
            <span><kbd>↵</kbd> select</span>
            <button
              className={styles.commandPaletteFooterButton}
              onClick={() => {
                if (typeof document !== "undefined" && document.activeElement instanceof HTMLElement) {
                  previouslyFocusedElementRef.current = document.activeElement;
                }
                setShortcutsOpen(true);
              }}
              type="button"
            >
              Keyboard shortcuts
            </button>
          </div>
        </div>
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

  function renderSkeleton(rows: number, asCard = false) {
    return Array.from({ length: rows }, (_, index) => (
      <div className={asCard ? styles.skeletonCard : styles.skeletonRow} key={index}>
        <div className={`${styles.skeleton} ${styles.skeletonWide}`} />
        <div className={`${styles.skeleton} ${styles.skeletonNarrow}`} />
      </div>
    ));
  }

  function renderEmptyState(
    icon: "database" | "workflow" | "agent" | "shield" | "inbox",
    title: string,
    hint: string,
    ctaLabel?: string,
    ctaAction?: () => void,
  ) {
    const iconMarkup =
      icon === "workflow" ? (
        <svg fill="none" viewBox="0 0 24 24">
          <path d="M6 6h4v4H6zM14 14h4v4h-4zM14 6h4v4h-4zM10 8h4M16 10v4" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.6" />
        </svg>
      ) : icon === "agent" ? (
        <svg fill="none" viewBox="0 0 24 24">
          <path d="M9 18h6M8 6h8M7 9h10v6H7zM12 4v2M9 13h.01M15 13h.01" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.6" />
        </svg>
      ) : icon === "shield" ? (
        <svg fill="none" viewBox="0 0 24 24">
          <path d="m12 3 7 3v5c0 4.3-2.6 7.3-7 10-4.4-2.7-7-5.7-7-10V6l7-3Z" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.6" />
        </svg>
      ) : icon === "inbox" ? (
        <svg fill="none" viewBox="0 0 24 24">
          <path d="M4 6h16v10H15l-2 3-2-3H4z" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.6" />
        </svg>
      ) : (
        <svg fill="none" viewBox="0 0 24 24">
          <path d="M5 6h14v12H5zM9 10h6M9 14h4" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.6" />
        </svg>
      );

    return (
      <div className={styles.emptyState}>
        <div className={styles.emptyStateIcon} aria-hidden="true">
          {iconMarkup}
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

  function getTabBadge(workspaceKey: WorkspaceKey, index: number): { value: string; danger?: boolean } | null {
    if (workspaceKey === "control-tower" && index === 1 && filteredAlerts.length > 0) {
      return { value: String(filteredAlerts.length) };
    }
    if (workspaceKey === "control-tower" && index === 2 && deadLetters.length > 0) {
      return { value: String(deadLetters.length), danger: true };
    }
    if (workspaceKey === "agents" && index === 2 && agentRuns.filter((run) => run.agentId === agentDraft.id).length > 0) {
      return { value: String(agentRuns.filter((run) => run.agentId === agentDraft.id).length) };
    }
    return null;
  }

  function renderTabBar() {
    const tabs = WORKSPACE_TABS[workspace];
    return (
      <div className={[styles.tabBar, styles.tabBarSticky, scrolledPast ? styles.tabBarCompact : ""].join(" ")}>
        {tabs.map((tabLabel, index) => {
          const badge = getTabBadge(workspace, index);
          return (
            <button
              className={index === activeTab ? styles.activeTab : styles.tab}
              data-tooltip={tabLabel}
              key={tabLabel}
              onClick={() => setActiveTab(index)}
              type="button"
            >
              {tabLabel}
              {workspace === "pages" && index === 0 && designerDirty ? <span className={styles.unsavedDot} data-tooltip="Unsaved changes" /> : null}
              {badge ? <span className={badge.danger ? styles.tabBadgeDanger : styles.tabBadge}>{badge.value}</span> : null}
            </button>
          );
        })}
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
            {manifest.objects.length === 0 ? (
              <div className={styles.emptyState}>
                <p>No objects defined yet.</p>
                <p className={styles.emptyStateHint}>Objects are the core data building blocks. Start by defining your first data model.</p>
                <button className={styles.emptyStateCta} onClick={() => { setSelectedObjectId(""); setObjectDraft(createBlankObjectDraft()); }} type="button">Create your first object</button>
              </div>
            ) : (
              manifest.objects.map((objectDefinition) => (
                <button
                  className={`${objectDefinition.id === selectedObject?.id ? styles.activeListItem : styles.listItem} ${styles.staggerChild}`}
                  data-object-list-id={objectDefinition.id}
                  key={objectDefinition.id}
                  onClick={() => {
                    setSelectedObjectId(objectDefinition.id);
                    setObjectDraft(createObjectDraftFromDefinition(objectDefinition));
                  }}
                  type="button"
                >
                  <span>{objectDefinition.label}</span>
                  <div className={styles.listItemMeta}>
                    <span className={styles.metaBadge}>{objectDefinition.fields.length} fields</span>
                    {objectDefinition.primaryFieldKey ? <span className={styles.metaBadge}>{objectDefinition.primaryFieldKey}</span> : null}
                  </div>
                </button>
              ))
            )}
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
                  <button className={styles.ghostButtonDanger} onClick={() => requestConfirmation("Delete object", `Are you sure you want to delete "${selectedObject.label}"? This action cannot be undone.`, "Delete", () => void handleObjectDelete(selectedObject.id))} type="button">
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
                  <small className={styles.formFieldHelp}>Unique lowercase_snake_case identifier</small>
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
                    <button
                      className={inlineDeleteCountdown(`object-field-${field.id}`) ? styles.ghostButton : styles.ghostButtonDanger}
                      onClick={() => {
                        if (inlineDeleteCountdown(`object-field-${field.id}`)) {
                          clearInlineDelete(true);
                          return;
                        }

                        scheduleInlineDelete(`object-field-${field.id}`, `${field.label} field`, () => handleFieldDelete(field.id));
                      }}
                      type="button"
                    >
                      {inlineDeleteCountdown(`object-field-${field.id}`)
                        ? `Undo (${inlineDeleteCountdown(`object-field-${field.id}`)}s)`
                        : "Delete"}
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
                  <small className={styles.formFieldHelp}>Controls masking in agent context. PII fields are redacted by default.</small>
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
                  <small className={styles.formFieldHelp}>Must have a value before save</small>
                </label>
                <label className={styles.checkboxField}>
                  <input checked={fieldDraft.unique} onChange={(event) => setFieldDraft((current) => ({ ...current, unique: event.target.checked }))} type="checkbox" />
                  <span>Unique</span>
                  <small className={styles.formFieldHelp}>No duplicate values across records</small>
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
                  className={`${page.id === selectedPage?.id ? styles.activeListItem : styles.listItem} ${styles.staggerChild}`}
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
                <article className={`${styles.metricCard} ${styles.staggerChild}`} key={template.key}>
                  <span>{template.label}</span>
                  <strong>{template.section.components.length} components</strong>
                  <p className={styles.metricMeta}>{template.description}</p>
                  <div className={styles.inlineList}>
                    <button className={styles.secondaryButton} onClick={() => addLayoutSection(template.key)} type="button">
                      Add to page
                    </button>
                    <button
                      className={styles.ghostButton}
                      disabled={!selectedSection}
                      onClick={() => {
                        if (!selectedSection || !layoutDraft) {
                          return;
                        }
                        const replacement = createSectionFromTemplate({
                          templateKey: template.key,
                          idFactory: createClientId,
                          objectKey: pageDraft.objectKey || undefined,
                          availableTemplates: bootstrap.designerCatalog.sectionTemplates,
                        });
                        updateLayoutSection(selectedSection.id, () => ({
                          ...replacement,
                          id: selectedSection.id,
                        }));
                      }}
                      type="button"
                    >
                      Replace selected
                    </button>
                  </div>
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
            {manifest.pages.length === 0
              ? renderEmptyState("database", "No pages yet.", "Create your first page to start shaping the runtime navigation and layout.", "New page", () => {
                  pushDesignerHistory();
                  setSelectedPageId("");
                  setPageDraft(createBlankPageDraft());
                  const blankLayout = createBlankLayout(createClientId("page"), "New page");
                  setLayoutDraft(blankLayout);
                  setSelectedSectionId(blankLayout.sections[0]?.id ?? "");
                  setSelectedComponentId(blankLayout.sections[0]?.components[0]?.id ?? "");
                  markDesignerDirty();
                })
              : manifest.pages.map((page) => {
                  const pageLayout = manifest.layouts.find((l) => l.key === page.layoutKey);
                  return (
                    <button
                      className={`${page.id === selectedPage?.id ? styles.activeListItem : styles.listItem} ${styles.staggerChild}`}
                      data-page-tree-id={page.id}
                      key={page.id}
                      onClick={() => applySelectedPage(page)}
                      type="button"
                    >
                      <span>{page.title}</span>
                      <div className={styles.listItemMeta}>
                        <span className={styles.metaBadge}>/{page.route}</span>
                        {pageLayout ? <span className={styles.metaBadge}>{pageLayout.sections.length} sections</span> : null}
                        {page.isHome ? <span className={styles.metaBadge}>Home</span> : null}
                      </div>
                    </button>
                  );
                })}
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
              <strong>{pageDiagnostics.length === 0 ? `${pageReadinessScore}% ready` : `${pageDiagnostics.length} issues`}</strong>
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
              <button className={styles.secondaryButton} onClick={() => void handleSavePageTemplate()} type="button">
                Save page as template
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
                      <button className={styles.secondaryButton} onClick={() => addComponentToSection(section.id)} type="button">
                        Add component
                      </button>
                      <div className={styles.overflowMenuWrap} data-section-menu-root="">
                        <button
                          aria-expanded={sectionActionMenuId === section.id}
                          className={styles.overflowMenuTrigger}
                          data-tooltip="More section actions"
                          onClick={() => setSectionActionMenuId((current) => (current === section.id ? null : section.id))}
                          type="button"
                        >
                          <span>•••</span>
                        </button>
                        {sectionActionMenuId === section.id ? (
                          <div className={styles.overflowMenu} role="menu">
                            <button
                              className={styles.overflowMenuItem}
                              onClick={() => {
                                setSelectedSectionId(section.id);
                                setSelectedComponentId("");
                                setSectionActionMenuId(null);
                              }}
                              role="menuitem"
                              type="button"
                            >
                              Inspect
                            </button>
                            <button
                              className={styles.overflowMenuItem}
                              onClick={() => {
                                duplicateSection(section.id);
                                setSectionActionMenuId(null);
                              }}
                              role="menuitem"
                              type="button"
                            >
                              Duplicate
                            </button>
                            <button
                              className={styles.overflowMenuItem}
                              onClick={() => {
                                setSelectedSectionId(section.id);
                                void handleSaveSectionTemplate(section.id);
                                setSectionActionMenuId(null);
                              }}
                              role="menuitem"
                              type="button"
                            >
                              Save as template
                            </button>
                            <button
                              className={inlineDeleteCountdown(`section-${section.id}`) ? styles.overflowMenuItem : styles.overflowMenuItemDanger}
                              onClick={() => {
                                if (inlineDeleteCountdown(`section-${section.id}`)) {
                                  clearInlineDelete(true);
                                } else {
                                  scheduleInlineDelete(`section-${section.id}`, `${section.title} section`, () => removeSection(section.id));
                                }
                                setSectionActionMenuId(null);
                              }}
                              role="menuitem"
                              type="button"
                            >
                              {inlineDeleteCountdown(`section-${section.id}`)
                                ? `Undo (${inlineDeleteCountdown(`section-${section.id}`)}s)`
                                : "Remove"}
                            </button>
                          </div>
                        ) : null}
                      </div>
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
              renderEmptyState("database", "Select a page to begin.", "Choose a page from the tree or create a new one to start composing the runtime.")
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
                  <button
                    className={inlineDeleteCountdown(`component-${selectedComponent.id}`) ? styles.ghostButton : styles.ghostButtonDanger}
                    onClick={() => {
                      if (inlineDeleteCountdown(`component-${selectedComponent.id}`)) {
                        clearInlineDelete(true);
                        return;
                      }

                      scheduleInlineDelete(`component-${selectedComponent.id}`, `${selectedComponent.title} component`, () => removeComponent(selectedSection!.id, selectedComponent.id));
                    }}
                    type="button"
                  >
                    {inlineDeleteCountdown(`component-${selectedComponent.id}`)
                      ? `Undo (${inlineDeleteCountdown(`component-${selectedComponent.id}`)}s)`
                      : "Remove"}
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
                  className={`${styles.listItem} ${styles.staggerChild}`}
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
                  <div className={diagnostic.severity === "blocking" ? styles.errorBanner : styles.sidebarPanel} key={diagnostic.id}>
                    <strong>{diagnostic.severity.toUpperCase()}</strong> {diagnostic.message}
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
                className={`${form.id === selectedForm?.id ? styles.activeListItem : styles.listItem} ${styles.staggerChild}`}
                data-form-list-id={form.id}
                key={form.id}
                onClick={() => {
                  setSelectedFormId(form.id);
                  setFormDraft(createFormDraftFromDefinition(form));
                  void refreshFormSubmissions(form.key);
                }}
                type="button"
              >
                <span>{form.title}</span>
                <div className={styles.listItemMeta}>
                  <span className={styles.metaBadge}>/{form.route}</span>
                  <span className={styles.metaBadge}>{form.deliveryMode}</span>
                  <span className={styles.metaBadge}>{form.fields.length} fields</span>
                  <span className={styles.metaBadge}>{form.steps.length} steps</span>
                </div>
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
                        <button
                          className={inlineDeleteCountdown(`form-field-${field.id}`) ? styles.ghostButton : styles.ghostButtonDanger}
                          onClick={() => {
                            if (inlineDeleteCountdown(`form-field-${field.id}`)) {
                              clearInlineDelete(true);
                              return;
                            }

                            scheduleInlineDelete(`form-field-${field.id}`, `${field.label} field`, () => removeFormField(field.id));
                          }}
                          type="button"
                        >
                          {inlineDeleteCountdown(`form-field-${field.id}`)
                            ? `Undo (${inlineDeleteCountdown(`form-field-${field.id}`)}s)`
                            : "Remove"}
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
          <div className={styles.sidebarMeta}>
            <span>Review state</span>
            <strong>{brandingDraft.mode}</strong>
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
              <div className={styles.inlineList}>
                <button
                  className={brandingDraft.mode === "draft" ? styles.primaryButton : styles.secondaryButton}
                  onClick={() => setBrandingDraft((current) => ({ ...current, mode: "draft" }))}
                  type="button"
                >
                  Draft
                </button>
                <button
                  className={brandingDraft.mode === "review" ? styles.primaryButton : styles.secondaryButton}
                  onClick={() => setBrandingDraft((current) => ({ ...current, mode: "review" }))}
                  type="button"
                >
                  Ready for review
                </button>
                <button
                  className={brandingDraft.mode === "approved" ? styles.primaryButton : styles.secondaryButton}
                  onClick={() => setBrandingDraft((current) => ({ ...current, mode: "approved" }))}
                  type="button"
                >
                  Approved
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
                    <div className={styles.colorFieldRow}>
                      <span
                        className={styles.colorSwatch}
                        style={{ background: isValidCssColor(String(brandingDraft[key as keyof TenantBrandingDefinition] ?? "")) ? String(brandingDraft[key as keyof TenantBrandingDefinition]) : "transparent" }}
                      />
                      <input
                        className={isValidCssColor(String(brandingDraft[key as keyof TenantBrandingDefinition] ?? "")) ? styles.input : `${styles.input} ${styles.inputError}`}
                        onChange={(event) => setBrandingDraft((current) => ({ ...current, [key]: event.target.value }))}
                        value={String(brandingDraft[key as keyof TenantBrandingDefinition] ?? "")}
                      />
                      <input
                        aria-label={`${label} color picker`}
                        className={styles.colorPicker}
                        onChange={(event) => setBrandingDraft((current) => ({ ...current, [key]: event.target.value }))}
                        type="color"
                        value={isValidCssColor(String(brandingDraft[key as keyof TenantBrandingDefinition] ?? "")) ? String(brandingDraft[key as keyof TenantBrandingDefinition]) : "#005292"}
                      />
                    </div>
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
                  <div className={styles.sidebarPanel}>
                    {brandingDraft.brandBookAssetId
                      ? "A brand book is attached. Move the theme into review once the suggested tokens and accessibility checks look right."
                      : "Attach a brand book to anchor review and approval against a source asset."}
                  </div>
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
                        <article className={`${styles.metricCard} ${styles.staggerChild}`} key={menu.id}>
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
                  <span>Profile page</span>
                  <select className={styles.select} onChange={(event) => setProfileDraft((current) => ({ ...current, profilePageKey: event.target.value }))} value={profileDraft.profilePageKey}>
                    <option value="">Choose page</option>
                    {manifest.pages.map((page) => (
                      <option key={page.id} value={page.key}>
                        {page.title} ({page.key})
                      </option>
                    ))}
                  </select>
                </label>
                <label className={styles.formField}>
                  <span>Settings page</span>
                  <select className={styles.select} onChange={(event) => setProfileDraft((current) => ({ ...current, settingsPageKey: event.target.value }))} value={profileDraft.settingsPageKey}>
                    <option value="">Choose page</option>
                    {manifest.pages.map((page) => (
                      <option key={page.id} value={page.key}>
                        {page.title} ({page.key})
                      </option>
                    ))}
                  </select>
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
              <div className={styles.sidebarPanel}>
                Route ownership stays metadata-driven through pages. Keep menu structure and landing rules here, then confirm page-level routes in the Pages workspace.
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
                    <div className={styles.inlineList}>
                      <button className={styles.ghostButton} onClick={() => void handleTestNotificationChannel(channel.key)} type="button">
                        Test send
                      </button>
                      <button className={styles.ghostButton} onClick={() => void handleToggleNotificationChannel(channel.key, false)} type="button">
                        Disable runtime
                      </button>
                      <button className={styles.ghostButton} onClick={() => void handleToggleNotificationChannel(channel.key, true)} type="button">
                        Enable runtime
                      </button>
                    </div>
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
      </div>
    );
  }

  function renderWorkflowsWorkspace() {
    const workflowDiagnostics: WorkflowDiagnostic[] = [];
    if (!workflowDraft.name.trim()) {
      workflowDiagnostics.push({ id: "workflow-name", severity: "blocking", category: "graph", message: "Workflow name is required." });
    }
    if (!workflowDraft.triggers.length) {
      workflowDiagnostics.push({ id: "workflow-trigger", severity: "blocking", category: "trigger", message: "Add at least one trigger." });
    }
    if (!workflowDraft.nodes.length) {
      workflowDiagnostics.push({ id: "workflow-node", severity: "blocking", category: "graph", message: "Add at least one node." });
    }
    if (workflowDraft.nodes.length > 1 && workflowDraft.edges.length === 0) {
      workflowDiagnostics.push({ id: "workflow-edges", severity: "blocking", category: "graph", message: "Connect nodes with at least one edge." });
    }
    const orphanNode = workflowDraft.nodes.find((node) => !workflowDraft.edges.some((edge) => edge.sourceId === node.id || edge.targetId === node.id) && workflowDraft.nodes.length > 1);
    if (orphanNode) {
      workflowDiagnostics.push({ id: "workflow-orphan", severity: "warning", category: "graph", message: `Node "${orphanNode.label}" is disconnected from the graph.` });
    }
    const invalidConditionEdge = workflowDraft.nodes
      .filter((node) => node.type === "condition")
      .find((node) => workflowDraft.edges.some((edge) => edge.sourceId === node.id && edge.label && !["true", "false", "yes", "no", "success", "failure", "else", "match"].includes(edge.label.trim().toLowerCase())));
    if (invalidConditionEdge) {
      workflowDiagnostics.push({ id: "workflow-branch-label", severity: "warning", category: "branching", message: `Condition node "${invalidConditionEdge.label}" has a branch label outside the supported set.` });
    }
    const incompleteConditionNode = workflowDraft.nodes
      .filter((node) => node.type === "condition")
      .find((node) => {
        const labels = workflowDraft.edges
          .filter((edge) => edge.sourceId === node.id)
          .map((edge) => (edge.label ?? "").trim().toLowerCase());
        const hasPositive = labels.some((label) => ["true", "yes", "success", "match"].includes(label));
        const hasNegative = labels.some((label) => ["false", "no", "failure", "else"].includes(label));
        return !hasPositive || !hasNegative;
      });
    if (incompleteConditionNode) {
      workflowDiagnostics.push({
        id: "workflow-branch-coverage",
        severity: "warning",
        category: "branching",
        message: `Condition node "${incompleteConditionNode.label}" should expose both positive and negative branches.`,
      });
    }
    const invalidSubflowNode = workflowDraft.nodes.find(
      (node) =>
        node.type === "subflow" &&
        !manifest.workflows.some((workflow) => {
          const subflowConfig = node.config as { workflowKey: string };
          return workflow.key === subflowConfig.workflowKey || workflow.id === subflowConfig.workflowKey;
        }),
    );
    if (invalidSubflowNode) {
      workflowDiagnostics.push({ id: "workflow-subflow", severity: "blocking", category: "subflow", message: `Subflow node "${invalidSubflowNode.label}" references an unknown workflow.` });
    }
    const selfReferentialSubflow = workflowDraft.nodes.find((node) => {
      if (node.type !== "subflow") {
        return false;
      }
      const subflowConfig = node.config as { workflowKey?: string };
      return Boolean(subflowConfig.workflowKey) && (subflowConfig.workflowKey === workflowDraft.key || subflowConfig.workflowKey === selectedWorkflow?.id);
    });
    if (selfReferentialSubflow) {
      workflowDiagnostics.push({
        id: "workflow-subflow-self",
        severity: "blocking",
        category: "subflow",
        message: `Subflow node "${selfReferentialSubflow.label}" cannot reference the workflow currently being edited.`,
      });
    }
    const rootNodes = workflowDraft.nodes.filter((node) => !workflowDraft.edges.some((edge) => edge.targetId === node.id));
    if (workflowDraft.nodes.length > 1 && rootNodes.length > 1) {
      workflowDiagnostics.push({
        id: "workflow-multiple-roots",
        severity: "warning",
        category: "graph",
        message: "The graph has multiple root nodes. Consider a single entry path or explicit trigger fan-out.",
      });
    }
    const reachableNodeIds = new Set<string>();
    const stack = [...rootNodes.map((node) => node.id)];
    while (stack.length > 0) {
      const nextNodeId = stack.pop();
      if (!nextNodeId || reachableNodeIds.has(nextNodeId)) {
        continue;
      }
      reachableNodeIds.add(nextNodeId);
      workflowDraft.edges
        .filter((edge) => edge.sourceId === nextNodeId)
        .forEach((edge) => {
          if (!reachableNodeIds.has(edge.targetId)) {
            stack.push(edge.targetId);
          }
        });
    }
    const unreachableNode = workflowDraft.nodes.find((node) => !reachableNodeIds.has(node.id));
    if (unreachableNode) {
      workflowDiagnostics.push({
        id: "workflow-unreachable",
        severity: "warning",
        category: "graph",
        message: `Node "${unreachableNode.label}" is not reachable from any root path.`,
      });
    }
    const hasSavedTestCase = manifest.workflowTests.some(
      (testCase: WorkflowTestCaseDefinition) => testCase.workflowKey === (workflowDraft.key || selectedWorkflow?.key),
    );
    if (workflowDraft.nodes.length > 0 && !hasSavedTestCase) {
      workflowDiagnostics.push({
        id: "workflow-test-case",
        severity: "info",
        category: "testing",
        message: "Save at least one draft test case before publishing this workflow.",
      });
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
          <div className={styles.sidebarSection}>
            <p className={styles.sidebarLabel}>Templates</p>
            <div className={styles.listStack}>
              {manifest.workflowTemplates.length === 0 ? (
                <div className={styles.sidebarPanel}>Save a workflow as a reusable template once it has a stable shape.</div>
              ) : (
                manifest.workflowTemplates.slice(0, 6).map((template: WorkflowTemplateDefinition) => (
                  <button
                    className={`${styles.listItem} ${styles.staggerChild}`}
                    key={template.id}
                    onClick={() => {
                      setWorkflowDraft(structuredClone(template.workflow));
                      setSelectedWorkflowId("");
                      setSelectedWorkflowNodeId(template.workflow.nodes[0]?.id ?? "");
                    }}
                    type="button"
                  >
                    <span>{template.name}</span>
                    <small>{template.source}</small>
                  </button>
                ))
              )}
            </div>
          </div>
          <div className={styles.listStack}>
            {manifest.workflows.map((workflow) => (
              <button
                className={`${workflow.id === selectedWorkflow?.id ? styles.activeListItem : styles.listItem} ${styles.staggerChild}`}
                data-workflow-list-id={workflow.id}
                key={workflow.id}
                onClick={() => selectWorkflow(workflow)}
                type="button"
              >
                <span>{workflow.name}</span>
                <div className={styles.listItemMeta}>
                  <span className={styles.metaBadge}>{workflow.nodes.length} nodes</span>
                  <span className={styles.metaBadge}>{workflow.edges.length} edges</span>
                  <span className={styles.metaBadge}>{workflow.status}</span>
                </div>
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
                    <button className={styles.secondaryButton} onClick={() => void handleSaveWorkflowTemplate()} type="button">
                      Save as template
                    </button>
                    <button className={styles.secondaryButton} onClick={() => void handleSaveWorkflowTestCase()} type="button">
                      Save test case
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
                    {workflowDiagnostics.length === 0
                      ? "Graph passes the current structural checks."
                      : workflowDiagnostics.map((diagnostic) => diagnostic.message).join(" ")}
                  </div>
                  {manifest.workflowTests.filter((testCase: WorkflowTestCaseDefinition) => testCase.workflowKey === (workflowDraft.key || selectedWorkflow?.key)).length > 0 ? (
                    <div className={styles.listStack}>
                      {manifest.workflowTests
                        .filter((testCase: WorkflowTestCaseDefinition) => testCase.workflowKey === (workflowDraft.key || selectedWorkflow?.key))
                        .slice(0, 4)
                        .map((testCase: WorkflowTestCaseDefinition) => (
                          <button
                            className={styles.listItem}
                            key={testCase.id}
                            onClick={() => setWorkflowTestPayload(JSON.stringify(testCase.payload, null, 2))}
                            type="button"
                          >
                            <span>{testCase.name}</span>
                            <small>{testCase.expectedStatus ?? "No expected status"}</small>
                          </button>
                        ))}
                    </div>
                  ) : null}
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
                        {"workflowKey" in selectedWorkflowNode.config ? (
                          <label className={styles.formField}>
                            <span>Subflow workflow</span>
                            <select
                              className={styles.select}
                              onChange={(event) =>
                                updateWorkflowNode(selectedWorkflowNode.id, (current) => ({
                                  ...current,
                                  config: { ...current.config, workflowKey: event.target.value },
                                }))
                              }
                              value={selectedWorkflowNode.config.workflowKey}
                            >
                              <option value="">Choose workflow</option>
                              {manifest.workflows.map((workflow) => (
                                <option key={workflow.id} value={workflow.key}>
                                  {workflow.name}
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
                className={`${agent.id === selectedAgent?.id ? styles.activeListItem : styles.listItem} ${styles.staggerChild}`}
                data-agent-list-id={agent.id}
                key={agent.id}
                onClick={() => selectAgent(agent)}
                type="button"
              >
                <span>{agent.name}</span>
                <div className={styles.listItemMeta}>
                  <span className={styles.metaBadge}>{agent.scope}</span>
                  <span className={styles.metaBadge}>{agent.zeroRetentionRequired ? "zero retention" : "standard"}</span>
                  {agent.modelProviderId ? <span className={styles.metaBadge}>{manifest.modelProviders.find((p) => p.id === agent.modelProviderId)?.name ?? "provider"}</span> : null}
                </div>
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
              <details className={styles.formSection} open>
                <summary>
                  <svg className={styles.formSectionChevron} fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" viewBox="0 0 10 10"><path d="M3 1l4 4-4 4" /></svg>
                  Identity
                </summary>
                <div className={styles.formSectionContent}>
                  <div className={styles.formGrid}>
                    <label className={styles.formField}>
                      <span>Name</span>
                      <input className={styles.input} onChange={(event) => setAgentDraft((current) => ({ ...current, name: event.target.value }))} value={agentDraft.name} />
                    </label>
                    <label className={styles.formField}>
                      <span>Key</span>
                      <input className={styles.input} onChange={(event) => setAgentDraft((current) => ({ ...current, key: event.target.value }))} value={agentDraft.key} />
                      <small className={styles.formFieldHelp}>Unique lowercase_snake_case identifier</small>
                    </label>
                    <label className={styles.formField}>
                      <span>Scope</span>
                      <select className={styles.select} onChange={(event) => setAgentDraft((current) => ({ ...current, scope: event.target.value as AgentDefinition["scope"] }))} value={agentDraft.scope}>
                        <option value="workspace">Workspace</option>
                        <option value="node">Node</option>
                      </select>
                      <small className={styles.formFieldHelp}>Workspace agents operate across all objects. Node agents run at a single workflow step.</small>
                    </label>
                    <label className={styles.formFieldSpan}>
                      <span>Description</span>
                      <textarea className={styles.textarea} onChange={(event) => setAgentDraft((current) => ({ ...current, description: event.target.value }))} value={agentDraft.description ?? ""} />
                    </label>
                  </div>
                </div>
              </details>
              <details className={styles.formSection} open>
                <summary>
                  <svg className={styles.formSectionChevron} fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" viewBox="0 0 10 10"><path d="M3 1l4 4-4 4" /></svg>
                  Model and prompt
                </summary>
                <div className={styles.formSectionContent}>
                  <div className={styles.formGrid}>
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
                    <label className={styles.formField}>
                      <span>Budget (USD)</span>
                      <input className={styles.input} min="0" onChange={(event) => setAgentDraft((current) => ({ ...current, costBudgetUsd: Number(event.target.value) || 0 }))} step="0.01" type="number" value={agentDraft.costBudgetUsd ?? 0} />
                      <small className={styles.formFieldHelp}>Maximum spend per execution before the agent is paused.</small>
                    </label>
                    <label className={styles.formFieldSpan}>
                      <span>Prompt</span>
                      <textarea className={styles.textarea} onChange={(event) => setAgentDraft((current) => ({ ...current, prompt: event.target.value }))} value={agentDraft.prompt} />
                    </label>
                    <label className={styles.formField}>
                      <span>Output schema</span>
                      <textarea className={styles.textarea} onChange={(event) => setAgentDraft((current) => ({ ...current, outputSchema: event.target.value }))} value={agentDraft.outputSchema ?? ""} />
                    </label>
                  </div>
                </div>
              </details>
              <details className={styles.formSection} open>
                <summary>
                  <svg className={styles.formSectionChevron} fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" viewBox="0 0 10 10"><path d="M3 1l4 4-4 4" /></svg>
                  Policy and safety
                </summary>
                <div className={styles.formSectionContent}>
                  <div className={styles.inlineList}>
                    <label className={styles.checkboxField}>
                      <input checked={agentDraft.zeroRetentionRequired} onChange={(event) => setAgentDraft((current) => ({ ...current, zeroRetentionRequired: event.target.checked }))} type="checkbox" />
                      <span>Zero retention required</span>
                      <small className={styles.formFieldHelp}>Provider must not store request or response data</small>
                    </label>
                    <label className={styles.checkboxField}>
                      <input checked={agentDraft.approvalPolicy?.required ?? false} onChange={(event) => setAgentDraft((current) => ({ ...current, approvalPolicy: { ...(current.approvalPolicy ?? { approverRole: "BUILDER_ADMIN", notes: "" }), required: event.target.checked } }))} type="checkbox" />
                      <span>Approval required</span>
                      <small className={styles.formFieldHelp}>Human sign-off before each run</small>
                    </label>
                  </div>
                </div>
              </details>
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
                    renderEmptyState("agent", "No activity yet.", "Agent previews and future runtime calls will appear here.")
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
            </>
          ) : null}

          {activeTab === 2 ? (
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
                <div className={styles.sidebarPanel}>
                  Governance: {agentDraft.approvalPolicy?.required ? "approval required" : "allowed"} ·
                  {agentDraft.outputSchema?.trim() ? " schema enforced" : " schema optional"} ·
                  {agentDraft.zeroRetentionRequired ? " zero retention" : " standard retention"}
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

  function renderControlTowerDetailPanel() {
    if (!controlTowerSelection || !controlTowerDetail) {
      return null;
    }

    function humanizeToken(value: string): string {
      return value
        .replace(/[_-]+/g, " ")
        .replace(/\s+/g, " ")
        .trim()
        .replace(/\b\w/g, (character) => character.toUpperCase());
    }

    function formatInlineValue(value: unknown): string {
      if (value == null) {
        return "No value";
      }
      if (typeof value === "string") {
        return value.length > 120 ? `${value.slice(0, 117)}...` : value;
      }
      if (typeof value === "number" || typeof value === "boolean") {
        return String(value);
      }
      if (Array.isArray(value)) {
        return value.length === 0 ? "No items" : `${value.length} item${value.length === 1 ? "" : "s"}`;
      }
      const entries = Object.entries(value as Record<string, unknown>);
      if (entries.length === 0) {
        return "Empty object";
      }
      return entries
        .slice(0, 3)
        .map(([key, nestedValue]) => `${humanizeToken(key)}: ${formatInlineValue(nestedValue)}`)
        .join(" · ");
    }

    function renderStructuredPanel(label: string, value: unknown) {
      if (value == null) {
        return null;
      }

      if (Array.isArray(value)) {
        return (
          <article className={styles.scopeCard}>
            <div className={styles.sectionHeader}>
              <div>
                <p className={styles.cardEyebrow}>{label}</p>
                <h3>{value.length} entries</h3>
              </div>
            </div>
            <div className={styles.listStack}>
              {value.length === 0 ? <div className={styles.sidebarPanel}>No entries recorded.</div> : null}
              {value.map((entry, index) => (
                <article className={styles.auditRow} key={`${label}-${index}`}>
                  <div>
                    <strong>{`Entry ${index + 1}`}</strong>
                    <p>{formatInlineValue(entry)}</p>
                  </div>
                </article>
              ))}
            </div>
          </article>
        );
      }

      if (typeof value === "object") {
        return (
          <article className={styles.scopeCard}>
            <div className={styles.sectionHeader}>
              <div>
                <p className={styles.cardEyebrow}>{label}</p>
                <h3>Structured snapshot</h3>
              </div>
            </div>
            <div className={styles.listStack}>
              {Object.entries(value as Record<string, unknown>).map(([key, nestedValue]) => (
                <article className={styles.auditRow} key={`${label}-${key}`}>
                  <div>
                    <strong>{humanizeToken(key)}</strong>
                    <p>{formatInlineValue(nestedValue)}</p>
                  </div>
                </article>
              ))}
            </div>
          </article>
        );
      }

      return (
        <article className={styles.scopeCard}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>{label}</p>
              <h3>Snapshot</h3>
            </div>
          </div>
          <div className={styles.sidebarPanel}>{formatInlineValue(value)}</div>
        </article>
      );
    }

    function renderTimeline(label: string, entries: Array<Record<string, unknown>>) {
      if (entries.length === 0) {
        return null;
      }
      return (
        <article className={styles.scopeCard}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>{label}</p>
              <h3>Timeline</h3>
            </div>
          </div>
          <div className={styles.listStack}>
            {entries.map((entry, index) => (
              <article className={styles.auditRow} key={`${label}-${index}`}>
                <div>
                  <strong>{String(entry.message ?? entry.summary ?? entry.lifecycleStage ?? entry.status ?? "Event")}</strong>
                  <p>{String(entry.level ?? entry.provider ?? entry.responseSummary ?? entry.errorMessage ?? "")}</p>
                </div>
                <span>{formatPlatformDateTime(String(entry.at ?? entry.createdAt ?? ""))}</span>
              </article>
            ))}
          </div>
        </article>
      );
    }

    function renderRelationList(input: {
      label: string;
      title: string;
      emptyLabel: string;
      items: Array<{
        id: string;
        primary: string;
        secondary?: string;
        tags?: string[];
        inspectType?: ControlTowerSelectionType;
      }>;
    }) {
      return (
        <article className={styles.scopeCard}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>{input.label}</p>
              <h3>{input.title}</h3>
            </div>
          </div>
          <div className={styles.listStack}>
            {input.items.length === 0 ? <div className={styles.sidebarPanel}>{input.emptyLabel}</div> : null}
            {input.items.map((item) => (
              <article className={styles.auditRow} key={`${input.label}-${item.id}`}>
                <div>
                  <strong>{item.primary}</strong>
                  <p>{item.secondary ?? item.id}</p>
                  {item.tags && item.tags.length > 0 ? (
                    <div className={styles.inlineList}>
                      {item.tags.map((tag) => (
                        <span className={styles.inlineTag} key={`${item.id}-${tag}`}>
                          {tag}
                        </span>
                      ))}
                    </div>
                  ) : null}
                </div>
                {item.inspectType ? (
                  <button className={styles.ghostButton} onClick={() => void handleInspectControlTowerDetail(item.inspectType!, item.id)} type="button">
                    Inspect
                  </button>
                ) : null}
              </article>
            ))}
          </div>
        </article>
      );
    }

    function renderAgentTracePanel(trace: AgentRunRecord["trace"]) {
      if (!trace) {
        return null;
      }

      return (
        <article className={styles.scopeCard}>
          <div className={styles.sectionHeader}>
            <div>
              <p className={styles.cardEyebrow}>Execution trace</p>
              <h3>
                {trace.providerKey} · {trace.providerModel}
              </h3>
            </div>
          </div>
          <div className={styles.sidebarPanel}>
            {trace.outputValidationPassed === undefined ? "Schema optional" : trace.outputValidationPassed ? "Schema passed" : "Schema failed"}
            {trace.handoffWorkflowKey ? ` · Handoff ${trace.handoffWorkflowKey}` : ""}
          </div>
          {trace.promptBlockSummary.length > 0 ? (
            <div className={styles.listStack}>
              {trace.promptBlockSummary.map((block) => (
                <article className={styles.auditRow} key={block.id}>
                  <div>
                    <strong>{block.label}</strong>
                    <p>{humanizeToken(block.kind)}</p>
                  </div>
                </article>
              ))}
            </div>
          ) : null}
          {trace.policyDecisions.length > 0 ? (
            <div className={styles.listStack}>
              {trace.policyDecisions.map((decision, index) => (
                <article className={styles.auditRow} key={`decision-${index}`}>
                  <div>
                    <strong>{`Policy decision ${index + 1}`}</strong>
                    <p>{decision}</p>
                  </div>
                </article>
              ))}
            </div>
          ) : null}
          {renderTimeline("Stage timeline", trace.stages as unknown as Array<Record<string, unknown>>)}
        </article>
      );
    }

    const workflowDetail = controlTowerSelection.type === "workflow-run" ? (controlTowerDetail as unknown as WorkflowRunDetail) : null;
    const agentDetail = controlTowerSelection.type === "agent-run" ? (controlTowerDetail as unknown as AgentRunDetail) : null;
    const deliveryDetail =
      controlTowerSelection.type === "delivery"
        ? (controlTowerDetail as unknown as {
            delivery: NotificationDeliveryRecord;
            attempts: NotificationAttemptRecord[];
            channelHealth: NotificationChannelHealth | null;
          })
        : null;
    const alertDetail =
      controlTowerSelection.type === "alert"
        ? (controlTowerDetail as unknown as {
            alert: PlatformAlertRecord;
            relatedDeliveries: NotificationDeliveryRecord[];
            relatedWorkflowRuns: PlatformWorkflowRunRecord[];
            relatedAgentRuns: AgentRunRecord[];
          })
        : null;
    const approvalDetail =
      controlTowerSelection.type === "approval"
        ? (controlTowerDetail as unknown as {
            task: PlatformApprovalTaskRecord;
            workflowRun: PlatformWorkflowRunRecord | null;
            agentRun: AgentRunRecord | null;
          })
        : null;
    const deadLetterDetail = controlTowerSelection.type === "dead-letter" ? (controlTowerDetail as unknown as DeadLetterRecord) : null;

    return (
      <section className={styles.panelWide}>
        <div className={styles.sectionHeader}>
          <div>
            <p className={styles.cardEyebrow}>Drill-down</p>
            <h2>
              {controlTowerSelection.type} · {controlTowerSelection.id}
            </h2>
          </div>
        </div>
        {workflowDetail ? (
          <div className={styles.scopeGrid}>
            <article className={styles.scopeCard}>
              <p className={styles.cardEyebrow}>Workflow run</p>
              <h3>{workflowDetail.run.workflowKey}</h3>
              <div className={styles.inlineList}>
                <span className={styles.inlineTag}>{workflowDetail.run.status}</span>
                {workflowDetail.run.pauseReason ? <span className={styles.inlineTag}>{workflowDetail.run.pauseReason}</span> : null}
                {workflowDetail.run.subflowRunId ? <span className={styles.inlineTag}>subflow linked</span> : null}
              </div>
              <div className={styles.sidebarPanel}>
                Created {formatPlatformDateTime(workflowDetail.run.createdAt)}
                {workflowDetail.run.startedAt ? ` · Started ${formatPlatformDateTime(workflowDetail.run.startedAt)}` : ""}
                {workflowDetail.run.finishedAt ? ` · Finished ${formatPlatformDateTime(workflowDetail.run.finishedAt)}` : ""}
              </div>
              <div className={styles.sidebarPanel}>
                Related cost ${workflowDetail.relatedAgentRuns.reduce((sum, run) => sum + run.costUsd, 0).toFixed(4)}
                {workflowDetail.run.nextRetryAt ? ` · Next retry ${formatPlatformDateTime(workflowDetail.run.nextRetryAt)}` : ""}
              </div>
              <div className={styles.inlineList}>
                <button className={styles.secondaryButton} onClick={() => void handleReplayWorkflowRun(workflowDetail.run.id)} type="button">
                  Replay run
                </button>
              </div>
            </article>
            {renderStructuredPanel("Input snapshot", workflowDetail.run.input)}
            {renderStructuredPanel("Output snapshot", workflowDetail.run.output)}
            {renderTimeline("Run logs", workflowDetail.run.logs)}
            {renderRelationList({
              label: "Run lineage",
              title: "Parent and child runs",
              emptyLabel: "No parent or child runs linked.",
              items: [
                ...(workflowDetail.parentRun
                  ? [
                      {
                        id: workflowDetail.parentRun.id,
                        primary: workflowDetail.parentRun.workflowKey,
                        secondary: `Parent run · ${workflowDetail.parentRun.status}`,
                        tags: ["parent"],
                        inspectType: "workflow-run" as const,
                      },
                    ]
                  : []),
                ...workflowDetail.childRuns.map((run) => ({
                  id: run.id,
                  primary: run.workflowKey,
                  secondary: `Child run · ${run.status}`,
                  tags: run.pauseReason ? [run.pauseReason] : [],
                  inspectType: "workflow-run" as const,
                })),
              ],
            })}
            {renderRelationList({
              label: "Related runs",
              title: "Agent executions",
              emptyLabel: "No agent runs linked to this workflow.",
              items: workflowDetail.relatedAgentRuns.map((run) => ({
                id: run.id,
                primary: run.agentKey,
                secondary: `${run.status} · ${run.modelProviderKey} · $${run.costUsd.toFixed(4)}`,
                tags: [run.runMode, run.approvalStatus],
                inspectType: "agent-run" as const,
              })),
            })}
            {renderRelationList({
              label: "Operational links",
              title: "Deliveries and approvals",
              emptyLabel: "No deliveries or approvals linked.",
              items: [
                ...workflowDetail.deliveries.map((delivery) => ({
                  id: delivery.id,
                  primary: delivery.ruleKey,
                  secondary: `${delivery.status} · ${delivery.channelKey}`,
                  tags: [delivery.severity],
                  inspectType: "delivery" as const,
                })),
                ...workflowDetail.approvals.map((task) => ({
                  id: task.id,
                  primary: task.nodeLabel,
                  secondary: `${task.status} · ${task.approverRole}`,
                  tags: [task.taskType ?? "workflow"],
                  inspectType: "approval" as const,
                })),
              ],
            })}
            {renderRelationList({
              label: "Dead letters",
              title: "Terminal failures",
              emptyLabel: "No dead letters recorded.",
              items: workflowDetail.deadLetters.map((entry) => ({
                id: entry.id,
                primary: entry.type,
                secondary: entry.reason,
                tags: ["dead letter"],
                inspectType: "dead-letter" as const,
              })),
            })}
          </div>
        ) : null}
        {agentDetail ? (
          <div className={styles.scopeGrid}>
            <article className={styles.scopeCard}>
              <p className={styles.cardEyebrow}>Agent run</p>
              <h3>{agentDetail.run.agentKey}</h3>
              <div className={styles.inlineList}>
                <span className={styles.inlineTag}>{agentDetail.run.status}</span>
                <span className={styles.inlineTag}>{agentDetail.run.runMode}</span>
                <span className={styles.inlineTag}>{agentDetail.run.approvalStatus}</span>
              </div>
              <div className={styles.sidebarPanel}>
                Provider {agentDetail.run.modelProviderKey} · ${agentDetail.run.costUsd.toFixed(4)} · {agentDetail.run.tokensIn + agentDetail.run.tokensOut} tokens
              </div>
              {agentDetail.run.schemaValidation ? (
                <div className={styles.sidebarPanel}>
                  Schema: {agentDetail.run.schemaValidation.passed ? "passed" : "failed"} · {agentDetail.run.schemaValidation.summary}
                </div>
              ) : null}
              <div className={styles.inlineList}>
                {agentDetail.run.status === "blocked" && agentDetail.run.approvalStatus === "approved" ? (
                  <button className={styles.secondaryButton} onClick={() => void handleResumeAgentRun(agentDetail.run.id)} type="button">
                    Resume
                  </button>
                ) : null}
                {agentDetail.parentWorkflowRun ? <span className={styles.inlineTag}>workflow linked</span> : null}
                {agentDetail.handoffWorkflowRun ? <span className={styles.inlineTag}>handoff linked</span> : null}
              </div>
            </article>
            {renderStructuredPanel("Input snapshot", agentDetail.run.input)}
            {renderStructuredPanel("Output snapshot", agentDetail.run.output)}
            {renderTimeline("Run logs", agentDetail.run.logs)}
            {renderAgentTracePanel(agentDetail.run.trace)}
            {renderRelationList({
              label: "Workflow links",
              title: "Parent and handoff workflows",
              emptyLabel: "No workflow lineage linked.",
              items: [
                ...(agentDetail.parentWorkflowRun
                  ? [
                      {
                        id: agentDetail.parentWorkflowRun.id,
                        primary: agentDetail.parentWorkflowRun.workflowKey,
                        secondary: `Parent workflow · ${agentDetail.parentWorkflowRun.status}`,
                        tags: ["parent"],
                        inspectType: "workflow-run" as const,
                      },
                    ]
                  : []),
                ...(agentDetail.handoffWorkflowRun
                  ? [
                      {
                        id: agentDetail.handoffWorkflowRun.id,
                        primary: agentDetail.handoffWorkflowRun.workflowKey,
                        secondary: `Handoff workflow · ${agentDetail.handoffWorkflowRun.status}`,
                        tags: ["handoff"],
                        inspectType: "workflow-run" as const,
                      },
                    ]
                  : []),
              ],
            })}
            {renderRelationList({
              label: "Operational links",
              title: "Deliveries and approvals",
              emptyLabel: "No deliveries or approvals linked.",
              items: [
                ...agentDetail.deliveries.map((delivery) => ({
                  id: delivery.id,
                  primary: delivery.ruleKey,
                  secondary: `${delivery.status} · ${delivery.channelKey}`,
                  tags: [delivery.severity],
                  inspectType: "delivery" as const,
                })),
                ...agentDetail.approvals.map((task) => ({
                  id: task.id,
                  primary: task.nodeLabel,
                  secondary: `${task.status} · ${task.approverRole}`,
                  tags: [task.taskType ?? "agent"],
                  inspectType: "approval" as const,
                })),
              ],
            })}
          </div>
        ) : null}
        {deliveryDetail ? (
          <div className={styles.scopeGrid}>
            <article className={styles.scopeCard}>
              <p className={styles.cardEyebrow}>Notification delivery</p>
              <h3>{deliveryDetail.delivery.ruleKey}</h3>
              <div className={styles.inlineList}>
                <span className={styles.inlineTag}>{deliveryDetail.delivery.status}</span>
                <span className={styles.inlineTag}>{deliveryDetail.delivery.channelKey}</span>
                <span className={styles.inlineTag}>{deliveryDetail.delivery.severity}</span>
              </div>
              <div className={styles.sidebarPanel}>
                Attempts {deliveryDetail.delivery.attemptCount ?? 0}/{deliveryDetail.delivery.maxAttempts ?? "—"}
                {deliveryDetail.delivery.nextRetryAt ? ` · Next retry ${formatPlatformDateTime(deliveryDetail.delivery.nextRetryAt)}` : ""}
              </div>
              {deliveryDetail.channelHealth ? (
                <div className={styles.sidebarPanel}>
                  Channel health: {deliveryDetail.channelHealth.status} · success rate {(deliveryDetail.channelHealth.successRate * 100).toFixed(0)}%
                </div>
              ) : null}
              <div className={styles.inlineList}>
                <button className={styles.secondaryButton} disabled={deliveryDetail.delivery.status === "sent"} onClick={() => void handleRetryDelivery(deliveryDetail.delivery.id)} type="button">
                  Retry delivery
                </button>
              </div>
            </article>
            {renderStructuredPanel("Resolved payload", deliveryDetail.delivery.resolvedPayload)}
            {renderTimeline("Delivery attempts", deliveryDetail.attempts as unknown as Array<Record<string, unknown>>)}
            {deliveryDetail.channelHealth ? renderStructuredPanel("Channel health", deliveryDetail.channelHealth) : null}
          </div>
        ) : null}
        {alertDetail ? (
          <div className={styles.scopeGrid}>
            <article className={styles.scopeCard}>
              <p className={styles.cardEyebrow}>Alert</p>
              <h3>{alertDetail.alert.title}</h3>
              <div className={styles.inlineList}>
                <span className={styles.inlineTag}>{alertDetail.alert.category}</span>
                <span className={styles.inlineTag}>{alertDetail.alert.severity}</span>
              </div>
              <div className={styles.sidebarPanel}>{alertDetail.alert.summary}</div>
              <div className={styles.inlineList}>
                <button className={styles.secondaryButton} disabled={Boolean(alertDetail.alert.acknowledgedAt)} onClick={() => void handleAcknowledgeAlert(alertDetail.alert.id)} type="button">
                  Acknowledge
                </button>
              </div>
            </article>
            {renderRelationList({
              label: "Related deliveries",
              title: "Delivery impact",
              emptyLabel: "No deliveries linked to this alert.",
              items: alertDetail.relatedDeliveries.map((delivery) => ({
                id: delivery.id,
                primary: delivery.ruleKey,
                secondary: `${delivery.status} · ${delivery.channelKey}`,
                tags: [delivery.severity],
                inspectType: "delivery" as const,
              })),
            })}
            {renderRelationList({
              label: "Related workflows",
              title: "Workflow runs",
              emptyLabel: "No workflow runs linked to this alert.",
              items: alertDetail.relatedWorkflowRuns.map((run) => ({
                id: run.id,
                primary: run.workflowKey,
                secondary: run.status,
                tags: run.pauseReason ? [run.pauseReason] : [],
                inspectType: "workflow-run" as const,
              })),
            })}
            {renderRelationList({
              label: "Related agents",
              title: "Agent runs",
              emptyLabel: "No agent runs linked to this alert.",
              items: alertDetail.relatedAgentRuns.map((run) => ({
                id: run.id,
                primary: run.agentKey,
                secondary: `${run.status} · ${run.modelProviderKey}`,
                tags: [run.runMode, run.approvalStatus],
                inspectType: "agent-run" as const,
              })),
            })}
          </div>
        ) : null}
        {approvalDetail ? (
          <div className={styles.scopeGrid}>
            <article className={styles.scopeCard}>
              <p className={styles.cardEyebrow}>Approval task</p>
              <h3>{approvalDetail.task.nodeLabel}</h3>
              <div className={styles.inlineList}>
                <span className={styles.inlineTag}>{approvalDetail.task.status}</span>
                <span className={styles.inlineTag}>{approvalDetail.task.approverRole}</span>
              </div>
              {approvalDetail.task.instructions ? <div className={styles.sidebarPanel}>{approvalDetail.task.instructions}</div> : null}
              <div className={styles.inlineList}>
                <button className={styles.secondaryButton} disabled={approvalDetail.task.status !== "pending"} onClick={() => void handleResolveApprovalTask(approvalDetail.task.id, "approved")} type="button">
                  Approve
                </button>
                <button className={styles.ghostButtonDanger} disabled={approvalDetail.task.status !== "pending"} onClick={() => void handleResolveApprovalTask(approvalDetail.task.id, "rejected")} type="button">
                  Reject
                </button>
              </div>
            </article>
            {renderRelationList({
              label: "Workflow link",
              title: "Workflow run",
              emptyLabel: "No workflow run linked.",
              items: approvalDetail.workflowRun
                ? [
                    {
                      id: approvalDetail.workflowRun.id,
                      primary: approvalDetail.workflowRun.workflowKey,
                      secondary: approvalDetail.workflowRun.status,
                      tags: [approvalDetail.task.taskType ?? "workflow"],
                      inspectType: "workflow-run" as const,
                    },
                  ]
                : [],
            })}
            {renderRelationList({
              label: "Agent link",
              title: "Agent run",
              emptyLabel: "No agent run linked.",
              items: approvalDetail.agentRun
                ? [
                    {
                      id: approvalDetail.agentRun.id,
                      primary: approvalDetail.agentRun.agentKey,
                      secondary: `${approvalDetail.agentRun.status} · ${approvalDetail.agentRun.modelProviderKey}`,
                      tags: [approvalDetail.agentRun.runMode, approvalDetail.agentRun.approvalStatus],
                      inspectType: "agent-run" as const,
                    },
                  ]
                : [],
            })}
          </div>
        ) : null}
        {deadLetterDetail ? (
          <div className={styles.scopeGrid}>
            <article className={styles.scopeCard}>
              <p className={styles.cardEyebrow}>Dead letter</p>
              <h3>{deadLetterDetail.type}</h3>
              <div className={styles.sidebarPanel}>
                {deadLetterDetail.reason} · Captured {formatPlatformDateTime(deadLetterDetail.createdAt)}
              </div>
            </article>
            {renderStructuredPanel("Payload", deadLetterDetail.payload)}
          </div>
        ) : null}
      </section>
    );
  }

  function renderControlTowerWorkspace() {
    return (
      <div className={styles.workspaceGrid}>
        <section className={styles.panelWide}>
          <div className={styles.sectionHeader}>
            <button className={styles.filterToggleRow} onClick={() => setControlTowerFiltersExpanded((prev) => !prev)} type="button">
              <div>
                <p className={styles.cardEyebrow}>Filters</p>
                <h2>Operator scope{activeFilterCount > 0 ? <span className={styles.filterBadge} style={{ marginLeft: 10, verticalAlign: "middle" }}>{activeFilterCount}</span> : null}</h2>
              </div>
              <svg className={styles.navGroupChevron} fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" viewBox="0 0 10 10" style={{ transform: controlTowerFiltersExpanded ? "rotate(90deg)" : "rotate(0deg)", transition: "transform 0.2s" }}><path d="M3 1l4 4-4 4" /></svg>
            </button>
            {activeFilterCount > 0 ? (
              <button className={styles.filterClearButton} onClick={() => setControlTowerFilters({ status: "all", workflowKey: "", agentKey: "", severity: "all", fromDate: "", toDate: "" })} type="button">Clear all</button>
            ) : null}
          </div>
          <div className={controlTowerFiltersExpanded ? styles.filterBody : `${styles.filterBody} ${styles.filterBodyCollapsed}`}>
            <div className={styles.formGrid}>
              <label className={styles.formField}>
                <span>Status</span>
                <select className={styles.select} onChange={(event) => setControlTowerFilters((current) => ({ ...current, status: event.target.value }))} value={controlTowerFilters.status}>
                  <option value="all">All</option>
                  <option value="QUEUED">Queued</option>
                  <option value="RUNNING">Running</option>
                  <option value="PAUSED">Paused</option>
                  <option value="SUCCEEDED">Succeeded</option>
                  <option value="FAILED">Failed</option>
                  <option value="queued">Queued agent</option>
                  <option value="running">Running agent</option>
                  <option value="blocked">Blocked agent</option>
                  <option value="succeeded">Succeeded agent</option>
                  <option value="failed">Failed agent</option>
                  <option value="retrying">Retrying delivery</option>
                  <option value="exhausted">Exhausted delivery</option>
                </select>
              </label>
              <label className={styles.formField}>
                <span>Workflow key</span>
                <input className={styles.input} onChange={(event) => setControlTowerFilters((current) => ({ ...current, workflowKey: event.target.value }))} value={controlTowerFilters.workflowKey} />
              </label>
              <label className={styles.formField}>
                <span>Agent key</span>
                <input className={styles.input} onChange={(event) => setControlTowerFilters((current) => ({ ...current, agentKey: event.target.value }))} value={controlTowerFilters.agentKey} />
              </label>
              <label className={styles.formField}>
                <span>Severity</span>
                <select className={styles.select} onChange={(event) => setControlTowerFilters((current) => ({ ...current, severity: event.target.value }))} value={controlTowerFilters.severity}>
                  <option value="all">All</option>
                  <option value="info">Info</option>
                  <option value="success">Success</option>
                  <option value="warning">Warning</option>
                  <option value="critical">Critical</option>
                </select>
              </label>
              <label className={styles.formField}>
                <span>From</span>
                <input className={styles.input} onChange={(event) => setControlTowerFilters((current) => ({ ...current, fromDate: event.target.value }))} type="date" value={controlTowerFilters.fromDate} />
              </label>
              <label className={styles.formField}>
                <span>To</span>
                <input className={styles.input} onChange={(event) => setControlTowerFilters((current) => ({ ...current, toDate: event.target.value }))} type="date" value={controlTowerFilters.toDate} />
              </label>
            </div>
          </div>
        </section>
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
                {isControlTowerLoading ? (
                  renderSkeleton(4)
                ) : filteredWorkflowRuns.length === 0 ? (
                  renderEmptyState("workflow", "No workflow runs yet.", "Publish a workflow and queue a run to see execution data here.", "Refresh", () => void refreshControlTower())
                ) : (
                  filteredWorkflowRuns.slice(0, controlTowerListLimits.workflowRuns).map((run) => (
                    <article className={`${styles.auditRow} ${styles.staggerChild}`} key={run.id}>
                      <div>
                        <strong>{run.workflowKey}</strong>
                        <p>
                          {run.status} · {run.logs.length} log events
                          {run.pauseReason ? ` · ${run.pauseReason}` : ""}
                        </p>
                      </div>
                      <div className={styles.inlineList}>
                        <span>{formatPlatformDateTime(run.createdAt)}</span>
                        <button className={styles.ghostButton} onClick={() => void handleInspectControlTowerDetail("workflow-run", run.id)} type="button">
                          Inspect
                        </button>
                        <button className={styles.ghostButton} onClick={() => void handleReplayWorkflowRun(run.id)} type="button">
                          Replay
                        </button>
                      </div>
                    </article>
                  ))
                )}
                {filteredWorkflowRuns.length > controlTowerListLimits.workflowRuns ? (
                  <button className={styles.showMoreButton} onClick={() => setControlTowerListLimits((prev) => ({ ...prev, workflowRuns: prev.workflowRuns + 24 }))} type="button">
                    Show more ({filteredWorkflowRuns.length - controlTowerListLimits.workflowRuns} remaining)
                  </button>
                ) : null}
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
                {isControlTowerLoading ? renderSkeleton(3, true) : filteredAgentRuns.slice(0, 6).map((run) => (
                  <article className={styles.workflowCard} key={`detail-${run.id}`}>
                    <strong>{run.agentKey}</strong>
                    <p>{run.logs[0]?.message ? String(run.logs[0].message) : "Simulation completed."}</p>
                    <div className={styles.inlineList}>
                      <span className={styles.inlineTag}>{run.modelProviderKey}</span>
                      <span className={styles.inlineTag}>{run.runMode}</span>
                      <span className={styles.inlineTag}>{run.approvalStatus}</span>
                      <span className={styles.inlineTag}>{run.tokensIn + run.tokensOut} tokens</span>
                      <button className={styles.ghostButton} onClick={() => void handleInspectControlTowerDetail("agent-run", run.id)} type="button">
                        Inspect
                      </button>
                      {run.status === "blocked" && run.approvalStatus === "approved" ? (
                        <button className={styles.ghostButton} onClick={() => void handleResumeAgentRun(run.id)} type="button">
                          Resume
                        </button>
                      ) : null}
                    </div>
                  </article>
                ))}
              </div>
            </section>
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
                    {isControlTowerLoading ? (
                      Array.from({ length: 4 }, (_, index) => (
                        <tr key={`cost-skeleton-${index}`}>
                          <td colSpan={5}>
                            <div className={styles.skeletonCard}>
                              <div className={`${styles.skeleton} ${styles.skeletonWide}`} />
                              <div className={`${styles.skeleton} ${styles.skeletonNarrow}`} />
                            </div>
                          </td>
                        </tr>
                      ))
                    ) : costLedger.length > 0 ? (
                      costLedger.slice(0, controlTowerListLimits.costs).map((entry) => (
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
                          {renderEmptyState("inbox", "No cost data yet.", "Costs appear after provider-backed runs or evaluations.")}
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </section>
            {renderControlTowerDetailPanel()}
          </>
        ) : null}

        {activeTab === 1 ? (
          <>
            <section className={styles.panelWide}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Alerts</p>
                  <h2>Budget, delivery, and runtime warnings</h2>
                </div>
              </div>
              <div className={styles.listStack}>
                {filteredAlerts.length === 0 ? (
                  renderEmptyState("inbox", "No alerts triggered.", "Alerts appear when budgets, deliveries, or runtime thresholds are breached.")
                ) : (
                  filteredAlerts.slice(0, controlTowerListLimits.alerts).map((alert) => (
                    <article className={styles.auditRow} key={alert.id}>
                      <div>
                        <strong>{alert.title}</strong>
                        <p>{alert.summary}</p>
                      </div>
                      <div className={styles.inlineList}>
                        <span className={styles.inlineTag}>{alert.category}</span>
                        <span className={styles.inlineTag}>{alert.severity}</span>
                        <button className={styles.ghostButton} onClick={() => void handleInspectControlTowerDetail("alert", alert.id)} type="button">
                          Inspect
                        </button>
                        <button className={styles.ghostButton} disabled={Boolean(alert.acknowledgedAt)} onClick={() => void handleAcknowledgeAlert(alert.id)} type="button">
                          {alert.acknowledgedAt ? "Acknowledged" : "Acknowledge"}
                        </button>
                      </div>
                    </article>
                  ))
                )}
                {filteredAlerts.length > controlTowerListLimits.alerts ? (
                  <button className={styles.showMoreButton} onClick={() => setControlTowerListLimits((prev) => ({ ...prev, alerts: prev.alerts + 24 }))} type="button">
                    Show more ({filteredAlerts.length - controlTowerListLimits.alerts} remaining)
                  </button>
                ) : null}
              </div>
            </section>
            <section className={styles.panelWide}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Approvals</p>
                  <h2>Pending human checkpoints</h2>
                </div>
                {visibleControlTowerApprovals.length > 0 ? <span className={styles.badge}>{visibleControlTowerApprovals.length} pending</span> : null}
              </div>
              <div className={styles.listStack}>
                {approvalTasks.length === 0 ? (
                  renderEmptyState("workflow", "No approval tasks pending.", "Approvals appear when workflows or agents require human sign-off before continuing.")
                ) : (
                  approvalTasks.map((task) => (
                    <article className={styles.workflowCard} key={task.id}>
                      <strong>{task.nodeLabel}</strong>
                      <p>
                        {task.workflowKey} · {task.status} · {task.approverRole}
                      </p>
                      {task.instructions ? <div className={styles.sidebarPanel}>{task.instructions}</div> : null}
                      <div className={styles.inlineList}>
                        <button className={styles.ghostButton} onClick={() => void handleInspectControlTowerDetail("approval", task.id)} type="button">
                          Inspect
                        </button>
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
            {renderControlTowerDetailPanel()}
          </>
        ) : null}

        {activeTab === 2 ? (
          <>
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
                    {filteredDeliveries.length > 0 ? (
                      filteredDeliveries.slice(0, controlTowerListLimits.deliveries).map((delivery) => (
                        <tr key={delivery.id}>
                          <td>{delivery.ruleKey}</td>
                          <td>{delivery.channelKey}</td>
                          <td>{delivery.status}</td>
                          <td>{delivery.severity}</td>
                          <td>
                            <div className={styles.inlineList}>
                              <span>{formatPlatformDateTime(delivery.deliveredAt ?? delivery.createdAt)}</span>
                              <button className={styles.ghostButton} onClick={() => void handleInspectControlTowerDetail("delivery", delivery.id)} type="button">
                                Inspect
                              </button>
                              <button className={styles.ghostButton} disabled={delivery.status === "sent"} onClick={() => void handleRetryDelivery(delivery.id)} type="button">
                                Retry
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={5}>
                          {renderEmptyState("inbox", "No deliveries yet.", "Notification deliveries will appear here after alerts or workflows fan out to channels.")}
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </section>
            <section className={styles.panelWide}>
              <div className={styles.sectionHeader}>
                <div>
                  <p className={styles.cardEyebrow}>Dead letters</p>
                  <h2>Escalated runtime failures</h2>
                </div>
                {deadLetters.length > 0 ? <span className={styles.badge}>{deadLetters.length} blocked</span> : null}
              </div>
              <div className={styles.listStack}>
                {deadLetters.length === 0 ? (
                  renderEmptyState("inbox", "No dead letters recorded.", "Dead letters capture messages that could not be delivered or processed after all retries.")
                ) : (
                  deadLetters.map((entry) => (
                    <article className={styles.auditRow} key={entry.id}>
                      <div>
                        <strong>{entry.type}</strong>
                        <p>{entry.reason}</p>
                      </div>
                      <div className={styles.inlineList}>
                        <span>{formatPlatformDateTime(entry.createdAt)}</span>
                        <button className={styles.ghostButton} onClick={() => void handleInspectControlTowerDetail("dead-letter", entry.id)} type="button">
                          Inspect
                        </button>
                      </div>
                    </article>
                  ))
                )}
              </div>
            </section>
            {renderControlTowerDetailPanel()}
          </>
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
            data-tooltip={sidebarCollapsed ? "Expand sidebar" : "Collapse sidebar"}
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
          <nav className={styles.navStack}>
            {WORKSPACE_GROUPS.map((group) => {
              const isCollapsed = collapsedGroups.has(group.key) && !group.workspaces.includes(workspace);
              const groupEntries = WORKSPACES.filter((entry) => group.workspaces.includes(entry.key));
              return (
                <div className={`${styles.navGroup} ${isCollapsed ? styles.navGroupCollapsed : ""}`} key={group.key}>
                  <button
                    className={styles.navGroupHeader}
                    onClick={() => setCollapsedGroups((prev) => { const next = new Set(prev); if (next.has(group.key)) next.delete(group.key); else next.add(group.key); return next; })}
                    type="button"
                  >
                    {group.label}
                    <svg className={styles.navGroupChevron} fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" viewBox="0 0 10 10"><path d="M3 1l4 4-4 4" /></svg>
                  </button>
                  <div className={styles.navGroupBody}>
                    {groupEntries.map((entry) => (
                      <button
                        className={`${entry.key === workspace ? styles.activeNavItem : styles.navItem} ${styles.staggerChild}`}
                        key={entry.key}
                        onClick={() => selectWorkspace(entry.key)}
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
                  </div>
                </div>
              );
            })}
          </nav>
        </div>
        <div className={styles.sidebarSection}>
          <p className={styles.sidebarLabel}>Publish</p>
          <div className={styles.sidebarPanel}>
            <ul className={styles.compactBulletList}>
              <li>Draft changes stay private until publish.</li>
              <li>Publishing freezes a versioned runtime manifest.</li>
              <li>Use preview links before shipping live.</li>
            </ul>
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

      <main className={styles.studioMain} ref={mainRef}>
        <nav className={[styles.breadcrumb, styles.breadcrumbSticky].join(" ")}>
          <button
            aria-label="Studio home"
            className={styles.breadcrumbIcon}
            onClick={() => {
              deselectCurrentSelection();
              setActiveTab(0);
              window.scrollTo({ top: 0, behavior: "smooth" });
            }}
            type="button"
          >
            <svg fill="none" viewBox="0 0 16 16">
              <path d="M2 7.2 8 2l6 5.2V14H9.6V9.8H6.4V14H2V7.2Z" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.4" />
            </svg>
          </button>
          <button type="button" className={styles.breadcrumbLink} onClick={() => { deselectCurrentSelection(); setActiveTab(0); window.scrollTo({ top: 0, behavior: "smooth" }); }}>Studio</button>
          <span className={styles.breadcrumbSep}>/</span>
          <button type="button" className={styles.breadcrumbLink} onClick={() => { deselectCurrentSelection(); setActiveTab(0); }}>{activeWorkspace.label}</button>
          <span className={styles.breadcrumbSep}>/</span>
          <span className={styles.breadcrumbActive}>{WORKSPACE_TABS[workspace][activeTab]}</span>
          {(() => {
            const entityLabel =
              workspace === "data-model" ? selectedObject?.label :
              workspace === "pages" ? selectedPage?.title :
              workspace === "forms" ? selectedForm?.title :
              workspace === "workflows" ? selectedWorkflow?.name :
              workspace === "agents" ? selectedAgent?.name :
              null;
            return entityLabel ? (
              <>
                <span className={styles.breadcrumbSep}>/</span>
                <button
                  className={styles.breadcrumbEntity}
                  onClick={() => {
                    if (workspace === "pages" && selectedPage?.id) {
                      document.querySelector<HTMLElement>(`[data-page-tree-id="${selectedPage.id}"]`)?.scrollIntoView({ block: "center", behavior: "smooth" });
                    }
                    if (workspace === "data-model" && selectedObject?.id) {
                      document.querySelector<HTMLElement>(`[data-object-list-id="${selectedObject.id}"]`)?.scrollIntoView({ block: "center", behavior: "smooth" });
                    }
                    if (workspace === "forms" && selectedForm?.id) {
                      document.querySelector<HTMLElement>(`[data-form-list-id="${selectedForm.id}"]`)?.scrollIntoView({ block: "center", behavior: "smooth" });
                    }
                    if (workspace === "workflows" && selectedWorkflow?.id) {
                      document.querySelector<HTMLElement>(`[data-workflow-list-id="${selectedWorkflow.id}"]`)?.scrollIntoView({ block: "center", behavior: "smooth" });
                    }
                    if (workspace === "agents" && selectedAgent?.id) {
                      document.querySelector<HTMLElement>(`[data-agent-list-id="${selectedAgent.id}"]`)?.scrollIntoView({ block: "center", behavior: "smooth" });
                    }
                  }}
                  type="button"
                >
                  {entityLabel}
                </button>
              </>
            ) : null;
          })()}
          <button
            className={styles.commandPaletteTrigger}
            onClick={() => {
              if (typeof document !== "undefined" && document.activeElement instanceof HTMLElement) {
                previouslyFocusedElementRef.current = document.activeElement;
              }
              setCommandPaletteOpen(true);
              setCommandPaletteQuery("");
              setCommandPaletteIndex(0);
            }}
            type="button"
          >
            <span className={styles.commandPaletteTriggerCopy}>
              <svg fill="none" viewBox="0 0 16 16">
                <circle cx="7" cy="7" r="4.75" stroke="currentColor" strokeWidth="1.5" />
                <path d="M10.5 10.5 14 14" stroke="currentColor" strokeLinecap="round" strokeWidth="1.5" />
              </svg>
              <span>Search workspaces, pages, workflows, and agents</span>
            </span>
            <span className={styles.kbdBadge}>⌘K</span>
          </button>
        </nav>

        <header className={[styles.studioHeader, scrolledPast ? styles.headerCompact : ""].join(" ")}>
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

        {renderTabBar()}

        <div className={styles.workspaceEnter} key={workspaceTransitionKey}>
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
        </div>

        {isPublishPreviewLoading ? (
          <section className={styles.publishPreviewStrip}>
            {renderSkeleton(3, true)}
          </section>
        ) : publishPreview ? (
          <section className={styles.publishPreviewStrip}>
            <div>
              <p className={styles.cardEyebrow}>Publish preview</p>
              <h3>Next activation will create v{publishPreview.nextVersionNumber}</h3>
            </div>
            <div className={styles.previewMetrics}>
              <div className={styles.previewMetric}>
                <span>Branding</span>
                <strong>
                  +{publishPreview.summary.branding.added.length} / ~{publishPreview.summary.branding.updated.length} / -
                  {publishPreview.summary.branding.removed.length}
                </strong>
              </div>
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
      {renderConfirmDialog()}
      {renderShortcutsDialog()}
      {renderCommandPalette()}
      {renderToastStack()}
    </div>
  );
}
