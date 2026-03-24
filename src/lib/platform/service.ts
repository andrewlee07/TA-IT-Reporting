import { nanoid } from "nanoid";
import { Prisma } from "@/generated/prisma/client";

import { prepareAgentInvocation } from "@/lib/platform/agent-gateway";
import {
  createSignedPlatformSessionValue,
  resolvePlatformActorIdentity,
  tryResolvePlatformActorIdentity,
  tryResolvePlatformViewAsState,
  type PlatformActorIdentity,
} from "@/lib/platform/auth";
import {
  countLayoutComponents,
  createDefaultPlacement,
  createDesignerCatalog,
  normalizeLayoutComponentDefinition,
  normalizeLayoutSectionDefinition,
} from "@/lib/platform/designer";
import { createStarterManifest } from "@/lib/platform/defaults";
import {
  acceptLocalInvite,
  createLocalInvite,
  createLocalMembership,
  createLocalWorkflowRun,
  createLocalAuditEvent,
  createLocalEnvironment,
  createLocalTenant,
  createLocalUser,
  createLocalVersion,
  deleteLocalPlatformRecord,
  getLocalActiveVersion,
  getLocalDraft,
  getLocalEnvironmentBySlug,
  getLocalInviteByToken,
  getLocalMembership,
  getLocalTenantBySlug,
  getLocalUserByEmail,
  listLocalAuditEvents,
  listLocalEnvironments,
  listLocalInvites,
  listLocalPlatformRecords,
  listLocalTenantMembershipSummariesForUser,
  listLocalTenants,
  listLocalWorkflowRuns,
  listLocalVersions,
  setLocalActiveVersion,
  updateLocalWorkflowRun,
  upsertLocalDraft,
  upsertLocalPlatformRecord,
} from "@/lib/platform/local-store";
import { ensureManifestConsistency, findObjectDefinition, touchManifest } from "@/lib/platform/manifest";
import { buildPublishedArtifacts } from "@/lib/platform/publish";
import { enqueueOutboxEvent, enqueueWorkflowRun } from "@/lib/platform/execution-bus";
import {
  PlatformError,
  PlatformForbiddenError,
  PlatformNotFoundError,
} from "@/lib/platform/errors";
import { assertRole } from "@/lib/platform/rbac";
import { validateRecordInput } from "@/lib/platform/records";
import { deliverNotification, executeAgentWithProvider, summarizeAgentRunCost } from "@/lib/platform/runtime-adapters";
import { createDefaultBranding } from "@/lib/platform/theme";
import { getObjectStorage } from "@/lib/storage";
import type {
  AgentDefinition,
  AgentRunRecord,
  AgentRunDetail,
  AgentTrace,
  AppShellDefinition,
  CostLedgerRecord,
  DeadLetterRecord,
  EventEnvelope,
  FieldDefinition,
  FormDefinition,
  FormFieldDefinition,
  FormStepDefinition,
  LayoutDefinition,
  MenuItemDefinition,
  ModelProviderDefinition,
  NotificationCenterDefinition,
  NotificationAttemptRecord,
  NotificationChannelHealth,
  NotificationDeliveryRecord,
  NotificationRuleMatchRecord,
  ObjectDefinition,
  PageDefinition,
  PlatformAlertRecord,
  PlatformApprovalTaskRecord,
  PlatformAgentEvalRecord,
  PlatformActor,
  PlatformAgentPreview,
  PlatformAuditEventRecord,
  PlatformBootstrap,
  PlatformEnvironmentSummary,
  PlatformFormSubmissionRecord,
  PlatformInviteRecord,
  PlatformManifest,
  PlatformPublishPageImpact,
  PlatformPublishPreview,
  PlatformPublishRouteImpact,
  PlatformPublishedVersionRecord,
  PlatformRecord,
  PlatformRole,
  PlatformSectionTemplate,
  PlatformSessionSummary,
  PlatformTenantSummary,
  SecurityPolicyDefinition,
  SubflowDefinition,
  TenantBrandingDefinition,
  WorkflowDefinition,
  WorkflowRunDetail,
  WorkflowTemplateDefinition,
  WorkflowTestCaseDefinition,
  PlatformWorkflowRunRecord,
} from "@/lib/platform/types";
import { getEnv } from "@/lib/env";
import { getPrisma } from "@/lib/prisma";

interface PlatformContext {
  tenantId: string;
  environmentId: string;
  tenant: PlatformTenantSummary;
  environment: PlatformEnvironmentSummary;
  actor: PlatformActor;
  draftManifest: PlatformManifest;
  activeVersion: PlatformPublishedVersionRecord | null;
  versions: PlatformPublishedVersionRecord[];
  auditEvents: PlatformAuditEventRecord[];
}

function systemMaintenanceActor(): PlatformActor {
  return {
    email: "platform-worker@local.test",
    name: "Platform Worker",
    role: "SUPER_ADMIN",
  };
}

type EditableFormFieldInput = Partial<Omit<FormFieldDefinition, "id" | "key">> &
  Pick<FormFieldDefinition, "label" | "type"> & {
  id?: string;
  key?: string;
};

type EditableFormStepInput = Partial<Omit<FormStepDefinition, "id" | "key">> &
  Pick<FormStepDefinition, "title"> & {
  id?: string;
  key?: string;
};

type EditableFormInput = Omit<FormDefinition, "id" | "key" | "fields" | "steps"> & {
  id?: string;
  key?: string;
  fields?: EditableFormFieldInput[];
  steps?: EditableFormStepInput[];
};

function toJsonValue(value: unknown): Prisma.InputJsonValue {
  return JSON.parse(JSON.stringify(value)) as Prisma.InputJsonValue;
}

function isPersistenceFallbackError(error: unknown): boolean {
  if (!(error instanceof Error)) {
    return false;
  }

  return [
    "DATABASE_URL is required",
    "User was denied access on the database",
    "Can't reach database server",
    "Connection refused",
    "connection pool",
    "does not exist",
    "The table",
    "The column",
    "Unknown field",
  ].some((message) => error.message.includes(message));
}

async function withLocalFallback<T>(action: () => Promise<T>, fallback: () => Promise<T>): Promise<T> {
  try {
    return await action();
  } catch (error) {
    if (!getEnv().PLATFORM_LOCAL_DEV_MODE || !isPersistenceFallbackError(error)) {
      throw error;
    }

    return fallback();
  }
}

function isLocalDevIdentity(identity: PlatformActorIdentity): boolean {
  const env = getEnv();
  return identity.source === "local_dev" && identity.email === env.PLATFORM_DEV_ACTOR_EMAIL;
}

function toTenantSummary(tenant: {
  id: string;
  slug: string;
  name: string;
  description: string | null;
  environments?: Array<{ slug: string; isDefault: boolean }>;
}): PlatformTenantSummary {
  return {
    id: tenant.id,
    slug: tenant.slug,
    name: tenant.name,
    description: tenant.description ?? undefined,
    defaultEnvironmentSlug: tenant.environments?.find((environment) => environment.isDefault)?.slug ?? getEnv().PLATFORM_DEFAULT_ENVIRONMENT_SLUG,
  };
}

function toEnvironmentSummary(environment: {
  id: string;
  slug: string;
  name: string;
  isDefault: boolean;
}): PlatformEnvironmentSummary {
  return {
    id: environment.id,
    slug: environment.slug,
    name: environment.name,
    isDefault: environment.isDefault,
  };
}

function toVersionRecord(version: {
  id: string;
  versionNumber: number;
  status: string;
  notes: string | null;
  manifestPath: string | null;
  gitCommitSha: string | null;
  activatedAt: Date | null;
  createdAt: Date;
}): PlatformPublishedVersionRecord {
  return {
    id: version.id,
    versionNumber: version.versionNumber,
    status: version.status as PlatformPublishedVersionRecord["status"],
    notes: version.notes,
    manifestPath: version.manifestPath,
    gitCommitSha: version.gitCommitSha,
    activatedAt: version.activatedAt?.toISOString() ?? null,
    createdAt: version.createdAt.toISOString(),
  };
}

function toAuditRecord(record: {
  id: string;
  action: string;
  resourceType: string;
  resourceId: string;
  summary: string;
  actorEmail: string | null;
  actorRole: string | null;
  createdAt: Date;
  payload: unknown;
}): PlatformAuditEventRecord {
  return {
    id: record.id,
    action: record.action,
    resourceType: record.resourceType,
    resourceId: record.resourceId,
    summary: record.summary,
    actorEmail: record.actorEmail,
    actorRole: record.actorRole as PlatformRole | null,
    createdAt: record.createdAt.toISOString(),
    payload: (record.payload as Record<string, unknown> | null) ?? null,
  };
}

function toPlatformRecord(record: {
  id: string;
  objectKey: string;
  data: unknown;
  createdAt: Date | string;
  updatedAt: Date | string;
}): PlatformRecord {
  return {
    id: record.id,
    objectKey: record.objectKey,
    data: record.data as Record<string, unknown>,
    createdAt: typeof record.createdAt === "string" ? record.createdAt : record.createdAt.toISOString(),
    updatedAt: typeof record.updatedAt === "string" ? record.updatedAt : record.updatedAt.toISOString(),
  };
}

function toWorkflowRunRecord(record: {
  id: string;
  workflowId: string;
  workflowKey: string;
  status: string;
  input?: unknown;
  output?: unknown;
  logs?: unknown;
  startedAt?: Date | string | null;
  finishedAt?: Date | string | null;
  createdAt: Date | string;
  updatedAt: Date | string;
}): PlatformWorkflowRunRecord {
  const input = (record.input as Record<string, unknown> | null) ?? null;
  const output = (record.output as Record<string, unknown> | null) ?? null;
  const runtimeState =
    output && typeof output.runtimeState === "object" && output.runtimeState
      ? (output.runtimeState as Record<string, unknown>)
      : null;

  return {
    id: record.id,
    workflowId: record.workflowId,
    workflowKey: record.workflowKey,
    status: record.status as PlatformWorkflowRunRecord["status"],
    input,
    output,
    logs: (record.logs as Array<Record<string, unknown>> | null) ?? [],
    parentWorkflowRunId: typeof input?.parentWorkflowRunId === "string" ? input.parentWorkflowRunId : null,
    replayedFromRunId:
      typeof input?.replayedFromRunId === "string"
        ? input.replayedFromRunId
        : typeof input?.replayOfRunId === "string"
          ? input.replayOfRunId
          : null,
    pauseReason:
      runtimeState && typeof runtimeState.pauseReason === "string"
        ? (runtimeState.pauseReason as PlatformWorkflowRunRecord["pauseReason"])
        : null,
    subflowRunId: runtimeState && typeof runtimeState.subflowRunId === "string" ? runtimeState.subflowRunId : null,
    nextRetryAt: runtimeState && typeof runtimeState.nextRetryAt === "string" ? runtimeState.nextRetryAt : null,
    startedAt:
      record.startedAt == null ? null : typeof record.startedAt === "string" ? record.startedAt : record.startedAt.toISOString(),
    finishedAt:
      record.finishedAt == null ? null : typeof record.finishedAt === "string" ? record.finishedAt : record.finishedAt.toISOString(),
    createdAt: typeof record.createdAt === "string" ? record.createdAt : record.createdAt.toISOString(),
    updatedAt: typeof record.updatedAt === "string" ? record.updatedAt : record.updatedAt.toISOString(),
  };
}

function buildInviteUrl(tenantSlug: string, token: string): string {
  return `${getEnv().APP_BASE_URL}/platform/login?tenant=${encodeURIComponent(tenantSlug)}&invite=${encodeURIComponent(token)}`;
}

function createInviteToken(): string {
  return `invite_${nanoid(24)}`;
}

function findFormDefinition(manifest: PlatformManifest, formIdOrKey: string): FormDefinition | undefined {
  return manifest.forms.find((form) => form.id === formIdOrKey || form.key === formIdOrKey || form.route === formIdOrKey);
}

function getFormSubmissionObjectKey(formKey: string): string {
  return `form_submission:${formKey}`;
}

const SYSTEM_OBJECT_KEYS = {
  outbox: "__system_outbox_event",
  delivery: "__system_notification_delivery",
  deliveryAttempt: "__system_notification_delivery_attempt",
  ruleMatch: "__system_notification_rule_match",
  channelHealth: "__system_notification_channel_health",
  alert: "__system_alert",
  agentRun: "__system_agent_run",
  costLedger: "__system_cost_ledger",
  deadLetter: "__system_dead_letter",
  approvalTask: "__system_approval_task",
} as const;

function getSystemObjectKey(value: keyof typeof SYSTEM_OBJECT_KEYS): string {
  return SYSTEM_OBJECT_KEYS[value];
}

function toFormSubmissionRecord(record: PlatformRecord): PlatformFormSubmissionRecord {
  return {
    id: record.id,
    formKey: String(record.data.formKey ?? ""),
    objectKey: typeof record.data.objectKey === "string" ? record.data.objectKey : undefined,
    status: (record.data.status as PlatformFormSubmissionRecord["status"]) ?? "submitted",
    data: (record.data.submission as Record<string, unknown>) ?? {},
    createdAt: record.createdAt,
    submittedAt: typeof record.data.submittedAt === "string" ? record.data.submittedAt : null,
    createdByEmail: typeof record.data.createdByEmail === "string" ? record.data.createdByEmail : null,
  };
}

function toEventEnvelope(record: PlatformRecord): EventEnvelope {
  return {
    id: record.id,
    type: String(record.data.type ?? "system.event"),
    tenantSlug: String(record.data.tenantSlug ?? ""),
    environmentSlug: String(record.data.environmentSlug ?? ""),
    emittedAt: typeof record.data.emittedAt === "string" ? record.data.emittedAt : record.createdAt,
    source: (record.data.source as EventEnvelope["source"]) ?? "system",
    resourceType: String(record.data.resourceType ?? "system"),
    resourceId: String(record.data.resourceId ?? record.id),
    payload: (record.data.payload as Record<string, unknown>) ?? {},
  };
}

function toNotificationDeliveryRecord(record: PlatformRecord): NotificationDeliveryRecord {
  return {
    id: record.id,
    eventId: String(record.data.eventId ?? ""),
    ruleKey: String(record.data.ruleKey ?? ""),
    channelKey: String(record.data.channelKey ?? ""),
    templateKey: String(record.data.templateKey ?? ""),
    status: (record.data.status as NotificationDeliveryRecord["status"]) ?? "pending",
    severity: (record.data.severity as NotificationDeliveryRecord["severity"]) ?? "info",
    subject: typeof record.data.subject === "string" ? record.data.subject : undefined,
    body: String(record.data.body ?? ""),
    destination: typeof record.data.destination === "string" ? record.data.destination : undefined,
    provider: typeof record.data.provider === "string" ? record.data.provider : undefined,
    providerResponseSummary: typeof record.data.providerResponseSummary === "string" ? record.data.providerResponseSummary : null,
    attemptCount: typeof record.data.attemptCount === "number" ? record.data.attemptCount : 0,
    maxAttempts: typeof record.data.maxAttempts === "number" ? record.data.maxAttempts : undefined,
    lastAttemptAt: typeof record.data.lastAttemptAt === "string" ? record.data.lastAttemptAt : null,
    nextRetryAt: typeof record.data.nextRetryAt === "string" ? record.data.nextRetryAt : null,
    disabledAt: typeof record.data.disabledAt === "string" ? record.data.disabledAt : null,
    resolvedPayload: (record.data.resolvedPayload as Record<string, unknown> | null) ?? null,
    createdAt: record.createdAt,
    deliveredAt: typeof record.data.deliveredAt === "string" ? record.data.deliveredAt : null,
    exhaustedAt: typeof record.data.exhaustedAt === "string" ? record.data.exhaustedAt : null,
    errorMessage: typeof record.data.errorMessage === "string" ? record.data.errorMessage : null,
  };
}

function toNotificationAttemptRecord(record: PlatformRecord): NotificationAttemptRecord {
  return {
    id: record.id,
    deliveryId: String(record.data.deliveryId ?? ""),
    eventId: String(record.data.eventId ?? ""),
    channelKey: String(record.data.channelKey ?? ""),
    status: (record.data.status as NotificationAttemptRecord["status"]) ?? "pending",
    provider: String(record.data.provider ?? "unknown"),
    destination: typeof record.data.destination === "string" ? record.data.destination : undefined,
    attemptNumber: typeof record.data.attemptNumber === "number" ? record.data.attemptNumber : 1,
    lifecycleStage:
      typeof record.data.lifecycleStage === "string"
        ? (record.data.lifecycleStage as NotificationAttemptRecord["lifecycleStage"])
        : undefined,
    responseSummary: typeof record.data.responseSummary === "string" ? record.data.responseSummary : null,
    errorMessage: typeof record.data.errorMessage === "string" ? record.data.errorMessage : null,
    createdAt: record.createdAt,
  };
}

function toNotificationChannelHealth(record: PlatformRecord): NotificationChannelHealth {
  return {
    id: record.id,
    channelKey: String(record.data.channelKey ?? ""),
    status: (record.data.status as NotificationChannelHealth["status"]) ?? "healthy",
    successCount: typeof record.data.successCount === "number" ? record.data.successCount : 0,
    failureCount: typeof record.data.failureCount === "number" ? record.data.failureCount : 0,
    consecutiveFailures: typeof record.data.consecutiveFailures === "number" ? record.data.consecutiveFailures : 0,
    successRate: typeof record.data.successRate === "number" ? record.data.successRate : 0,
    lastDeliveredAt: typeof record.data.lastDeliveredAt === "string" ? record.data.lastDeliveredAt : null,
    lastFailedAt: typeof record.data.lastFailedAt === "string" ? record.data.lastFailedAt : null,
    disabledAt: typeof record.data.disabledAt === "string" ? record.data.disabledAt : null,
    updatedAt: typeof record.data.updatedAt === "string" ? record.data.updatedAt : record.updatedAt,
  };
}

function toNotificationRuleMatchRecord(record: PlatformRecord): NotificationRuleMatchRecord {
  return {
    id: record.id,
    eventId: String(record.data.eventId ?? ""),
    ruleKey: String(record.data.ruleKey ?? ""),
    templateKey: String(record.data.templateKey ?? ""),
    matchedAt: typeof record.data.matchedAt === "string" ? record.data.matchedAt : record.createdAt,
    channelKeys: Array.isArray(record.data.channelKeys) ? record.data.channelKeys.map((value) => String(value)) : [],
    payload: (record.data.payload as Record<string, unknown>) ?? {},
  };
}

function toPlatformAlertRecord(record: PlatformRecord): PlatformAlertRecord {
  return {
    id: record.id,
    category: (record.data.category as PlatformAlertRecord["category"]) ?? "runtime",
    severity: (record.data.severity as PlatformAlertRecord["severity"]) ?? "info",
    title: String(record.data.title ?? "Platform alert"),
    summary: String(record.data.summary ?? ""),
    sourceId: String(record.data.sourceId ?? record.id),
    createdAt: record.createdAt,
    acknowledgedAt: typeof record.data.acknowledgedAt === "string" ? record.data.acknowledgedAt : null,
  };
}

function toApprovalTaskRecord(record: PlatformRecord): PlatformApprovalTaskRecord {
  return {
    id: record.id,
    workflowRunId: String(record.data.workflowRunId ?? ""),
    workflowKey: String(record.data.workflowKey ?? ""),
    nodeId: String(record.data.nodeId ?? ""),
    nodeLabel: String(record.data.nodeLabel ?? ""),
    taskType: (record.data.taskType as PlatformApprovalTaskRecord["taskType"]) ?? "workflow_node",
    agentRunId: typeof record.data.agentRunId === "string" ? record.data.agentRunId : null,
    agentKey: typeof record.data.agentKey === "string" ? record.data.agentKey : null,
    approverRole: (record.data.approverRole as PlatformApprovalTaskRecord["approverRole"]) ?? "BUILDER_ADMIN",
    status: (record.data.status as PlatformApprovalTaskRecord["status"]) ?? "pending",
    instructions: typeof record.data.instructions === "string" ? record.data.instructions : null,
    createdAt: record.createdAt,
    resolvedAt: typeof record.data.resolvedAt === "string" ? record.data.resolvedAt : null,
  };
}

function toAgentRunRecord(record: PlatformRecord): AgentRunRecord {
  return {
    id: record.id,
    agentId: String(record.data.agentId ?? ""),
    agentKey: String(record.data.agentKey ?? ""),
    status: (record.data.status as AgentRunRecord["status"]) ?? "queued",
    runMode: (record.data.runMode as AgentRunRecord["runMode"]) ?? "simulation",
    approvalStatus: (record.data.approvalStatus as AgentRunRecord["approvalStatus"]) ?? "not_required",
    input: (record.data.input as Record<string, unknown>) ?? {},
    output: (record.data.output as Record<string, unknown>) ?? null,
    logs: (record.data.logs as Array<Record<string, unknown>>) ?? [],
    modelProviderKey: String(record.data.modelProviderKey ?? ""),
    costUsd: typeof record.data.costUsd === "number" ? record.data.costUsd : 0,
    tokensIn: typeof record.data.tokensIn === "number" ? record.data.tokensIn : 0,
    tokensOut: typeof record.data.tokensOut === "number" ? record.data.tokensOut : 0,
    trace: (record.data.trace as AgentTrace | null) ?? null,
    approvalTaskId: typeof record.data.approvalTaskId === "string" ? record.data.approvalTaskId : null,
    handoffWorkflowRunId: typeof record.data.handoffWorkflowRunId === "string" ? record.data.handoffWorkflowRunId : null,
    parentWorkflowRunId: typeof record.data.parentWorkflowRunId === "string" ? record.data.parentWorkflowRunId : null,
    outputValidationPassed: typeof record.data.outputValidationPassed === "boolean" ? record.data.outputValidationPassed : null,
    schemaValidation:
      typeof record.data.schemaValidation === "object" && record.data.schemaValidation
        ? (record.data.schemaValidation as AgentRunRecord["schemaValidation"])
        : null,
    createdAt: record.createdAt,
    completedAt: typeof record.data.completedAt === "string" ? record.data.completedAt : null,
  };
}

function toCostLedgerRecord(record: PlatformRecord): CostLedgerRecord {
  return {
    id: record.id,
    category: (record.data.category as CostLedgerRecord["category"]) ?? "agent_run",
    referenceId: String(record.data.referenceId ?? ""),
    providerKey: typeof record.data.providerKey === "string" ? record.data.providerKey : undefined,
    amountUsd: typeof record.data.amountUsd === "number" ? record.data.amountUsd : 0,
    tokensIn: typeof record.data.tokensIn === "number" ? record.data.tokensIn : undefined,
    tokensOut: typeof record.data.tokensOut === "number" ? record.data.tokensOut : undefined,
    createdAt: record.createdAt,
    summary: String(record.data.summary ?? ""),
  };
}

function toDeadLetterRecord(record: PlatformRecord): DeadLetterRecord {
  return {
    id: record.id,
    eventId: String(record.data.eventId ?? ""),
    type: String(record.data.type ?? ""),
    reason: String(record.data.reason ?? ""),
    payload: (record.data.payload as Record<string, unknown>) ?? {},
    createdAt: record.createdAt,
  };
}

function toInviteRecord(record: {
  id: string;
  tenantId: string;
  email: string;
  role: string;
  token: string;
  status: string;
  inviteUrl: string;
  createdByEmail?: string | null | undefined;
  createdAt: Date | string;
  expiresAt: Date | string;
  acceptedAt?: Date | string | null;
} & { tenant?: { slug: string } }, tenantSlug?: string): PlatformInviteRecord {
  return {
    id: record.id,
    tenantId: record.tenantId,
    tenantSlug: tenantSlug ?? record.tenant?.slug ?? "",
    email: record.email,
    role: record.role as PlatformRole,
    status: record.status as PlatformInviteRecord["status"],
    token: record.token,
    inviteUrl: record.inviteUrl,
    createdByEmail: record.createdByEmail,
    createdAt: typeof record.createdAt === "string" ? record.createdAt : record.createdAt.toISOString(),
    expiresAt: typeof record.expiresAt === "string" ? record.expiresAt : record.expiresAt.toISOString(),
    acceptedAt:
      record.acceptedAt == null ? null : typeof record.acceptedAt === "string" ? record.acceptedAt : record.acceptedAt.toISOString(),
  };
}

async function listDbMembershipSummariesForUser(userId: string) {
  const prisma = getPrisma();
  const memberships = await prisma.platformMembership.findMany({
    where: { userId },
    include: {
      tenant: {
        include: {
          environments: {
            where: { isDefault: true },
            take: 1,
          },
        },
      },
    },
    orderBy: [
      {
        tenant: {
          name: "asc",
        },
      },
    ],
  });

  return memberships.map((membership) => ({
    tenantId: membership.tenantId,
    tenantSlug: membership.tenant.slug,
    tenantName: membership.tenant.name,
    defaultEnvironmentSlug: membership.tenant.environments[0]?.slug ?? getEnv().PLATFORM_DEFAULT_ENVIRONMENT_SLUG,
    role: membership.role as PlatformRole,
  }));
}

async function ensureDbContext(input: {
  tenantSlug: string;
  identity: PlatformActorIdentity;
  environmentSlug?: string;
}): Promise<PlatformContext> {
  const prisma = getPrisma();
  const tenantName = input.tenantSlug
    .split("-")
    .map((segment) => segment.charAt(0).toUpperCase() + segment.slice(1))
    .join(" ");
  let tenant = await prisma.platformTenant.findUnique({
    where: { slug: input.tenantSlug },
  });

  if (!tenant) {
    if (!isLocalDevIdentity(input.identity)) {
      throw new PlatformNotFoundError("Platform tenant not found.");
    }

    tenant = await prisma.platformTenant.create({
      data: {
        slug: input.tenantSlug,
        name: tenantName,
        description: `Adaptive platform tenant for ${tenantName}.`,
      },
    });
  }

  const environmentSlug = input.environmentSlug ?? getEnv().PLATFORM_DEFAULT_ENVIRONMENT_SLUG;
  let environment = await prisma.platformEnvironment.findUnique({
    where: {
      tenantId_slug: {
        tenantId: tenant.id,
        slug: environmentSlug,
      },
    },
  });

  if (!environment) {
    if (!isLocalDevIdentity(input.identity)) {
      throw new PlatformNotFoundError("Platform environment not found.");
    }

    environment = await prisma.platformEnvironment.create({
      data: {
        tenantId: tenant.id,
        slug: environmentSlug,
        name: environmentSlug.charAt(0).toUpperCase() + environmentSlug.slice(1),
        isDefault: true,
      },
    });
  }

  const user = await prisma.platformUser.upsert({
    where: { email: input.identity.email },
    update: {
      displayName: input.identity.name,
    },
    create: {
      email: input.identity.email,
      displayName: input.identity.name,
    },
  });

  let membership = await prisma.platformMembership.findUnique({
    where: {
      userId_tenantId: {
        userId: user.id,
        tenantId: tenant.id,
      },
    },
  });

  if (!membership) {
    if (!isLocalDevIdentity(input.identity)) {
      throw new PlatformForbiddenError("User is not a member of this tenant.");
    }

    membership = await prisma.platformMembership.create({
      data: {
        userId: user.id,
        tenantId: tenant.id,
        role: getEnv().PLATFORM_DEV_ACTOR_ROLE,
      },
    });
  }

  const actor: PlatformActor = {
    email: user.email,
    name: user.displayName,
    role: membership.role as PlatformRole,
  };

  const existingDraft = await prisma.platformDraft.findUnique({
    where: {
      tenantId_environmentId: {
        tenantId: tenant.id,
        environmentId: environment.id,
      },
    },
  });

  const seededManifest = createStarterManifest(tenant.slug, tenant.name);
  seededManifest.environment = {
    slug: environment.slug,
    name: environment.name,
  };

  const draftRecord = existingDraft
    ? await prisma.platformDraft.update({
        where: { id: existingDraft.id },
        data: {
          manifest: toJsonValue({
            ...(existingDraft.manifest as unknown as PlatformManifest),
            tenant: {
              slug: tenant.slug,
              name: tenant.name,
              description: tenant.description ?? undefined,
            },
            environment: {
              slug: environment.slug,
              name: environment.name,
            },
          }),
        },
      })
    : await prisma.platformDraft.create({
        data: {
          tenantId: tenant.id,
          environmentId: environment.id,
          manifest: toJsonValue(seededManifest),
        },
      });
  const normalizedDraftManifest = ensureManifestConsistency(draftRecord.manifest as unknown as PlatformManifest);

  const versions = await prisma.platformPublishedVersion.findMany({
    where: {
      tenantId: tenant.id,
      environmentId: environment.id,
    },
    orderBy: [{ createdAt: "desc" }],
  });

  const auditEvents = await prisma.platformAuditEvent.findMany({
    where: {
      tenantId: tenant.id,
      environmentId: environment.id,
    },
    orderBy: [{ createdAt: "desc" }],
    take: 40,
  });

  return {
    tenantId: tenant.id,
    environmentId: environment.id,
    tenant: {
      ...toTenantSummary({
        ...tenant,
        environments: [{ slug: environment.slug, isDefault: environment.isDefault }],
      }),
      defaultEnvironmentSlug: environment.slug,
    },
    environment: toEnvironmentSummary(environment),
    actor,
    draftManifest: normalizedDraftManifest,
    activeVersion: versions.find((version) => version.status === "ACTIVE") ? toVersionRecord(versions.find((version) => version.status === "ACTIVE")!) : null,
    versions: versions.map(toVersionRecord),
    auditEvents: auditEvents.map(toAuditRecord),
  };
}

async function ensureLocalContext(input: {
  tenantSlug: string;
  identity: PlatformActorIdentity;
  environmentSlug?: string;
}): Promise<PlatformContext> {
  const tenantName = input.tenantSlug
    .split("-")
    .map((segment) => segment.charAt(0).toUpperCase() + segment.slice(1))
    .join(" ");
  const environmentSlug = input.environmentSlug ?? getEnv().PLATFORM_DEFAULT_ENVIRONMENT_SLUG;

  let tenant = await getLocalTenantBySlug(input.tenantSlug);
  if (!tenant) {
    tenant = await createLocalTenant({
      slug: input.tenantSlug,
      name: tenantName,
      description: `Adaptive platform tenant for ${tenantName}.`,
      defaultEnvironmentSlug: environmentSlug,
    });
  }

  let environment = await getLocalEnvironmentBySlug({
    tenantId: tenant.id,
    slug: environmentSlug,
  });

  if (!environment) {
    environment = await createLocalEnvironment({
      tenantId: tenant.id,
      slug: environmentSlug,
      name: environmentSlug.charAt(0).toUpperCase() + environmentSlug.slice(1),
      isDefault: true,
    });
  }

  const user = (await getLocalUserByEmail(input.identity.email)) ??
    (await createLocalUser({
      email: input.identity.email,
      displayName: input.identity.name,
    }));
  let membership = await getLocalMembership({
    tenantId: tenant.id,
    userId: user.id,
  });

  if (!membership) {
    if (!isLocalDevIdentity(input.identity)) {
      throw new PlatformForbiddenError("User is not a member of this tenant.");
    }

    membership = await createLocalMembership({
      tenantId: tenant.id,
      userId: user.id,
      role: getEnv().PLATFORM_DEV_ACTOR_ROLE,
    });
  }

  const actor: PlatformActor = {
    email: user.email,
    name: user.displayName,
    role: membership.role,
  };

  let draft = await getLocalDraft({
    tenantId: tenant.id,
    environmentId: environment.id,
  });

  if (!draft) {
    const manifest = createStarterManifest(tenant.slug, tenant.name);
    manifest.environment = {
      slug: environment.slug,
      name: environment.name,
    };
    draft = await upsertLocalDraft({
      tenantId: tenant.id,
      environmentId: environment.id,
      manifest,
    });
  }

  const versions = await listLocalVersions({
    tenantId: tenant.id,
    environmentId: environment.id,
  });
  const activeVersion = await getLocalActiveVersion({
    tenantId: tenant.id,
    environmentId: environment.id,
  });
  const auditEvents = await listLocalAuditEvents({
    tenantId: tenant.id,
    environmentId: environment.id,
  });

  return {
    tenantId: tenant.id,
    environmentId: environment.id,
    tenant,
    environment,
    actor,
    draftManifest: ensureManifestConsistency(draft.manifest),
    activeVersion,
    versions,
    auditEvents,
  };
}

async function getPlatformContext(input: {
  tenantSlug: string;
  identity: PlatformActorIdentity;
  environmentSlug?: string;
}): Promise<PlatformContext> {
  return withLocalFallback(() => ensureDbContext(input), () => ensureLocalContext(input));
}

async function getPlatformContextFromRequest(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformContext> {
  return getPlatformContext({
    tenantSlug: input.tenantSlug,
    identity: resolvePlatformActorIdentity(input.request),
    environmentSlug: input.environmentSlug,
  });
}

async function getPlatformSessionSummary(input?: {
  request?: Request | Headers;
  fallbackActor?: PlatformActor | null;
}): Promise<PlatformSessionSummary> {
  const identity = tryResolvePlatformActorIdentity(input?.request);
  if (!identity) {
    return {
      actor: null,
      memberships: [],
      source: "none",
    };
  }

  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const user = await prisma.platformUser.findUnique({
        where: { email: identity.email },
      });

      const memberships = user ? await listDbMembershipSummariesForUser(user.id) : [];
      const actor =
        input?.fallbackActor ??
        (user
          ? {
              email: user.email,
              name: user.displayName,
              role: memberships[0]?.role ?? (getEnv().PLATFORM_DEV_ACTOR_ROLE as PlatformRole),
            }
          : null);

      return {
        actor,
        memberships,
        source: identity.source,
      };
    },
    async () => {
      const user = await getLocalUserByEmail(identity.email);
      const memberships = user ? await listLocalTenantMembershipSummariesForUser(user.id) : [];
      const actor =
        input?.fallbackActor ??
        (user
          ? {
              email: user.email,
              name: user.displayName,
              role: memberships[0]?.role ?? (getEnv().PLATFORM_DEV_ACTOR_ROLE as PlatformRole),
            }
          : null);

      return {
        actor,
        memberships,
        source: identity.source,
      };
    },
  );
}

async function listNotificationMaintenanceScopes(): Promise<Array<{
  context: PlatformContext;
  manifest: PlatformManifest;
}>> {
  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const environments = await prisma.platformEnvironment.findMany({
        include: {
          tenant: true,
          versions: {
            where: {
              status: "ACTIVE",
            },
            orderBy: [{ createdAt: "desc" }],
            take: 1,
          },
        },
      });

      return environments
        .filter((environment) => environment.versions[0]?.manifest)
        .map((environment) => {
          const activeVersion = environment.versions[0]!;
          const manifest = ensureManifestConsistency(activeVersion.manifest as unknown as PlatformManifest);
          return {
            context: {
              tenantId: environment.tenantId,
              environmentId: environment.id,
              tenant: {
                ...toTenantSummary({
                  id: environment.tenant.id,
                  slug: environment.tenant.slug,
                  name: environment.tenant.name,
                  description: environment.tenant.description,
                  environments: [{ slug: environment.slug, isDefault: environment.isDefault }],
                }),
                defaultEnvironmentSlug: environment.isDefault
                  ? environment.slug
                  : getEnv().PLATFORM_DEFAULT_ENVIRONMENT_SLUG,
              },
              environment: toEnvironmentSummary({
                id: environment.id,
                slug: environment.slug,
                name: environment.name,
                isDefault: environment.isDefault,
              }),
              actor: systemMaintenanceActor(),
              draftManifest: manifest,
              activeVersion: toVersionRecord(activeVersion),
              versions: [toVersionRecord(activeVersion)],
              auditEvents: [],
            },
            manifest,
          };
        });
    },
    async () => {
      const tenants = await listLocalTenants();
      const scopes: Array<{
        context: PlatformContext;
        manifest: PlatformManifest;
      }> = [];

      for (const tenant of tenants) {
        const environments = await listLocalEnvironments(tenant.id);
        for (const environment of environments) {
          const activeVersion = await getLocalActiveVersion({
            tenantId: tenant.id,
            environmentId: environment.id,
          });
          if (!activeVersion) {
            continue;
          }

          const manifest = ensureManifestConsistency(activeVersion.manifest);
          scopes.push({
            context: {
              tenantId: tenant.id,
              environmentId: environment.id,
              tenant,
              environment,
              actor: systemMaintenanceActor(),
              draftManifest: manifest,
              activeVersion,
              versions: [activeVersion],
              auditEvents: [],
            },
            manifest,
          });
        }
      }

      return scopes;
    },
  );
}

async function listTenantInvitesForContext(context: PlatformContext): Promise<PlatformInviteRecord[]> {
  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const invites = await prisma.platformTenantInvite.findMany({
        where: {
          tenantId: context.tenantId,
        },
        include: {
          tenant: true,
        },
        orderBy: [{ createdAt: "desc" }],
      });
      return invites.map((invite) => toInviteRecord(invite, context.tenant.slug));
    },
    async () => {
      const invites = await listLocalInvites(context.tenantId);
      return invites.map((invite) => toInviteRecord(invite, context.tenant.slug));
    },
  );
}

async function listTenantScopedRecords(input: {
  tenantId: string;
  environmentId: string;
  objectKey: string;
}): Promise<PlatformRecord[]> {
  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const records = await prisma.platformRecord.findMany({
        where: {
          tenantId: input.tenantId,
          environmentId: input.environmentId,
          objectKey: input.objectKey,
        },
        orderBy: [{ updatedAt: "desc" }],
      });
      return records.map(toPlatformRecord);
    },
    async () => {
      const records = await listLocalPlatformRecords(input);
      return records.map(toPlatformRecord);
    },
  );
}

async function upsertTenantScopedRecord(input: {
  tenantId: string;
  environmentId: string;
  objectKey: string;
  data: Record<string, unknown>;
  actor: PlatformActor;
  recordId?: string;
}): Promise<PlatformRecord> {
  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const record = input.recordId
        ? await prisma.platformRecord.findFirst({
            where: {
              id: input.recordId,
              tenantId: input.tenantId,
              environmentId: input.environmentId,
              objectKey: input.objectKey,
            },
          })
        : null;

      if (input.recordId && !record) {
        throw new PlatformNotFoundError("Record not found in the active tenant scope.");
      }

      const savedRecord = record
        ? await prisma.platformRecord.update({
            where: { id: record.id },
            data: {
              data: toJsonValue(input.data),
              updatedByEmail: input.actor.email,
            },
          })
        : await prisma.platformRecord.create({
            data: {
              tenantId: input.tenantId,
              environmentId: input.environmentId,
              objectKey: input.objectKey,
              data: toJsonValue(input.data),
              createdByEmail: input.actor.email,
              updatedByEmail: input.actor.email,
            },
          });

      return toPlatformRecord(savedRecord);
    },
    async () => {
      const record = await upsertLocalPlatformRecord({
        tenantId: input.tenantId,
        environmentId: input.environmentId,
        objectKey: input.objectKey,
        recordId: input.recordId,
        data: input.data,
        actor: input.actor,
      });
      return toPlatformRecord(record);
    },
  );
}

async function saveDraftManifest(input: {
  tenantId: string;
  environmentId: string;
  manifest: PlatformManifest;
}): Promise<void> {
  await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const existing = await prisma.platformDraft.findUnique({
        where: {
          tenantId_environmentId: {
            tenantId: input.tenantId,
            environmentId: input.environmentId,
          },
        },
      });

      if (existing) {
        await prisma.platformDraft.update({
          where: { id: existing.id },
          data: {
            manifest: toJsonValue(input.manifest),
          },
        });
        return;
      }

      await prisma.platformDraft.create({
        data: {
          tenantId: input.tenantId,
          environmentId: input.environmentId,
          manifest: toJsonValue(input.manifest),
        },
      });
    },
    async () => {
      await upsertLocalDraft({
        tenantId: input.tenantId,
        environmentId: input.environmentId,
        manifest: input.manifest,
      });
    },
  );
}

async function appendAuditEvent(input: {
  tenantId: string;
  environmentId: string;
  actor: PlatformActor;
  action: string;
  resourceType: string;
  resourceId: string;
  summary: string;
  payload?: Record<string, unknown>;
}): Promise<void> {
  await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      await prisma.platformAuditEvent.create({
        data: {
          tenantId: input.tenantId,
          environmentId: input.environmentId,
          actorEmail: input.actor.email,
          actorRole: input.actor.role,
          action: input.action,
          resourceType: input.resourceType,
          resourceId: input.resourceId,
          summary: input.summary,
          payload: input.payload ? toJsonValue(input.payload) : toJsonValue(null),
        },
      });
    },
    async () => {
      await createLocalAuditEvent({
        tenantId: input.tenantId,
        environmentId: input.environmentId,
        actor: input.actor,
        action: input.action,
        resourceType: input.resourceType,
        resourceId: input.resourceId,
        summary: input.summary,
        payload: input.payload,
      });
    },
  );
}

async function updateDraftWithMutation<T>(input: {
  tenantSlug: string;
  request?: Request | Headers;
  environmentSlug?: string;
  minimumRole?: PlatformRole;
  action: string;
  resourceType: string;
  resourceId: string;
  summary: string;
  mutate: (manifest: PlatformManifest) => { manifest: PlatformManifest; result: T };
}): Promise<T> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, input.minimumRole ?? "BUILDER_ADMIN");
  const { manifest, result } = input.mutate(structuredClone(context.draftManifest));
  const nextManifest = ensureManifestConsistency(touchManifest(manifest));

  await saveDraftManifest({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    manifest: nextManifest,
  });

  await appendAuditEvent({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    actor: context.actor,
    action: input.action,
    resourceType: input.resourceType,
    resourceId: input.resourceId,
    summary: input.summary,
  });

  return result;
}

function sanitizeKey(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .replace(/_{2,}/g, "_");
}

function nextId(prefix: string): string {
  return `${prefix}-${nanoid(8)}`;
}

function upsertById<T extends { id: string }>(items: T[], item: T): T[] {
  const filtered = items.filter((candidate) => candidate.id !== item.id);
  return [item, ...filtered];
}

function normalizeOptionalText(value: string | null | undefined): string | undefined {
  const normalized = value?.trim();
  return normalized ? normalized : undefined;
}

function requireNonEmptyText(value: string | null | undefined, label: string): string {
  const normalized = value?.trim() ?? "";
  if (!normalized) {
    throw new PlatformError(`${label} is required.`);
  }

  return normalized;
}

function requireManifestKey(value: string | null | undefined, fallback: string, label: string): string {
  const key = sanitizeKey(value ?? fallback);
  if (!key) {
    throw new PlatformError(`${label} must contain at least one letter or number.`);
  }

  return key;
}

function normalizeRoute(value: string | null | undefined, fallback: string): string {
  const route = (value ?? fallback)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9/]+/g, "-")
    .replace(/\/{2,}/g, "/")
    .replace(/^-+|-+$/g, "")
    .replace(/^-?\//, "")
    .replace(/\/-?$/, "");

  if (!route) {
    throw new PlatformError("Page route must contain at least one letter or number.");
  }

  return route;
}

function assertUniqueValue<T extends { id: string }>(
  items: T[],
  candidateId: string,
  getValue: (item: T) => string,
  nextValue: string,
  label: string,
): void {
  if (items.some((item) => item.id !== candidateId && getValue(item).toLowerCase() === nextValue.toLowerCase())) {
    throw new PlatformError(`${label} must be unique.`);
  }
}

function assertExists(value: unknown, message: string): asserts value {
  if (!value) {
    throw new PlatformNotFoundError(message);
  }
}

function createPageTemplateFromDraft(input: {
  manifest: PlatformManifest;
  pageKey: string;
  templateKey?: string;
  label?: string;
  description?: string;
}): PlatformManifest["pageTemplates"][number] {
  const page = input.manifest.pages.find((candidate) => candidate.key === input.pageKey || candidate.id === input.pageKey);
  assertExists(page, "Page template source page is invalid.");
  const layout = input.manifest.layouts.find((candidate) => candidate.key === page.layoutKey || candidate.id === page.layoutKey);
  assertExists(layout, "Page template source layout is invalid.");

  return {
    key: requireManifestKey(input.templateKey, page.key, "Page template key"),
    label: requireNonEmptyText(input.label ?? page.title, "Page template label"),
    description: normalizeOptionalText(input.description) ?? page.description ?? `Reusable template based on ${page.title}.`,
    source: "tenant",
    page: {
      key: page.key,
      title: page.title,
      description: page.description,
      objectKey: page.objectKey,
      isHome: page.isHome ?? false,
      previewNote: page.previewNote,
    },
    layout: {
      key: layout.key,
      name: layout.name,
      mobileColumns: layout.mobileColumns,
      tabletColumns: layout.tabletColumns,
      desktopColumns: layout.desktopColumns,
      sections: layout.sections.map((section) => ({
        ...section,
        id: "template-section",
        components: section.components.map((component) => ({
          ...component,
          id: "template-component",
        })),
      })),
    },
  };
}

function createSectionTemplateFromDraft(input: {
  layout: LayoutDefinition;
  sectionId: string;
  templateKey?: string;
  label?: string;
  description?: string;
}): PlatformSectionTemplate {
  const section = input.layout.sections.find((candidate) => candidate.id === input.sectionId);
  assertExists(section, "Section template source section is invalid.");

  return {
    key: requireManifestKey(input.templateKey, section.title, "Section template key"),
    label: requireNonEmptyText(input.label ?? section.title, "Section template label"),
    description: normalizeOptionalText(input.description) ?? section.description ?? `Reusable section based on ${section.title}.`,
    source: "tenant",
    section: {
      title: section.title,
      description: section.description,
      kind: section.kind,
      columns: section.columns,
      templateKey: section.templateKey,
      placement: section.placement,
      visibilityRule: section.visibilityRule,
      components: section.components.map((component) => ({
        ...component,
        id: "template-component",
      })),
    },
  };
}

function validateLayoutBindings(manifest: PlatformManifest, sections: LayoutDefinition["sections"]): void {
  for (const section of sections) {
    requireNonEmptyText(section.title, "Section title");
    for (const component of section.components) {
      requireNonEmptyText(component.title, "Component title");
      const objectKey = component.binding?.objectKey ?? component.objectKey;
      const workflowKey = component.binding?.workflowKey ?? component.workflowKey;
      const agentId = component.binding?.agentId ?? component.agentId;
      const relatedObjectKey = component.binding?.relatedObjectKey ?? component.relatedObjectKey;

      if (["record_table", "record_form", "related_records"].includes(component.kind) && !objectKey) {
        throw new PlatformError(`Component "${component.title}" requires an object binding.`);
      }
      if (component.kind === "workflow_launcher" && !workflowKey) {
        throw new PlatformError(`Component "${component.title}" requires a workflow binding.`);
      }
      if (["agent_summary", "agent_panel"].includes(component.kind) && !agentId) {
        throw new PlatformError(`Component "${component.title}" requires an agent binding.`);
      }

      if (objectKey) {
        const objectDefinition = manifest.objects.find((candidate) => candidate.key === objectKey || candidate.id === objectKey);
        assertExists(objectDefinition, `Component object binding "${objectKey}" is invalid.`);
      }
      if (workflowKey) {
        const workflow = manifest.workflows.find((candidate) => candidate.key === workflowKey || candidate.id === workflowKey);
        assertExists(workflow, `Component workflow binding "${workflowKey}" is invalid.`);
      }
      if (agentId) {
        const agent = manifest.agents.find((candidate) => candidate.id === agentId || candidate.key === agentId);
        assertExists(agent, `Component agent binding "${agentId}" is invalid.`);
      }
      if (relatedObjectKey) {
        const relatedObject = manifest.objects.find(
          (candidate) => candidate.key === relatedObjectKey || candidate.id === relatedObjectKey,
        );
        assertExists(relatedObject, `Related object binding "${relatedObjectKey}" is invalid.`);
      }
    }
  }
}

function diffPreviewBucket<T>(input: {
  draftItems: T[];
  activeItems: T[];
  getKey: (item: T) => string;
  getLabel: (item: T) => string;
}) {
  const activeMap = new Map(input.activeItems.map((item) => [input.getKey(item), item]));
  const draftMap = new Map(input.draftItems.map((item) => [input.getKey(item), item]));

  const added = input.draftItems.filter((item) => !activeMap.has(input.getKey(item))).map(input.getLabel);
  const updated = input.draftItems
    .filter((item) => {
      const active = activeMap.get(input.getKey(item));
      return active && JSON.stringify(active) !== JSON.stringify(item);
    })
    .map(input.getLabel);
  const removed = input.activeItems.filter((item) => !draftMap.has(input.getKey(item))).map(input.getLabel);

  return {
    added,
    updated,
    removed,
  };
}

function buildPublishPageImpacts(input: {
  draftManifest: PlatformManifest;
  activeManifest: PlatformManifest | null;
}): PlatformPublishPageImpact[] {
  const activePages = new Map((input.activeManifest?.pages ?? []).map((page) => [page.key, page]));
  const activeLayouts = new Map((input.activeManifest?.layouts ?? []).map((layout) => [layout.key, layout]));
  const draftPages = new Map(input.draftManifest.pages.map((page) => [page.key, page]));
  const draftLayouts = new Map(input.draftManifest.layouts.map((layout) => [layout.key, layout]));

  const impacts = input.draftManifest.pages.map((page): PlatformPublishPageImpact | null => {
    const currentLayout = draftLayouts.get(page.layoutKey);
    const activePage = activePages.get(page.key);
    const activeLayout = activePage ? activeLayouts.get(activePage.layoutKey) : undefined;
    const status = !activePage
      ? ("added" as const)
      : JSON.stringify({ page, layout: currentLayout }) !== JSON.stringify({ page: activePage, layout: activeLayout })
        ? ("updated" as const)
        : null;

    if (!status) {
      return null;
    }

    return {
      pageKey: page.key,
      title: page.title,
      route: page.route,
      layoutKey: page.layoutKey,
      status,
      sectionCount: currentLayout?.sections.length ?? 0,
      componentCount: countLayoutComponents(currentLayout?.sections ?? []),
    };
  });

  const removed = (input.activeManifest?.pages ?? [])
    .filter((page) => !draftPages.has(page.key))
    .map((page) => {
      const layout = activeLayouts.get(page.layoutKey);
      return {
        pageKey: page.key,
        title: page.title,
        route: page.route,
        layoutKey: page.layoutKey,
        status: "removed" as const,
        sectionCount: layout?.sections.length ?? 0,
        componentCount: countLayoutComponents(layout?.sections ?? []),
      };
    });

  return [...impacts.filter((impact): impact is PlatformPublishPageImpact => impact !== null), ...removed];
}

function buildPublishRouteImpacts(input: {
  draftPages: PageDefinition[];
  activePages: PageDefinition[];
}): PlatformPublishRouteImpact[] {
  const activeMap = new Map(input.activePages.map((page) => [page.key, page]));
  const draftKeys = new Set(input.draftPages.map((page) => page.key));

  const changed: PlatformPublishRouteImpact[] = input.draftPages.flatMap((page): PlatformPublishRouteImpact[] => {
    const activePage = activeMap.get(page.key);
    if (!activePage) {
      return [{ pageKey: page.key, route: page.route, status: "added" as const }];
    }
    if (activePage.route !== page.route) {
      return [{ pageKey: page.key, route: page.route, status: "updated" as const }];
    }
    return [];
  });

  const removed = input.activePages
    .filter((page) => !draftKeys.has(page.key))
    .map((page) => ({
      pageKey: page.key,
      route: page.route,
      status: "removed" as const,
    }));

  return [...changed, ...removed];
}

export async function getPlatformBootstrap(input: {
  tenantSlug?: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformBootstrap> {
  const tenantSlug = input.tenantSlug ?? getEnv().PLATFORM_DEFAULT_TENANT_SLUG;
  const context = await getPlatformContextFromRequest({
    tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const [session, invites] = await Promise.all([
    getPlatformSessionSummary({
      request: input.request,
      fallbackActor: context.actor,
    }),
    listTenantInvitesForContext(context),
  ]);

  return {
    tenant: context.tenant,
    environment: context.environment,
    actor: context.actor,
    session,
    viewAs: tryResolvePlatformViewAsState(input.request),
    draftManifest: context.draftManifest,
    activeVersion: context.activeVersion,
    versions: context.versions,
    auditEvents: context.auditEvents,
    invites,
    designerCatalog: createDesignerCatalog(context.draftManifest),
  };
}

export async function listObjectDefinitions(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<ObjectDefinition[]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.objects;
}

export async function getCurrentPlatformSession(input?: {
  request?: Request | Headers;
}): Promise<PlatformSessionSummary> {
  return getPlatformSessionSummary({ request: input?.request });
}

export async function listTenantInvites(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformInviteRecord[]> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "SUPER_ADMIN");
  return listTenantInvitesForContext(context);
}

export async function createTenantInvite(input: {
  tenantSlug: string;
  email: string;
  role: PlatformRole;
  expiresInDays?: number;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformInviteRecord> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "SUPER_ADMIN");

  const expiresAt = new Date(Date.now() + (input.expiresInDays ?? 7) * 24 * 60 * 60 * 1000);
  const token = createInviteToken();
  const inviteUrl = buildInviteUrl(context.tenant.slug, token);

  const invite = await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const created = await prisma.platformTenantInvite.create({
        data: {
          tenantId: context.tenantId,
          email: input.email.trim().toLowerCase(),
          role: input.role,
          token,
          status: "pending",
          inviteUrl,
          createdByEmail: context.actor.email,
          expiresAt,
        },
        include: {
          tenant: true,
        },
      });
      return toInviteRecord(created, context.tenant.slug);
    },
    async () => {
      const created = await createLocalInvite({
        tenantId: context.tenantId,
        tenantSlug: context.tenant.slug,
        email: input.email.trim().toLowerCase(),
        role: input.role,
        token,
        inviteUrl,
        createdByEmail: context.actor.email,
        expiresAt: expiresAt.toISOString(),
      });
      return toInviteRecord(created, context.tenant.slug);
    },
  );

  await appendAuditEvent({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    actor: context.actor,
    action: "access.invite_created",
    resourceType: "invite",
    resourceId: invite.id,
    summary: `Created ${invite.role} invite for ${invite.email}.`,
    payload: {
      email: invite.email,
      role: invite.role,
      expiresAt: invite.expiresAt,
    },
  });

  return invite;
}

export async function acceptPlatformInvite(input: {
  token: string;
  email: string;
  name: string;
}): Promise<{
  sessionValue: string;
  tenantSlug: string;
  session: PlatformSessionSummary;
}> {
  const normalizedEmail = input.email.trim().toLowerCase();
  const normalizedName = requireNonEmptyText(input.name, "Display name");

  const accepted = await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const invite = await prisma.platformTenantInvite.findUnique({
        where: { token: input.token },
        include: {
          tenant: {
            include: {
              environments: {
                where: { isDefault: true },
                take: 1,
              },
            },
          },
        },
      });
      assertExists(invite, "Invite not found.");

      if (invite.email.toLowerCase() !== normalizedEmail) {
        throw new PlatformForbiddenError("Invite email does not match this session.");
      }
      if (invite.status === "revoked") {
        throw new PlatformForbiddenError("Invite has been revoked.");
      }
      if (invite.expiresAt.getTime() < Date.now()) {
        throw new PlatformForbiddenError("Invite has expired.");
      }

      const user = await prisma.platformUser.upsert({
        where: { email: normalizedEmail },
        update: {
          displayName: normalizedName,
        },
        create: {
          email: normalizedEmail,
          displayName: normalizedName,
        },
      });

      await prisma.platformMembership.upsert({
        where: {
          userId_tenantId: {
            userId: user.id,
            tenantId: invite.tenantId,
          },
        },
        update: {
          role: invite.role,
        },
        create: {
          tenantId: invite.tenantId,
          userId: user.id,
          role: invite.role,
        },
      });

      const updatedInvite = await prisma.platformTenantInvite.update({
        where: { id: invite.id },
        data: {
          status: "accepted",
          acceptedByUserId: user.id,
          acceptedAt: new Date(),
        },
        include: {
          tenant: true,
        },
      });

      return {
        invite: toInviteRecord(updatedInvite, invite.tenant.slug),
        actor: {
          email: user.email,
          name: user.displayName,
          role: invite.role as PlatformRole,
        } satisfies PlatformActor,
      };
    },
    async () => {
      const invite = await getLocalInviteByToken(input.token);
      assertExists(invite, "Invite not found.");

      if (invite.email.toLowerCase() !== normalizedEmail) {
        throw new PlatformForbiddenError("Invite email does not match this session.");
      }
      if (invite.status === "revoked") {
        throw new PlatformForbiddenError("Invite has been revoked.");
      }
      if (new Date(invite.expiresAt).getTime() < Date.now()) {
        throw new PlatformForbiddenError("Invite has expired.");
      }

      const user = (await getLocalUserByEmail(normalizedEmail)) ??
        (await createLocalUser({
          email: normalizedEmail,
          displayName: normalizedName,
        }));

      const membership = await getLocalMembership({
        tenantId: invite.tenantId,
        userId: user.id,
      });

      if (!membership) {
        await createLocalMembership({
          tenantId: invite.tenantId,
          userId: user.id,
          role: invite.role,
        });
      }

      const updatedInvite = await acceptLocalInvite({
        inviteId: invite.id,
        userId: user.id,
      });

      return {
        invite: toInviteRecord(updatedInvite, invite.tenantSlug),
        actor: {
          email: user.email,
          name: user.displayName,
          role: invite.role,
        } satisfies PlatformActor,
      };
    },
  );

  const sessionValue = createSignedPlatformSessionValue({
    email: accepted.actor.email,
    name: accepted.actor.name,
  });

  const session = await getPlatformSessionSummary({
    request: new Headers({
      cookie: `ta_platform_session=${encodeURIComponent(sessionValue)}`,
    }),
    fallbackActor: accepted.actor,
  });

  return {
    sessionValue,
    tenantSlug: accepted.invite.tenantSlug,
    session,
  };
}

export async function saveObjectDefinition(input: {
  tenantSlug: string;
  object: Partial<ObjectDefinition> & Pick<ObjectDefinition, "label" | "pluralLabel">;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<ObjectDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "object.saved",
    resourceType: "object",
    resourceId: input.object.id ?? input.object.key ?? input.object.label,
    summary: `Saved object definition ${input.object.label}.`,
    mutate: (manifest) => {
      const existingObject = manifest.objects.find(
        (candidate) => candidate.id === input.object.id || candidate.key === input.object.key,
      );
      const label = requireNonEmptyText(input.object.label, "Object label");
      const pluralLabel = requireNonEmptyText(input.object.pluralLabel, "Object plural label");
      const objectId = existingObject?.id ?? input.object.id ?? nextId("obj");
      const objectKey = requireManifestKey(input.object.key, label, "Object key");
      assertUniqueValue(manifest.objects, objectId, (candidate) => candidate.key, objectKey, "Object key");
      assertUniqueValue(manifest.objects, objectId, (candidate) => candidate.label, label, "Object label");
      const nextObject: ObjectDefinition = {
        id: objectId,
        key: objectKey,
        label,
        pluralLabel,
        description: normalizeOptionalText(input.object.description),
        icon: input.object.icon ?? existingObject?.icon ?? "layout",
        primaryFieldKey:
          normalizeOptionalText(input.object.primaryFieldKey) ??
          existingObject?.primaryFieldKey ??
          requireManifestKey(undefined, label, "Object primary field"),
        allowCreate: input.object.allowCreate ?? true,
        allowUpdate: input.object.allowUpdate ?? true,
        allowDelete: input.object.allowDelete ?? false,
        fields: input.object.fields ?? existingObject?.fields ?? [
          {
            id: nextId("fld"),
            key: requireManifestKey(undefined, label, "Object field key"),
            label,
            type: "text",
            required: true,
            unique: true,
            sensitivity: "internal",
            validations: [
              {
                id: nextId("val"),
                type: "required",
                message: `${label} is required.`,
              },
            ],
          },
        ],
        relationships: input.object.relationships ?? existingObject?.relationships ?? [],
        views: input.object.views ?? existingObject?.views ?? [],
      };

      if (nextObject.fields.length === 0) {
        throw new PlatformError("Objects must include at least one field.");
      }

      for (const field of nextObject.fields) {
        field.label = requireNonEmptyText(field.label, "Field label");
        field.key = requireManifestKey(field.key, field.label, "Field key");
      }

      const fieldKeySet = new Set<string>();
      const fieldLabelSet = new Set<string>();
      for (const field of nextObject.fields) {
        if (fieldKeySet.has(field.key.toLowerCase())) {
          throw new PlatformError(`Field key "${field.key}" must be unique within the object.`);
        }
        if (fieldLabelSet.has(field.label.toLowerCase())) {
          throw new PlatformError(`Field label "${field.label}" must be unique within the object.`);
        }
        fieldKeySet.add(field.key.toLowerCase());
        fieldLabelSet.add(field.label.toLowerCase());
      }

      nextObject.primaryFieldKey = nextObject.fields.some((field) => field.key === nextObject.primaryFieldKey)
        ? nextObject.primaryFieldKey
        : nextObject.fields[0]?.key ?? nextObject.primaryFieldKey;

      manifest.objects = upsertById(manifest.objects, nextObject);
      return {
        manifest,
        result: nextObject,
      };
    },
  });
}

export async function deleteObjectDefinition(input: {
  tenantSlug: string;
  objectId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<void> {
  await updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "object.deleted",
    resourceType: "object",
    resourceId: input.objectId,
    summary: `Deleted object definition ${input.objectId}.`,
    mutate: (manifest) => {
      const objectDefinition = manifest.objects.find((object) => object.id === input.objectId || object.key === input.objectId);
      assertExists(objectDefinition, "Object definition not found.");

      manifest.objects = manifest.objects.filter((object) => object.id !== objectDefinition.id);
      manifest.pages = manifest.pages.filter((page) => page.objectKey !== objectDefinition.key);
      manifest.layouts = manifest.layouts.filter((layout) => layout.pageKey && manifest.pages.some((page) => page.key === layout.pageKey));
      return {
        manifest,
        result: undefined,
      };
    },
  });
}

export async function listFieldDefinitions(input: {
  tenantSlug: string;
  objectId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<FieldDefinition[]> {
  const bootstrap = await getPlatformBootstrap(input);
  const objectDefinition = bootstrap.draftManifest.objects.find((object) => object.id === input.objectId || object.key === input.objectId);
  assertExists(objectDefinition, "Object definition not found.");

  return objectDefinition.fields;
}

export async function saveFieldDefinition(input: {
  tenantSlug: string;
  objectId: string;
  field: Partial<FieldDefinition> & Pick<FieldDefinition, "label" | "type">;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<FieldDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "field.saved",
    resourceType: "field",
    resourceId: input.field.id ?? input.field.key ?? input.field.label,
    summary: `Saved field ${input.field.label}.`,
    mutate: (manifest) => {
      const objectDefinition = manifest.objects.find((object) => object.id === input.objectId || object.key === input.objectId);
      assertExists(objectDefinition, "Object definition not found.");

      const fieldId = input.field.id ?? nextId("fld");
      const label = requireNonEmptyText(input.field.label, "Field label");
      const key = requireManifestKey(input.field.key, label, "Field key");
      assertUniqueValue(objectDefinition.fields, fieldId, (field) => field.key, key, "Field key");
      assertUniqueValue(objectDefinition.fields, fieldId, (field) => field.label, label, "Field label");
      const nextField: FieldDefinition = {
        id: fieldId,
        key,
        label,
        type: input.field.type,
        description: normalizeOptionalText(input.field.description),
        required: input.field.required ?? false,
        unique: input.field.unique ?? false,
        sensitivity: input.field.sensitivity ?? "internal",
        placeholder: normalizeOptionalText(input.field.placeholder),
        options: input.field.options,
        defaultValue: input.field.defaultValue,
        validations: input.field.validations ?? [],
        calculation: input.field.calculation ?? null,
      };

      objectDefinition.fields = upsertById(objectDefinition.fields, nextField);
      if (!objectDefinition.primaryFieldKey) {
        objectDefinition.primaryFieldKey = nextField.key;
      }

      return {
        manifest,
        result: nextField,
      };
    },
  });
}

export async function deleteFieldDefinition(input: {
  tenantSlug: string;
  objectId: string;
  fieldId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<void> {
  await updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "field.deleted",
    resourceType: "field",
    resourceId: input.fieldId,
    summary: `Deleted field ${input.fieldId}.`,
    mutate: (manifest) => {
      const objectDefinition = manifest.objects.find((object) => object.id === input.objectId || object.key === input.objectId);
      assertExists(objectDefinition, "Object definition not found.");

      objectDefinition.fields = objectDefinition.fields.filter((field) => field.id !== input.fieldId && field.key !== input.fieldId);
      if (!objectDefinition.fields.some((field) => field.key === objectDefinition.primaryFieldKey)) {
        objectDefinition.primaryFieldKey = objectDefinition.fields[0]?.key ?? objectDefinition.primaryFieldKey;
      }

      return {
        manifest,
        result: undefined,
      };
    },
  });
}

export async function listPageDefinitions(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PageDefinition[]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.pages;
}

export async function savePageDefinition(input: {
  tenantSlug: string;
  page: Partial<PageDefinition> & Pick<PageDefinition, "title">;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PageDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "page.saved",
    resourceType: "page",
    resourceId: input.page.id ?? input.page.key ?? input.page.title,
    summary: `Saved page ${input.page.title}.`,
    mutate: (manifest) => {
      const existingPage = manifest.pages.find((candidate) => candidate.id === input.page.id || candidate.key === input.page.key);
      const title = requireNonEmptyText(input.page.title, "Page title");
      const pageId = existingPage?.id ?? input.page.id ?? nextId("page");
      const pageKey = requireManifestKey(input.page.key, title, "Page key");
      const route = normalizeRoute(input.page.route, pageKey.replace(/_/g, "-"));
      const layoutKey = requireManifestKey(input.page.layoutKey, existingPage?.layoutKey ?? pageKey, "Layout key");
      assertUniqueValue(manifest.pages, pageId, (candidate) => candidate.key, pageKey, "Page key");
      assertUniqueValue(manifest.pages, pageId, (candidate) => candidate.title, title, "Page title");
      assertUniqueValue(manifest.pages, pageId, (candidate) => candidate.route, route, "Page route");

      if (input.page.objectKey) {
        const boundObject = manifest.objects.find((candidate) => candidate.key === input.page.objectKey || candidate.id === input.page.objectKey);
        assertExists(boundObject, "Page object binding is invalid.");
      }

      const nextPage: PageDefinition = {
        id: pageId,
        key: pageKey,
        title,
        route,
        description: normalizeOptionalText(input.page.description),
        layoutKey,
        objectKey: input.page.objectKey ?? existingPage?.objectKey,
        isHome: input.page.isHome ?? existingPage?.isHome ?? manifest.pages.length === 0,
        previewNote: normalizeOptionalText(input.page.previewNote),
      };

      manifest.pages = upsertById(manifest.pages, nextPage);
      if (nextPage.isHome) {
        manifest.pages = manifest.pages.map((page) =>
          page.id === nextPage.id
            ? page
            : {
                ...page,
                isHome: false,
              },
        );
      }

      if (!manifest.layouts.some((layout) => layout.key === layoutKey)) {
        manifest.layouts = upsertById(manifest.layouts, {
          id: nextId("layout"),
          key: layoutKey,
          name: title,
          pageKey,
          mobileColumns: 1,
          tabletColumns: 2,
          desktopColumns: 12,
          sections: [
            {
              id: nextId("section"),
              title,
              kind: "grid",
              columns: 12,
              placement: createDefaultPlacement({
                zone: "main",
                span: 12,
                minHeight: 420,
              }),
              components: [
                normalizeLayoutComponentDefinition({
                  id: nextId("component"),
                  kind: input.page.objectKey ? "record_table" : "rich_text",
                  title,
                  description: input.page.description,
                  objectKey: input.page.objectKey,
                  width: 12,
                  stylePreset: input.page.objectKey ? "record-grid" : "editorial",
                  placement: createDefaultPlacement({
                    zone: "main",
                    span: 12,
                    minHeight: input.page.objectKey ? 360 : 260,
                  }),
                  binding: input.page.objectKey
                    ? {
                        objectKey: input.page.objectKey,
                      }
                    : undefined,
                  props: input.page.objectKey ? {} : { body: "Add components to this page from the builder." },
                }),
              ],
            },
          ],
        });
      }

      return {
        manifest,
        result: nextPage,
      };
    },
  });
}

export async function listLayoutDefinitions(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<LayoutDefinition[]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.layouts;
}

export async function saveLayoutDefinition(input: {
  tenantSlug: string;
  layout: LayoutDefinition;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<LayoutDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "layout.saved",
    resourceType: "layout",
    resourceId: input.layout.id,
    summary: `Saved layout ${input.layout.name}.`,
    mutate: (manifest) => {
      const page = manifest.pages.find((candidate) => candidate.key === input.layout.pageKey || candidate.id === input.layout.pageKey);
      assertExists(page, "Layout page reference is invalid.");
      const layoutId = input.layout.id;
      const layoutName = requireNonEmptyText(input.layout.name, "Layout name");
      const layoutKey = requireManifestKey(input.layout.key, layoutName, "Layout key");
      assertUniqueValue(manifest.layouts, layoutId, (candidate) => candidate.key, layoutKey, "Layout key");
      const normalizedSections = input.layout.sections.map((section) =>
        normalizeLayoutSectionDefinition({
          ...section,
          placement: createDefaultPlacement({
            ...section.placement,
            span: section.placement?.span ?? 12,
          }),
          components: section.components.map((component) =>
            normalizeLayoutComponentDefinition({
              ...component,
              placement: createDefaultPlacement({
                ...component.placement,
                span: component.placement?.span ?? component.width,
              }),
            }),
          ),
        }),
      );
      validateLayoutBindings(manifest, normalizedSections);

      manifest.layouts = upsertById(manifest.layouts, {
        ...input.layout,
        name: layoutName,
        key: layoutKey,
        pageKey: page.key,
        sections: normalizedSections,
      });
      return {
        manifest,
        result: manifest.layouts[0]!,
      };
    },
  });
}

export async function listPageTemplates(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformManifest["pageTemplates"]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.pageTemplates;
}

export async function savePageTemplateDefinition(input: {
  tenantSlug: string;
  pageKey: string;
  templateKey?: string;
  label?: string;
  description?: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformManifest["pageTemplates"][number]> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "page_template.saved",
    resourceType: "page_template",
    resourceId: input.templateKey ?? input.pageKey,
    summary: `Saved reusable page template for ${input.pageKey}.`,
    mutate: (manifest) => {
      const template = createPageTemplateFromDraft({
        manifest,
        pageKey: input.pageKey,
        templateKey: input.templateKey,
        label: input.label,
        description: input.description,
      });
      manifest.pageTemplates = [
        template,
        ...manifest.pageTemplates.filter((candidate) => candidate.key !== template.key),
      ];
      return {
        manifest,
        result: template,
      };
    },
  });
}

export async function listSectionTemplates(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformSectionTemplate[]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.sectionTemplates;
}

export async function saveSectionTemplateDefinition(input: {
  tenantSlug: string;
  layoutKey: string;
  sectionId: string;
  templateKey?: string;
  label?: string;
  description?: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformSectionTemplate> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "section_template.saved",
    resourceType: "section_template",
    resourceId: input.templateKey ?? input.sectionId,
    summary: `Saved reusable section template for ${input.sectionId}.`,
    mutate: (manifest) => {
      const layout = manifest.layouts.find((candidate) => candidate.key === input.layoutKey || candidate.id === input.layoutKey);
      assertExists(layout, "Section template layout is invalid.");
      const template = createSectionTemplateFromDraft({
        layout,
        sectionId: input.sectionId,
        templateKey: input.templateKey,
        label: input.label,
        description: input.description,
      });
      manifest.sectionTemplates = [
        template,
        ...manifest.sectionTemplates.filter((candidate) => candidate.key !== template.key),
      ];
      return {
        manifest,
        result: template,
      };
    },
  });
}

export async function listMenuDefinitions(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<MenuItemDefinition[]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.menus.sort((left, right) => left.order - right.order);
}

export async function saveMenuDefinition(input: {
  tenantSlug: string;
  menu: Partial<MenuItemDefinition> & Pick<MenuItemDefinition, "label" | "pageKey">;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<MenuItemDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "menu.saved",
    resourceType: "menu",
    resourceId: input.menu.id ?? input.menu.key ?? input.menu.label,
    summary: `Saved menu item ${input.menu.label}.`,
    mutate: (manifest) => {
      const menuId = input.menu.id ?? nextId("menu");
      const label = requireNonEmptyText(input.menu.label, "Menu label");
      const page = manifest.pages.find((candidate) => candidate.key === input.menu.pageKey || candidate.id === input.menu.pageKey);
      assertExists(page, "Menu page reference is invalid.");
      const key = requireManifestKey(input.menu.key, label, "Menu key");
      assertUniqueValue(manifest.menus, menuId, (candidate) => candidate.key, key, "Menu key");
      assertUniqueValue(manifest.menus, menuId, (candidate) => candidate.label, label, "Menu label");
      assertUniqueValue(manifest.menus, menuId, (candidate) => candidate.pageKey, page.key, "Menu page");
      const menu: MenuItemDefinition = {
        id: menuId,
        key,
        label,
        description: normalizeOptionalText(input.menu.description),
        icon: input.menu.icon ?? "dot",
        pageKey: page.key,
        order: input.menu.order ?? manifest.menus.length,
        group: normalizeOptionalText(input.menu.group) ?? "Workspace",
        groupKey: normalizeOptionalText(input.menu.groupKey),
        visibleToRoles: input.menu.visibleToRoles?.length ? input.menu.visibleToRoles : ["SUPER_ADMIN", "BUILDER_ADMIN", "USER"],
        highlight: input.menu.highlight ?? false,
        badgeBindingKey: normalizeOptionalText(input.menu.badgeBindingKey),
      };

      manifest.menus = upsertById(manifest.menus, menu);
      return {
        manifest,
        result: menu,
      };
    },
  });
}

export async function getAppShellDefinition(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<AppShellDefinition> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.appShell;
}

export async function saveAppShellDefinition(input: {
  tenantSlug: string;
  appShell: Partial<AppShellDefinition> & Pick<AppShellDefinition, "productName">;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<AppShellDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "app_shell.saved",
    resourceType: "app_shell",
    resourceId: input.appShell.productName,
    summary: `Saved app shell ${input.appShell.productName}.`,
    mutate: (manifest) => {
      const nextAppShell: AppShellDefinition = {
        ...manifest.appShell,
        ...input.appShell,
        productName: requireNonEmptyText(input.appShell.productName, "Product name"),
        menuGroups: [...(input.appShell.menuGroups ?? [])].sort((left, right) => left.order - right.order),
        quickActions: input.appShell.quickActions ?? [],
        announcementSlots: input.appShell.announcementSlots ?? [],
        badgeBindings: input.appShell.badgeBindings ?? [],
        visibilityRules: input.appShell.visibilityRules ?? [],
        navigationMode: input.appShell.navigationMode ?? input.appShell.menuStyle ?? "sidebar",
      };

      manifest.appShell = nextAppShell;
      return {
        manifest,
        result: nextAppShell,
      };
    },
  });
}

export async function getNotificationCenterDefinition(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<NotificationCenterDefinition> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.notifications;
}

export async function saveNotificationCenterDefinition(input: {
  tenantSlug: string;
  notifications: NotificationCenterDefinition;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<NotificationCenterDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "notifications.saved",
    resourceType: "notification_center",
    resourceId: "notifications",
    summary: "Saved notification center configuration.",
    mutate: (manifest) => {
      for (const channel of input.notifications.channels) {
        requireNonEmptyText(channel.key, "Notification channel key");
        requireNonEmptyText(channel.name, "Notification channel name");
      }
      for (const template of input.notifications.templates) {
        requireNonEmptyText(template.key, "Notification template key");
        requireNonEmptyText(template.name, "Notification template name");
      }
      for (const rule of input.notifications.rules) {
        requireNonEmptyText(rule.key, "Notification rule key");
        requireNonEmptyText(rule.name, "Notification rule name");
      }

      manifest.notifications = input.notifications;
      return {
        manifest,
        result: manifest.notifications,
      };
    },
  });
}

export async function listWorkflowDefinitions(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<WorkflowDefinition[]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.workflows;
}

export async function saveWorkflowDefinition(input: {
  tenantSlug: string;
  workflow: Partial<WorkflowDefinition> & Pick<WorkflowDefinition, "name">;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<WorkflowDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "workflow.saved",
    resourceType: "workflow",
    resourceId: input.workflow.id ?? input.workflow.key ?? input.workflow.name,
    summary: `Saved workflow ${input.workflow.name}.`,
    mutate: (manifest) => {
      const workflowId = input.workflow.id ?? nextId("wf");
      const name = requireNonEmptyText(input.workflow.name, "Workflow name");
      const key = requireManifestKey(input.workflow.key, name, "Workflow key");
      assertUniqueValue(manifest.workflows, workflowId, (candidate) => candidate.key, key, "Workflow key");
      assertUniqueValue(manifest.workflows, workflowId, (candidate) => candidate.name, name, "Workflow name");

      if (input.workflow.objectKey) {
        const objectDefinition = manifest.objects.find((candidate) => candidate.key === input.workflow.objectKey || candidate.id === input.workflow.objectKey);
        assertExists(objectDefinition, "Workflow object binding is invalid.");
      }

      const nodes = (input.workflow.nodes ?? []).map((node) => ({
        ...node,
        label: requireNonEmptyText(node.label, "Workflow node label"),
      }));
      const triggers = (input.workflow.triggers ?? []).map((trigger) => ({
        ...trigger,
        label: requireNonEmptyText(trigger.label, "Workflow trigger label"),
      }));
      const edges = input.workflow.edges ?? [];
      const nodeIds = new Set(nodes.map((node) => node.id));
      for (const edge of edges) {
        if (!nodeIds.has(edge.sourceId) || !nodeIds.has(edge.targetId)) {
          throw new PlatformError("Workflow edges must connect existing nodes.");
        }
      }
      for (const node of nodes) {
        if (node.type === "subflow") {
          const subflowConfig = node.config as { workflowKey: string };
          const referencedWorkflow = manifest.workflows.find(
            (candidate) => candidate.key === subflowConfig.workflowKey || candidate.id === subflowConfig.workflowKey,
          );
          assertExists(referencedWorkflow, `Workflow subflow "${subflowConfig.workflowKey}" is invalid.`);
        }
      }

      const workflow: WorkflowDefinition = {
        id: workflowId,
        key,
        name,
        description: normalizeOptionalText(input.workflow.description),
        objectKey: input.workflow.objectKey,
        status: input.workflow.status ?? "draft",
        triggers,
        nodes,
        edges,
      };

      manifest.workflows = upsertById(manifest.workflows, workflow);
      return {
        manifest,
        result: workflow,
      };
    },
  });
}

export async function listWorkflowTemplates(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<WorkflowTemplateDefinition[]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.workflowTemplates;
}

export async function saveWorkflowTemplateDefinition(input: {
  tenantSlug: string;
  workflowId: string;
  templateKey?: string;
  name?: string;
  description?: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<WorkflowTemplateDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "workflow_template.saved",
    resourceType: "workflow_template",
    resourceId: input.templateKey ?? input.workflowId,
    summary: `Saved reusable workflow template for ${input.workflowId}.`,
    mutate: (manifest) => {
      const workflow = manifest.workflows.find((candidate) => candidate.id === input.workflowId || candidate.key === input.workflowId);
      assertExists(workflow, "Workflow template source workflow is invalid.");
      const template: WorkflowTemplateDefinition = {
        id: nextId("wf-template"),
        key: requireManifestKey(input.templateKey, workflow.key, "Workflow template key"),
        name: requireNonEmptyText(input.name ?? workflow.name, "Workflow template name"),
        description: normalizeOptionalText(input.description) ?? workflow.description ?? `Reusable template based on ${workflow.name}.`,
        source: "tenant",
        workflow: structuredClone(workflow),
      };
      manifest.workflowTemplates = [
        template,
        ...manifest.workflowTemplates.filter((candidate) => candidate.key !== template.key),
      ];
      return {
        manifest,
        result: template,
      };
    },
  });
}

export async function listSubflows(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<SubflowDefinition[]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.subflows;
}

export async function saveSubflowDefinition(input: {
  tenantSlug: string;
  subflow: Partial<SubflowDefinition> & Pick<SubflowDefinition, "name" | "workflowKey">;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<SubflowDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "subflow.saved",
    resourceType: "subflow",
    resourceId: input.subflow.id ?? input.subflow.key ?? input.subflow.name,
    summary: `Saved subflow ${input.subflow.name}.`,
    mutate: (manifest) => {
      const workflow = manifest.workflows.find(
        (candidate) => candidate.id === input.subflow.workflowKey || candidate.key === input.subflow.workflowKey,
      );
      assertExists(workflow, "Subflow workflow reference is invalid.");
      const subflowId = input.subflow.id ?? nextId("subflow");
      const name = requireNonEmptyText(input.subflow.name, "Subflow name");
      const key = requireManifestKey(input.subflow.key, name, "Subflow key");
      const nextSubflow: SubflowDefinition = {
        id: subflowId,
        key,
        name,
        description: normalizeOptionalText(input.subflow.description),
        workflowKey: workflow.key,
      };
      manifest.subflows = upsertById(manifest.subflows, nextSubflow);
      return {
        manifest,
        result: nextSubflow,
      };
    },
  });
}

export async function listWorkflowTests(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<WorkflowTestCaseDefinition[]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.workflowTests;
}

export async function saveWorkflowTestCaseDefinition(input: {
  tenantSlug: string;
  testCase: Partial<WorkflowTestCaseDefinition> & Pick<WorkflowTestCaseDefinition, "name" | "workflowKey">;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<WorkflowTestCaseDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "workflow_test.saved",
    resourceType: "workflow_test",
    resourceId: input.testCase.id ?? input.testCase.key ?? input.testCase.name,
    summary: `Saved workflow test ${input.testCase.name}.`,
    mutate: (manifest) => {
      const workflow = manifest.workflows.find(
        (candidate) => candidate.id === input.testCase.workflowKey || candidate.key === input.testCase.workflowKey,
      );
      assertExists(workflow, "Workflow test workflow reference is invalid.");
      const testCaseId = input.testCase.id ?? nextId("wf-testcase");
      const name = requireNonEmptyText(input.testCase.name, "Workflow test name");
      const key = requireManifestKey(input.testCase.key, name, "Workflow test key");
      const nextTestCase: WorkflowTestCaseDefinition = {
        id: testCaseId,
        key,
        name,
        workflowKey: workflow.key,
        description: normalizeOptionalText(input.testCase.description),
        payload: input.testCase.payload ?? {},
        expectedStatus: input.testCase.expectedStatus,
        expectedLogFragments: input.testCase.expectedLogFragments ?? [],
        expectApproval: input.testCase.expectApproval ?? false,
        expectWait: input.testCase.expectWait ?? false,
      };
      manifest.workflowTests = upsertById(manifest.workflowTests, nextTestCase);
      return {
        manifest,
        result: nextTestCase,
      };
    },
  });
}

export async function listAgentDefinitions(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<AgentDefinition[]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.agents;
}

export async function saveAgentDefinition(input: {
  tenantSlug: string;
  agent: Partial<AgentDefinition> & Pick<AgentDefinition, "name" | "modelProviderId" | "prompt" | "scope">;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<AgentDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "agent.saved",
    resourceType: "agent",
    resourceId: input.agent.id ?? input.agent.key ?? input.agent.name,
    summary: `Saved agent ${input.agent.name}.`,
    mutate: (manifest) => {
      const agentId = input.agent.id ?? nextId("agent");
      const name = requireNonEmptyText(input.agent.name, "Agent name");
      const key = requireManifestKey(input.agent.key, name, "Agent key");
      const prompt = requireNonEmptyText(input.agent.prompt, "Agent prompt");
      assertUniqueValue(manifest.agents, agentId, (candidate) => candidate.key, key, "Agent key");
      assertUniqueValue(manifest.agents, agentId, (candidate) => candidate.name, name, "Agent name");

      const provider = manifest.modelProviders.find(
        (candidate) => candidate.id === input.agent.modelProviderId || candidate.key === input.agent.modelProviderId,
      );
      assertExists(provider, "Agent model provider is invalid.");
      if (!manifest.securityPolicy.allowedModelProviderKeys.includes(provider.key)) {
        throw new PlatformError("Selected model provider is not allowed by tenant security policy.");
      }
      if ((input.agent.zeroRetentionRequired ?? true) && !provider.supportsZeroRetention) {
        throw new PlatformError("Selected model provider does not support zero retention.");
      }

      const allowedToolIds = input.agent.allowedToolIds ?? [];
      for (const toolId of allowedToolIds) {
        const tool = manifest.tools.find((candidate) => candidate.id === toolId || candidate.key === toolId);
        assertExists(tool, `Agent tool "${toolId}" is invalid.`);
      }

      const objectKeys = input.agent.objectKeys ?? [];
      for (const objectKey of objectKeys) {
        const objectDefinition = manifest.objects.find((candidate) => candidate.key === objectKey || candidate.id === objectKey);
        assertExists(objectDefinition, `Agent object scope "${objectKey}" is invalid.`);
      }

      const handoffWorkflowKeys = input.agent.handoffWorkflowKeys ?? [];
      for (const workflowKey of handoffWorkflowKeys) {
        const workflow = manifest.workflows.find((candidate) => candidate.id === workflowKey || candidate.key === workflowKey);
        assertExists(workflow, `Agent handoff workflow "${workflowKey}" is invalid.`);
      }

      const agent: AgentDefinition = {
        id: agentId,
        key,
        name,
        description: normalizeOptionalText(input.agent.description),
        scope: input.agent.scope,
        modelProviderId: provider.id,
        prompt,
        promptBlocks: (input.agent.promptBlocks ?? []).map((block) => ({
          ...block,
          id: block.id || nextId("prompt-block"),
          label: requireNonEmptyText(block.label, "Prompt block label"),
          content: requireNonEmptyText(block.content, "Prompt block content"),
        })),
        allowedToolIds,
        objectKeys,
        handoffWorkflowKeys,
        outputSchema: normalizeOptionalText(input.agent.outputSchema),
        evalPolicy: input.agent.evalPolicy
          ? {
              rubric: requireNonEmptyText(input.agent.evalPolicy.rubric, "Agent eval rubric"),
              samplePrompt: requireNonEmptyText(input.agent.evalPolicy.samplePrompt, "Agent sample prompt"),
              passingScore: input.agent.evalPolicy.passingScore,
            }
          : undefined,
        costBudgetUsd: input.agent.costBudgetUsd,
        approvalPolicy: input.agent.approvalPolicy
          ? {
              required: input.agent.approvalPolicy.required,
              approverRole: input.agent.approvalPolicy.approverRole,
              notes: normalizeOptionalText(input.agent.approvalPolicy.notes),
            }
          : undefined,
        zeroRetentionRequired: input.agent.zeroRetentionRequired ?? true,
      };

      manifest.agents = upsertById(manifest.agents, agent);
      return {
        manifest,
        result: agent,
      };
    },
  });
}

export async function listModelProviders(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<ModelProviderDefinition[]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.modelProviders;
}

export async function saveModelProviderDefinition(input: {
  tenantSlug: string;
  provider: ModelProviderDefinition;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<ModelProviderDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    minimumRole: "SUPER_ADMIN",
    action: "model.saved",
    resourceType: "model_provider",
    resourceId: input.provider.id,
    summary: `Saved model provider ${input.provider.name}.`,
    mutate: (manifest) => {
      const providerId = input.provider.id;
      const name = requireNonEmptyText(input.provider.name, "Model provider name");
      const key = requireManifestKey(input.provider.key, name, "Model provider key");
      assertUniqueValue(manifest.modelProviders, providerId, (candidate) => candidate.key, key, "Model provider key");
      assertUniqueValue(manifest.modelProviders, providerId, (candidate) => candidate.name, name, "Model provider name");

      const provider: ModelProviderDefinition = {
        ...input.provider,
        name,
        key,
        model: requireNonEmptyText(input.provider.model, "Model name"),
        apiKeySecretRef: requireNonEmptyText(input.provider.apiKeySecretRef, "API key secret reference"),
        endpoint: normalizeOptionalText(input.provider.endpoint),
      };

      manifest.modelProviders = upsertById(manifest.modelProviders, provider);
      return {
        manifest,
        result: provider,
      };
    },
  });
}

export async function getSecurityPolicy(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<SecurityPolicyDefinition> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.securityPolicy;
}

export async function saveSecurityPolicy(input: {
  tenantSlug: string;
  securityPolicy: SecurityPolicyDefinition;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<SecurityPolicyDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    minimumRole: "SUPER_ADMIN",
    action: "security.saved",
    resourceType: "security_policy",
    resourceId: "default",
    summary: "Updated tenant security policy.",
    mutate: (manifest) => {
      const maskingPolicy = manifest.maskingPolicies.find((candidate) => candidate.key === input.securityPolicy.defaultMaskingPolicyKey);
      assertExists(maskingPolicy, "Default masking policy is invalid.");
      for (const providerKey of input.securityPolicy.allowedModelProviderKeys) {
        const provider = manifest.modelProviders.find((candidate) => candidate.key === providerKey);
        assertExists(provider, `Allowed model provider "${providerKey}" is invalid.`);
      }

      manifest.securityPolicy = {
        ...input.securityPolicy,
        allowedModelProviderKeys: [...new Set(input.securityPolicy.allowedModelProviderKeys)],
      };
      return {
        manifest,
        result: manifest.securityPolicy,
      };
    },
  });
}

export async function publishDraftManifest(input: {
  tenantSlug: string;
  notes?: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformPublishedVersionRecord> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "BUILDER_ADMIN");
  const nextVersionNumber = Math.max(0, ...context.versions.map((version) => version.versionNumber)) + 1;
  const versionId = nextId("version");
  const artifacts = await buildPublishedArtifacts({
    manifest: context.draftManifest,
    tenantSlug: context.tenant.slug,
    environmentSlug: context.environment.slug,
    versionNumber: nextVersionNumber,
    versionId,
  });

  await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      await prisma.$transaction(async (tx) => {
        await tx.platformPublishedVersion.updateMany({
          where: {
            tenantId: context.tenantId,
            environmentId: context.environmentId,
            status: "ACTIVE",
          },
          data: {
            status: "ROLLED_BACK",
          },
        });

        await tx.platformPublishedVersion.create({
          data: {
            id: versionId,
            tenantId: context.tenantId,
            environmentId: context.environmentId,
            versionNumber: nextVersionNumber,
            status: "ACTIVE",
            manifest: toJsonValue(artifacts.manifest),
            manifestPath: artifacts.manifestPath,
            gitCommitSha: artifacts.gitCommitSha,
            notes: input.notes ?? null,
            activatedAt: new Date(),
          },
        });

        const existingDraft = await tx.platformDraft.findUnique({
          where: {
            tenantId_environmentId: {
              tenantId: context.tenantId,
              environmentId: context.environmentId,
            },
          },
        });

        if (existingDraft) {
          await tx.platformDraft.update({
            where: { id: existingDraft.id },
            data: {
              manifest: toJsonValue(artifacts.manifest),
            },
          });
        }
      });
    },
    async () => {
      const version = await createLocalVersion({
        id: versionId,
        tenantId: context.tenantId,
        environmentId: context.environmentId,
        versionNumber: nextVersionNumber,
        manifest: artifacts.manifest,
        status: "ACTIVE",
        notes: input.notes ?? null,
        manifestPath: artifacts.manifestPath,
        gitCommitSha: artifacts.gitCommitSha,
        activatedAt: new Date().toISOString(),
      });

      await setLocalActiveVersion({
        tenantId: context.tenantId,
        environmentId: context.environmentId,
        versionId: version.id,
      });

      await upsertLocalDraft({
        tenantId: context.tenantId,
        environmentId: context.environmentId,
        manifest: artifacts.manifest,
      });
    },
  );

  await appendAuditEvent({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    actor: context.actor,
    action: "publish.created",
    resourceType: "published_version",
    resourceId: versionId,
    summary: `Published version v${nextVersionNumber}.`,
    payload: {
      versionNumber: nextVersionNumber,
      manifestPath: artifacts.manifestPath,
      gitCommitSha: artifacts.gitCommitSha,
    },
  });

  return {
    id: versionId,
    versionNumber: nextVersionNumber,
    status: "ACTIVE",
    notes: input.notes ?? null,
    manifestPath: artifacts.manifestPath,
    gitCommitSha: artifacts.gitCommitSha,
    activatedAt: new Date().toISOString(),
    createdAt: new Date().toISOString(),
  };
}

export async function listPublishedVersions(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformPublishedVersionRecord[]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.versions;
}

export async function rollbackPublishedVersion(input: {
  tenantSlug: string;
  versionId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformPublishedVersionRecord> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "BUILDER_ADMIN");
  const versionToActivate = context.versions.find((version) => version.id === input.versionId);
  assertExists(versionToActivate, "Published version not found.");

  await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      await prisma.$transaction(async (tx) => {
        await tx.platformPublishedVersion.updateMany({
          where: {
            tenantId: context.tenantId,
            environmentId: context.environmentId,
            status: "ACTIVE",
          },
          data: {
            status: "ROLLED_BACK",
          },
        });

        await tx.platformPublishedVersion.update({
          where: { id: input.versionId },
          data: {
            status: "ACTIVE",
            activatedAt: new Date(),
          },
        });
      });
    },
    async () => {
      await setLocalActiveVersion({
        tenantId: context.tenantId,
        environmentId: context.environmentId,
        versionId: input.versionId,
      });
    },
  );

  await appendAuditEvent({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    actor: context.actor,
    action: "publish.rollback",
    resourceType: "published_version",
    resourceId: input.versionId,
    summary: `Rolled back to v${versionToActivate.versionNumber}.`,
    payload: {
      versionNumber: versionToActivate.versionNumber,
    },
  });

  return {
    ...versionToActivate,
    status: "ACTIVE",
    activatedAt: new Date().toISOString(),
  };
}

export async function listAuditEvents(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformAuditEventRecord[]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.auditEvents;
}

async function loadActiveRuntimeManifest(context: PlatformContext): Promise<PlatformManifest | null> {
  if (!context.activeVersion) {
    return null;
  }

  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const version = await prisma.platformPublishedVersion.findUnique({
        where: { id: context.activeVersion!.id },
      });

      return version?.manifest ? ensureManifestConsistency(version.manifest as unknown as PlatformManifest) : null;
    },
    async () => {
      const versions = await listLocalVersions({
        tenantId: context.tenantId,
        environmentId: context.environmentId,
      });
      const manifest = versions.find((version) => version.id === context.activeVersion!.id)?.manifest;
      return manifest ? ensureManifestConsistency(manifest) : null;
    },
  );
}

export async function getPublicRuntimeManifest(input: {
  tenantSlug: string;
  environmentSlug?: string;
}): Promise<PlatformManifest | null> {
  const environmentSlug = input.environmentSlug ?? getEnv().PLATFORM_DEFAULT_ENVIRONMENT_SLUG;

  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const tenant = await prisma.platformTenant.findUnique({
        where: { slug: input.tenantSlug },
      });
      if (!tenant) {
        return null;
      }

      const environment = await prisma.platformEnvironment.findUnique({
        where: {
          tenantId_slug: {
            tenantId: tenant.id,
            slug: environmentSlug,
          },
        },
      });
      if (!environment) {
        return null;
      }

      const version = await prisma.platformPublishedVersion.findFirst({
        where: {
          tenantId: tenant.id,
          environmentId: environment.id,
          status: "ACTIVE",
        },
        orderBy: { createdAt: "desc" },
      });

      return version?.manifest ? ensureManifestConsistency(version.manifest as unknown as PlatformManifest) : null;
    },
    async () => {
      const tenant = await getLocalTenantBySlug(input.tenantSlug);
      if (!tenant) {
        return null;
      }

      const environment = await getLocalEnvironmentBySlug({
        tenantId: tenant.id,
        slug: environmentSlug,
      });
      if (!environment) {
        return null;
      }

      const activeVersion = await getLocalActiveVersion({
        tenantId: tenant.id,
        environmentId: environment.id,
      });
      if (!activeVersion) {
        return null;
      }

      const versions = await listLocalVersions({
        tenantId: tenant.id,
        environmentId: environment.id,
      });
      const manifest = versions.find((version) => version.id === activeVersion.id)?.manifest;
      return manifest ? ensureManifestConsistency(manifest) : null;
    },
  );
}

export async function getRuntimeManifest(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
  draftFallback?: boolean;
}): Promise<PlatformManifest | null> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });

  const activeManifest = await loadActiveRuntimeManifest(context);
  if (activeManifest) {
    return activeManifest;
  }

  return input.draftFallback ? context.draftManifest : null;
}

export async function getDraftPreviewManifest(input: {
  tenantSlug: string;
  route?: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformManifest> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "BUILDER_ADMIN");

  await appendAuditEvent({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    actor: context.actor,
    action: "preview.opened",
    resourceType: "draft_preview",
    resourceId: input.route ?? "workspace",
    summary: `Opened draft preview${input.route ? ` for ${input.route}` : ""}.`,
    payload: {
      route: input.route ?? null,
      manifestUpdatedAt: context.draftManifest.metadata.draftUpdatedAt,
    },
  });

  return context.draftManifest;
}

export async function saveBrandingDefinition(input: {
  tenantSlug: string;
  branding: Partial<TenantBrandingDefinition> & Pick<TenantBrandingDefinition, "themeName" | "primaryColor" | "secondaryColor" | "accentColor" | "surfaceColor" | "textColor" | "pageBackground" | "fontFamily">;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<TenantBrandingDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "branding.saved",
    resourceType: "branding",
    resourceId: "tenant-branding",
    summary: `Saved branding theme ${input.branding.themeName}.`,
    mutate: (manifest) => {
      const nextBranding: TenantBrandingDefinition = {
        ...createDefaultBranding(manifest.tenant.name),
        ...manifest.branding,
        ...input.branding,
        themeName: requireNonEmptyText(input.branding.themeName, "Theme name"),
        primaryColor: requireNonEmptyText(input.branding.primaryColor, "Primary color"),
        secondaryColor: requireNonEmptyText(input.branding.secondaryColor, "Secondary color"),
        accentColor: requireNonEmptyText(input.branding.accentColor, "Accent color"),
        surfaceColor: requireNonEmptyText(input.branding.surfaceColor, "Surface color"),
        textColor: requireNonEmptyText(input.branding.textColor, "Text color"),
        pageBackground: requireNonEmptyText(input.branding.pageBackground, "Page background"),
        fontFamily: requireNonEmptyText(input.branding.fontFamily, "Font family"),
        notes: normalizeOptionalText(input.branding.notes),
        mode: input.branding.mode ?? manifest.branding?.mode ?? "draft",
        assets: manifest.branding?.assets ?? [],
      };

      manifest.branding = nextBranding;
      return {
        manifest,
        result: nextBranding,
      };
    },
  });
}

export async function uploadBrandAsset(input: {
  tenantSlug: string;
  kind: "logo" | "icon" | "brand_book" | "reference";
  label: string;
  fileName: string;
  contentType: string;
  buffer: Buffer;
  environmentSlug?: string;
  request?: Request | Headers;
}) {
  const fileName = input.fileName.replace(/[^a-zA-Z0-9._-]+/g, "-");
  const storageKey = `platform-assets/${input.tenantSlug}/${nanoid(10)}/${fileName}`;
  await getObjectStorage().putBuffer(storageKey, input.buffer, input.contentType);

  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "branding.asset_uploaded",
    resourceType: "branding_asset",
    resourceId: storageKey,
    summary: `Uploaded ${input.kind.replace(/_/g, " ")} asset ${input.label}.`,
    mutate: (manifest) => {
      const assetId = nextId("brand-asset");
      const asset = {
        id: assetId,
        kind: input.kind,
        label: requireNonEmptyText(input.label, "Asset label"),
        fileName,
        contentType: input.contentType,
        storageKey,
        url: `/api/platform/tenants/${input.tenantSlug}/branding/assets/${assetId}`,
        uploadedAt: new Date().toISOString(),
      };

      manifest.branding = {
        ...createDefaultBranding(manifest.tenant.name),
        ...manifest.branding,
        assets: [asset, ...(manifest.branding?.assets ?? [])],
        logoAssetId: input.kind === "logo" ? assetId : manifest.branding?.logoAssetId,
        iconAssetId: input.kind === "icon" ? assetId : manifest.branding?.iconAssetId,
        brandBookAssetId: input.kind === "brand_book" ? assetId : manifest.branding?.brandBookAssetId,
      };

      return {
        manifest,
        result: asset,
      };
    },
  });
}

export async function saveProfileConfiguration(input: {
  tenantSlug: string;
  pageTitle: string;
  visibleFieldKeys: string[];
  profilePageKey?: string;
  settingsPageKey?: string;
  environmentSlug?: string;
  request?: Request | Headers;
}) {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "profiles.saved",
    resourceType: "profiles",
    resourceId: "tenant-profile-config",
    summary: `Saved profile settings for ${input.tenantSlug}.`,
    mutate: (manifest) => {
      manifest.profiles = {
        ...manifest.profiles,
        pageTitle: requireNonEmptyText(input.pageTitle, "Profile page title"),
        visibleFieldKeys: input.visibleFieldKeys,
        profilePageKey: normalizeOptionalText(input.profilePageKey),
        settingsPageKey: normalizeOptionalText(input.settingsPageKey),
      };

      manifest.appShell = {
        ...manifest.appShell,
        profilePageKey: normalizeOptionalText(input.profilePageKey) ?? manifest.appShell.profilePageKey,
        settingsPageKey: normalizeOptionalText(input.settingsPageKey) ?? manifest.appShell.settingsPageKey,
      };

      return {
        manifest,
        result: manifest.profiles,
      };
    },
  });
}

export async function listFormDefinitions(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<FormDefinition[]> {
  const bootstrap = await getPlatformBootstrap(input);
  return bootstrap.draftManifest.forms;
}

export async function saveFormDefinition(input: {
  tenantSlug: string;
  form: Partial<EditableFormInput> & Pick<EditableFormInput, "title" | "deliveryMode" | "submitLabel" | "successMessage">;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<FormDefinition> {
  return updateDraftWithMutation({
    tenantSlug: input.tenantSlug,
    request: input.request,
    environmentSlug: input.environmentSlug,
    action: "form.saved",
    resourceType: "form",
    resourceId: input.form.id ?? input.form.key ?? input.form.title,
    summary: `Saved form ${input.form.title}.`,
    mutate: (manifest) => {
      const formId = input.form.id ?? nextId("form");
      const title = requireNonEmptyText(input.form.title, "Form title");
      const key = requireManifestKey(input.form.key, title, "Form key");
      const route = normalizeRoute(input.form.route, key);
      assertUniqueValue(manifest.forms, formId, (form) => form.key, key, "Form key");
      assertUniqueValue(manifest.forms, formId, (form) => form.route, route, "Form route");

      const fields: FormFieldDefinition[] = (input.form.fields ?? []).map((field, index) => {
        const label = requireNonEmptyText(field.label, `Form field ${index + 1} label`);
        const fieldKey = requireManifestKey(field.key, label, `Form field ${index + 1} key`);
        return {
          id: field.id ?? nextId("form-field"),
          key: fieldKey,
          label,
          type: field.type,
          description: normalizeOptionalText(field.description),
          helpText: normalizeOptionalText(field.helpText),
          tooltip: normalizeOptionalText(field.tooltip),
          required: field.required ?? false,
          placeholder: normalizeOptionalText(field.placeholder),
          options: field.options ?? [],
          defaultValue: field.defaultValue,
          validations: field.validations ?? [],
          calculation: field.calculation ?? null,
          mandatoryRule: field.mandatoryRule,
        };
      });

      const steps: FormStepDefinition[] = (input.form.steps ?? []).map((step, index) => ({
        id: step.id ?? nextId("form-step"),
        key: requireManifestKey(step.key, step.title, `Form step ${index + 1} key`),
        title: requireNonEmptyText(step.title, `Form step ${index + 1} title`),
        description: normalizeOptionalText(step.description),
        fieldKeys: step.fieldKeys ?? [],
        visibilityRule: step.visibilityRule,
      }));

      const nextForm: FormDefinition = {
        id: formId,
        key,
        title,
        description: normalizeOptionalText(input.form.description),
        route,
        objectKey: normalizeOptionalText(input.form.objectKey),
        deliveryMode: input.form.deliveryMode,
        submitLabel: requireNonEmptyText(input.form.submitLabel, "Submit label"),
        successMessage: requireNonEmptyText(input.form.successMessage, "Success message"),
        saveAndResume: input.form.saveAndResume ?? true,
        requireAuthentication: input.form.requireAuthentication ?? false,
        analyticsEnabled: input.form.analyticsEnabled ?? true,
        fields,
        steps,
      };

      manifest.forms = upsertById(manifest.forms, nextForm);
      return {
        manifest,
        result: nextForm,
      };
    },
  });
}

export async function listFormSubmissions(input: {
  tenantSlug: string;
  formKey: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformFormSubmissionRecord[]> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "BUILDER_ADMIN");
  const form = findFormDefinition(context.draftManifest, input.formKey);
  assertExists(form, "Form definition not found.");
  const objectKey = getFormSubmissionObjectKey(form.key);

  const records = await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const rows = await prisma.platformRecord.findMany({
        where: {
          tenantId: context.tenantId,
          environmentId: context.environmentId,
          objectKey,
        },
        orderBy: [{ createdAt: "desc" }],
        take: 50,
      });
      return rows.map(toPlatformRecord);
    },
    async () => {
      const rows = await listLocalPlatformRecords({
        tenantId: context.tenantId,
        environmentId: context.environmentId,
        objectKey,
      });
      return rows.map(toPlatformRecord);
    },
  );

  return records.map(toFormSubmissionRecord);
}

export async function submitFormSubmission(input: {
  tenantSlug: string;
  formKey: string;
  data: Record<string, unknown>;
  status?: PlatformFormSubmissionRecord["status"];
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformFormSubmissionRecord> {
  const manifest = await getPublicRuntimeManifest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
  });
  if (!manifest) {
    throw new PlatformError("No published runtime is active for this tenant.", 409);
  }

  const form = findFormDefinition(manifest, input.formKey);
  assertExists(form, "Form definition not found in the active runtime.");

  const identity = tryResolvePlatformActorIdentity(input.request);
  if (form.requireAuthentication && !identity) {
    throw new PlatformForbiddenError("Authentication is required to submit this form.");
  }

  const formAsObject = {
    id: `obj-${form.key}`,
    key: form.key,
    label: form.title,
    pluralLabel: `${form.title}s`,
    icon: "form",
    primaryFieldKey: form.fields[0]?.key ?? "submission",
    allowCreate: true,
    allowUpdate: true,
    allowDelete: false,
    fields: form.fields.map((field) => ({
      ...field,
      unique: false,
      sensitivity: "public" as const,
    })),
    relationships: [],
    views: [],
  };
  const validation = validateRecordInput({
    objectDefinition: formAsObject,
    manifest,
    rawData: input.data,
    existingRecords: [],
  });

  if (validation.errors.length > 0) {
    throw new PlatformError(validation.errors.join(" "));
  }

  const environmentSlug = input.environmentSlug ?? manifest.environment.slug;
  const submissionObjectKey = getFormSubmissionObjectKey(form.key);
  const actorEmail = identity?.email ?? null;

  const saved = await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const tenant = await prisma.platformTenant.findUnique({
        where: { slug: input.tenantSlug },
      });
      assertExists(tenant, "Platform tenant not found.");
      const environment = await prisma.platformEnvironment.findUnique({
        where: {
          tenantId_slug: {
            tenantId: tenant.id,
            slug: environmentSlug,
          },
        },
      });
      assertExists(environment, "Platform environment not found.");
      const row = await prisma.platformRecord.create({
        data: {
          tenantId: tenant.id,
          environmentId: environment.id,
          objectKey: submissionObjectKey,
          data: toJsonValue({
            formKey: form.key,
            objectKey: form.objectKey,
            status: input.status ?? "submitted",
            submission: validation.data,
            submittedAt: new Date().toISOString(),
            createdByEmail: actorEmail,
          }),
          createdByEmail: actorEmail,
          updatedByEmail: actorEmail,
        },
      });
      return toPlatformRecord(row);
    },
    async () => {
      const tenant = await getLocalTenantBySlug(input.tenantSlug);
      assertExists(tenant, "Platform tenant not found.");
      const environment = await getLocalEnvironmentBySlug({
        tenantId: tenant.id,
        slug: environmentSlug,
      });
      assertExists(environment, "Platform environment not found.");
      const row = await upsertLocalPlatformRecord({
        tenantId: tenant.id,
        environmentId: environment.id,
        objectKey: submissionObjectKey,
        data: {
          formKey: form.key,
          objectKey: form.objectKey,
          status: input.status ?? "submitted",
          submission: validation.data,
          submittedAt: new Date().toISOString(),
          createdByEmail: actorEmail,
        },
        actor: {
          email: actorEmail ?? "public@tenant.local",
          name: actorEmail ?? "Public submitter",
          role: "USER",
        },
      });
      return toPlatformRecord(row);
    },
  );

  await emitPlatformEvent({
    tenantSlug: input.tenantSlug,
    environmentSlug,
    type: "form.submission.received",
    source: "form",
    resourceType: "form_submission",
    resourceId: saved.id,
    payload: {
      formKey: form.key,
      status: input.status ?? "submitted",
    },
  });

  return toFormSubmissionRecord(saved);
}

export async function runWorkflowTest(input: {
  tenantSlug: string;
  workflowId: string;
  testCaseId?: string;
  payload?: Record<string, unknown>;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformWorkflowRunRecord> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "BUILDER_ADMIN");

  const workflow = context.draftManifest.workflows.find(
    (candidate) => candidate.id === input.workflowId || candidate.key === input.workflowId,
  );
  assertExists(workflow, "Workflow definition not found.");
  const savedTestCase = input.testCaseId
    ? context.draftManifest.workflowTests.find((candidate) => candidate.id === input.testCaseId || candidate.key === input.testCaseId)
    : undefined;
  if (savedTestCase && savedTestCase.workflowKey !== workflow.key) {
    throw new PlatformError("Workflow test case does not belong to the selected workflow.");
  }
  const payload = input.payload ?? savedTestCase?.payload ?? {};

  const logs = [
    {
      level: "info",
      message: `Started draft test for ${workflow.name}.`,
      at: new Date().toISOString(),
    },
    ...workflow.nodes
      .sort((left, right) => left.position.y - right.position.y || left.position.x - right.position.x)
      .map((node) => ({
        level: "info",
        message: `Simulated ${node.type} node "${node.label}".`,
        at: new Date().toISOString(),
      })),
    {
      level: "info",
      message: `Completed draft test for ${workflow.name}.`,
      at: new Date().toISOString(),
    },
  ];

  await appendAuditEvent({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    actor: context.actor,
    action: "workflow.test_ran",
    resourceType: "workflow",
    resourceId: workflow.id,
    summary: `Ran draft test for workflow ${workflow.name}.`,
    payload: {
      payload,
      testCaseId: savedTestCase?.id ?? null,
    },
  });

  return {
    id: nextId("wf-test"),
    workflowId: workflow.id,
    workflowKey: workflow.key,
    status: "SUCCEEDED",
    input: payload,
    output: {
      completedNodes: workflow.nodes.length,
      mode: "draft-test",
      testCaseId: savedTestCase?.id ?? null,
    },
    logs,
    startedAt: new Date().toISOString(),
    finishedAt: new Date().toISOString(),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}

export async function evaluateAgentDefinition(input: {
  tenantSlug: string;
  agentId: string;
  objectKey?: string;
  sampleSize?: number;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformAgentEvalRecord> {
  const preview = await previewAgentInvocation({
    tenantSlug: input.tenantSlug,
    agentId: input.agentId,
    objectKey: input.objectKey,
    sampleSize: input.sampleSize,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });

  const score = [
    preview.metadata.allowedByPolicy,
    preview.metadata.masked,
    !preview.agent.zeroRetentionRequired || preview.provider.supportsZeroRetention,
  ].filter(Boolean).length * 33;

  return {
    id: nextId("agent-eval"),
    agentId: preview.agent.id,
    agentKey: preview.agent.key,
    score: Math.min(score, 100),
    summary:
      score >= 99
        ? "Provider policy, masking, and retention posture all passed for the sampled invocation."
        : "The sampled invocation surfaced one or more policy or retention concerns.",
    createdAt: new Date().toISOString(),
    result: {
      metadata: preview.metadata,
      provider: preview.provider,
      sampleSize: preview.sampleSize,
    },
  };
}

export async function listPlatformRecords(input: {
  tenantSlug: string;
  objectKey: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformRecord[]> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const runtimeManifest = await getRuntimeManifest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  if (!runtimeManifest) {
    throw new PlatformError("No published runtime is active for this tenant.", 409);
  }

  const objectDefinition = findObjectDefinition(runtimeManifest, input.objectKey);
  assertExists(objectDefinition, "Object definition not found in the active runtime.");

  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const records = await prisma.platformRecord.findMany({
        where: {
          tenantId: context.tenantId,
          environmentId: context.environmentId,
          objectKey: objectDefinition.key,
        },
        orderBy: [{ updatedAt: "desc" }],
      });

      return records.map(toPlatformRecord);
    },
    async () => {
      const records = await listLocalPlatformRecords({
        tenantId: context.tenantId,
        environmentId: context.environmentId,
        objectKey: objectDefinition.key,
      });
      return records.map(toPlatformRecord);
    },
  );
}

export async function savePlatformRecord(input: {
  tenantSlug: string;
  objectKey: string;
  data: Record<string, unknown>;
  recordId?: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformRecord> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const runtimeManifest = await getRuntimeManifest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });

  if (!runtimeManifest) {
    throw new PlatformError("No published runtime is active for this tenant.", 409);
  }

  const objectDefinition = findObjectDefinition(runtimeManifest, input.objectKey);
  assertExists(objectDefinition, "Object definition not found in the active runtime.");

  const existingRecords = await listPlatformRecords({
    tenantSlug: input.tenantSlug,
    objectKey: objectDefinition.key,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const validation = validateRecordInput({
    objectDefinition,
    manifest: runtimeManifest,
    rawData: input.data,
    existingRecords,
    currentRecordId: input.recordId,
  });

  if (validation.errors.length > 0) {
    throw new PlatformError(validation.errors.join(" "));
  }

  const saved = await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const record = input.recordId
        ? await prisma.platformRecord.findFirst({
            where: {
              id: input.recordId,
              tenantId: context.tenantId,
              environmentId: context.environmentId,
              objectKey: objectDefinition.key,
            },
          })
        : null;

      if (input.recordId && !record) {
        throw new PlatformNotFoundError("Record not found in the active tenant scope.");
      }

      const savedRecord = record
        ? await prisma.platformRecord.update({
            where: { id: record.id },
            data: {
              data: toJsonValue(validation.data),
              updatedByEmail: context.actor.email,
            },
          })
        : await prisma.platformRecord.create({
            data: {
              tenantId: context.tenantId,
              environmentId: context.environmentId,
              objectKey: objectDefinition.key,
              data: toJsonValue(validation.data),
              createdByEmail: context.actor.email,
              updatedByEmail: context.actor.email,
            },
          });

      return toPlatformRecord(savedRecord);
    },
    async () => {
      if (input.recordId) {
        const existingLocalRecord = (
          await listLocalPlatformRecords({
            tenantId: context.tenantId,
            environmentId: context.environmentId,
            objectKey: objectDefinition.key,
          })
        ).find((record) => record.id === input.recordId);
        if (!existingLocalRecord) {
          throw new PlatformNotFoundError("Record not found in the active tenant scope.");
        }
      }

      const record = await upsertLocalPlatformRecord({
        tenantId: context.tenantId,
        environmentId: context.environmentId,
        objectKey: objectDefinition.key,
        recordId: input.recordId,
        data: validation.data,
        actor: context.actor,
      });
      return toPlatformRecord(record);
    },
  );

  await appendAuditEvent({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    actor: context.actor,
    action: input.recordId ? "record.updated" : "record.created",
    resourceType: "record",
    resourceId: saved.id,
    summary: `${input.recordId ? "Updated" : "Created"} ${objectDefinition.label} record.`,
    payload: {
      objectKey: objectDefinition.key,
    },
  });

  return saved;
}

export async function deletePlatformRecord(input: {
  tenantSlug: string;
  objectKey: string;
  recordId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<void> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const runtimeManifest = await getRuntimeManifest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  if (!runtimeManifest) {
    throw new PlatformError("No published runtime is active for this tenant.", 409);
  }

  const objectDefinition = findObjectDefinition(runtimeManifest, input.objectKey);
  assertExists(objectDefinition, "Object definition not found in the active runtime.");

  await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const record = await prisma.platformRecord.findFirst({
        where: {
          id: input.recordId,
          tenantId: context.tenantId,
          environmentId: context.environmentId,
          objectKey: objectDefinition.key,
        },
      });
      if (!record) {
        throw new PlatformNotFoundError("Record not found in the active tenant scope.");
      }

      await prisma.platformRecord.delete({
        where: { id: record.id },
      });
    },
    async () => {
      const existingLocalRecord = (
        await listLocalPlatformRecords({
          tenantId: context.tenantId,
          environmentId: context.environmentId,
          objectKey: objectDefinition.key,
        })
      ).find((record) => record.id === input.recordId);
      if (!existingLocalRecord) {
        throw new PlatformNotFoundError("Record not found in the active tenant scope.");
      }

      await deleteLocalPlatformRecord({
        tenantId: context.tenantId,
        environmentId: context.environmentId,
        objectKey: objectDefinition.key,
        recordId: input.recordId,
      });
    },
  );

  await appendAuditEvent({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    actor: context.actor,
    action: "record.deleted",
    resourceType: "record",
    resourceId: input.recordId,
    summary: `Deleted ${objectDefinition.key} record.`,
    payload: {
      objectKey: objectDefinition.key,
    },
  });
}

export async function listWorkflowRuns(input: {
  tenantSlug: string;
  workflowId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformWorkflowRunRecord[]> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const manifest = await getRuntimeManifest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });

  if (!manifest) {
    throw new PlatformError("No published runtime is active for this tenant.", 409);
  }

  const workflow = manifest.workflows.find((candidate) => candidate.id === input.workflowId || candidate.key === input.workflowId);
  assertExists(workflow, "Workflow not found in the active runtime.");

  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const runs = await prisma.platformWorkflowRun.findMany({
        where: {
          tenantId: context.tenantId,
          environmentId: context.environmentId,
          workflowId: workflow.id,
          workflowKey: workflow.key,
        },
        orderBy: [{ createdAt: "desc" }],
        take: 25,
      });

      return runs.map(toWorkflowRunRecord);
    },
    async () => {
      const runs = await listLocalWorkflowRuns({
        tenantId: context.tenantId,
        environmentId: context.environmentId,
        workflowId: workflow.id,
        workflowKey: workflow.key,
      });
      return runs.map(toWorkflowRunRecord);
    },
  );
}

export async function listAllWorkflowRuns(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformWorkflowRunRecord[]> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });

  return withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const runs = await prisma.platformWorkflowRun.findMany({
        where: {
          tenantId: context.tenantId,
          environmentId: context.environmentId,
        },
        orderBy: [{ updatedAt: "desc" }],
        take: 50,
      });

      return runs.map(toWorkflowRunRecord);
    },
    async () => {
      const runs = await listLocalWorkflowRuns({
        tenantId: context.tenantId,
        environmentId: context.environmentId,
      });
      return runs.map(toWorkflowRunRecord);
    },
  );
}

export async function queueWorkflowRun(input: {
  tenantSlug: string;
  workflowId: string;
  payload?: Record<string, unknown>;
  meta?: {
    parentWorkflowRunId?: string;
    replayedFromRunId?: string;
    parentAgentRunId?: string;
  };
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformWorkflowRunRecord> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const manifest = (await getRuntimeManifest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  })) as PlatformManifest | null;

  if (!manifest) {
    throw new PlatformError("No published runtime is active for this tenant.", 409);
  }

  const workflow = manifest.workflows.find((candidate) => candidate.id === input.workflowId || candidate.key === input.workflowId);
  assertExists(workflow, "Workflow not found in the active runtime.");
  const runInput = {
    ...(input.payload ?? {}),
    ...(input.meta ?? {}),
  };

  const run = await withLocalFallback(
    async () => {
      const prisma = getPrisma();
      const record = await prisma.platformWorkflowRun.create({
        data: {
          tenantId: context.tenantId,
          environmentId: context.environmentId,
          workflowId: workflow.id,
          workflowKey: workflow.key,
          status: "QUEUED",
          input: Object.keys(runInput).length > 0 ? toJsonValue(runInput) : toJsonValue(null),
          logs: toJsonValue([]),
        },
      });

      return toWorkflowRunRecord(record);
    },
    async () => {
      const record = await createLocalWorkflowRun({
        tenantId: context.tenantId,
        environmentId: context.environmentId,
        workflowId: workflow.id,
        workflowKey: workflow.key,
        input: Object.keys(runInput).length > 0 ? runInput : undefined,
      });

      return toWorkflowRunRecord(record);
    },
  );

  await appendAuditEvent({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    actor: context.actor,
    action: "workflow.run.queued",
    resourceType: "workflow_run",
    resourceId: run.id,
    summary: `Queued workflow ${workflow.name}.`,
    payload: {
      workflowId: workflow.id,
      workflowKey: workflow.key,
    },
  });

  await emitPlatformEvent({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
    type: "workflow.run.queued",
    source: "workflow",
    resourceType: "workflow_run",
    resourceId: run.id,
    payload: {
      workflowKey: workflow.key,
      status: run.status,
    },
  });

  await enqueueWorkflowRun(run.id).catch(() => {
    // Polling worker remains the fallback path when Redis/BullMQ is unavailable.
  });

  return run;
}

export async function replayWorkflowRun(input: {
  tenantSlug: string;
  runId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformWorkflowRunRecord> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "BUILDER_ADMIN");

  const original = (await listAllWorkflowRuns({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  })).find((candidate) => candidate.id === input.runId);
  assertExists(original, "Workflow run not found.");

  const replay = await queueWorkflowRun({
    tenantSlug: input.tenantSlug,
    workflowId: original.workflowId,
    payload: {
      ...(original.input ?? {}),
      replayedAt: new Date().toISOString(),
    },
    meta: {
      replayedFromRunId: original.id,
      parentWorkflowRunId: original.parentWorkflowRunId ?? undefined,
    },
    environmentSlug: input.environmentSlug,
    request: input.request,
  });

  await appendAuditEvent({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    actor: context.actor,
    action: "workflow.run.replayed",
    resourceType: "workflow_run",
    resourceId: replay.id,
    summary: `Replayed workflow run ${original.id}.`,
    payload: {
      originalRunId: original.id,
      replayRunId: replay.id,
    },
  });

  return replay;
}

export async function getPublishPreview(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformPublishPreview> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "BUILDER_ADMIN");

  const activeManifest = await loadActiveRuntimeManifest(context);

  return {
    tenantSlug: context.tenant.slug,
    environmentSlug: context.environment.slug,
    activeVersionNumber: context.activeVersion?.versionNumber ?? null,
    nextVersionNumber: Math.max(0, ...context.versions.map((version) => version.versionNumber)) + 1,
    summary: {
      objects: diffPreviewBucket({
        draftItems: context.draftManifest.objects,
        activeItems: activeManifest?.objects ?? [],
        getKey: (item) => item.key,
        getLabel: (item) => item.label,
      }),
      branding: diffPreviewBucket({
        draftItems: [context.draftManifest.branding],
        activeItems: activeManifest ? [activeManifest.branding] : [],
        getKey: () => "tenant-branding",
        getLabel: (item) => item.themeName,
      }),
      pages: diffPreviewBucket({
        draftItems: context.draftManifest.pages,
        activeItems: activeManifest?.pages ?? [],
        getKey: (item) => item.key,
        getLabel: (item) => item.title,
      }),
      layouts: diffPreviewBucket({
        draftItems: context.draftManifest.layouts,
        activeItems: activeManifest?.layouts ?? [],
        getKey: (item) => item.key,
        getLabel: (item) => item.name,
      }),
      menus: diffPreviewBucket({
        draftItems: context.draftManifest.menus,
        activeItems: activeManifest?.menus ?? [],
        getKey: (item) => item.key,
        getLabel: (item) => item.label,
      }),
      workflows: diffPreviewBucket({
        draftItems: context.draftManifest.workflows,
        activeItems: activeManifest?.workflows ?? [],
        getKey: (item) => item.key,
        getLabel: (item) => item.name,
      }),
      agents: diffPreviewBucket({
        draftItems: context.draftManifest.agents,
        activeItems: activeManifest?.agents ?? [],
        getKey: (item) => item.key,
        getLabel: (item) => item.name,
      }),
      modelProviders: diffPreviewBucket({
        draftItems: context.draftManifest.modelProviders,
        activeItems: activeManifest?.modelProviders ?? [],
        getKey: (item) => item.key,
        getLabel: (item) => item.name,
      }),
      notifications: diffPreviewBucket({
        draftItems: context.draftManifest.notifications.rules,
        activeItems: activeManifest?.notifications.rules ?? [],
        getKey: (item) => item.key,
        getLabel: (item) => item.name,
      }),
    },
    pageImpacts: buildPublishPageImpacts({
      draftManifest: context.draftManifest,
      activeManifest,
    }),
    routeImpacts: buildPublishRouteImpacts({
      draftPages: context.draftManifest.pages,
      activePages: activeManifest?.pages ?? [],
    }),
  };
}

export async function previewAgentInvocation(input: {
  tenantSlug: string;
  agentId: string;
  objectKey?: string;
  sampleSize?: number;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformAgentPreview> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "BUILDER_ADMIN");

  const runtimeManifest = await getRuntimeManifest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const manifest = runtimeManifest ?? context.draftManifest;

  const agent = manifest.agents.find((candidate) => candidate.id === input.agentId || candidate.key === input.agentId);
  assertExists(agent, "Agent definition not found in the active runtime.");

  const objectKey = input.objectKey ?? agent.objectKeys[0];
  if (!objectKey) {
    throw new PlatformError("Agent preview requires an object scope.");
  }
  if (agent.objectKeys.length > 0 && !agent.objectKeys.includes(objectKey)) {
    throw new PlatformError("Agent preview object is outside the agent scope.");
  }

  const objectDefinition = findObjectDefinition(manifest, objectKey);
  assertExists(objectDefinition, "Agent preview object is invalid.");
  const records = await listTenantScopedRecords({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: objectDefinition.key,
  });
  const sampleSize = Math.min(Math.max(input.sampleSize ?? 3, 1), 10);
  const invocation = prepareAgentInvocation({
    manifest,
    agentId: agent.id,
    objectKey,
    records: records.slice(0, sampleSize),
  });

  await appendAuditEvent({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    actor: context.actor,
    action: "agent.previewed",
    resourceType: "agent",
    resourceId: agent.id,
    summary: `Prepared masked preview for agent ${agent.name}.`,
    payload: {
      objectKey,
      sampleSize,
      providerKey: invocation.provider.key,
    },
  });

  const recentActivity = (await listAuditEvents({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  }))
    .filter((event) => event.resourceType === "agent" && event.resourceId === agent.id)
    .slice(0, 10);

  return {
    agent: {
      id: invocation.agent.id,
      key: invocation.agent.key,
      name: invocation.agent.name,
      scope: invocation.agent.scope,
      zeroRetentionRequired: invocation.agent.zeroRetentionRequired,
    },
    provider: {
      id: invocation.provider.id,
      key: invocation.provider.key,
      name: invocation.provider.name,
      provider: invocation.provider.provider,
      model: invocation.provider.model,
      supportsZeroRetention: invocation.provider.supportsZeroRetention,
      allowedForSensitiveData: invocation.provider.allowedForSensitiveData,
      status: invocation.provider.status,
    },
    objectKey,
    sampleSize,
    maskedRecords: invocation.inputRecords,
    metadata: {
      masked: invocation.metadata.masked,
      zeroRetentionRequired: invocation.metadata.zeroRetentionRequired,
      allowedToolIds: invocation.agent.allowedToolIds,
      allowedObjectKeys: invocation.agent.objectKeys,
      allowedByPolicy: manifest.securityPolicy.allowedModelProviderKeys.includes(invocation.provider.key),
    },
    recentActivity,
  };
}

function interpolateTemplate(template: string, payload: Record<string, unknown>): string {
  return template.replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, (_match, token: string) => {
    const value = payload[token];
    if (value === undefined || value === null) {
      return "";
    }
    return String(value);
  });
}

function getNotificationMaxAttempts(kind: NotificationCenterDefinition["channels"][number]["kind"]): number {
  switch (kind) {
    case "in_app":
      return 1;
    case "email":
      return 3;
    case "webhook":
    case "slack_style":
      return 4;
    default:
      return 1;
  }
}

function getNotificationRetryDelayMs(
  kind: NotificationCenterDefinition["channels"][number]["kind"],
  attemptNumber: number,
): number | null {
  if (kind === "in_app") {
    return null;
  }

  const schedule = [60_000, 5 * 60_000, 15 * 60_000, 60 * 60_000];
  return schedule[Math.max(0, Math.min(attemptNumber - 1, schedule.length - 1))] ?? schedule[schedule.length - 1]!;
}

function getNextNotificationRetryAt(
  kind: NotificationCenterDefinition["channels"][number]["kind"],
  attemptNumber: number,
  maxAttempts: number,
): string | null {
  if (attemptNumber >= maxAttempts) {
    return null;
  }

  const delayMs = getNotificationRetryDelayMs(kind, attemptNumber);
  if (!delayMs) {
    return null;
  }

  return new Date(Date.now() + delayMs).toISOString();
}

async function getNotificationChannelHealthRecord(input: {
  context: PlatformContext;
  channelKey: string;
}) {
  const existing = (
    await listTenantScopedRecords({
      tenantId: input.context.tenantId,
      environmentId: input.context.environmentId,
      objectKey: getSystemObjectKey("channelHealth"),
    })
  )
    .map(toNotificationChannelHealth)
    .find((candidate) => candidate.channelKey === input.channelKey);

  return existing ?? null;
}

async function recordNotificationAttempt(input: {
  context: PlatformContext;
  deliveryId: string;
  eventId: string;
  channelKey: string;
  status: NotificationAttemptRecord["status"];
  lifecycleStage?: NotificationAttemptRecord["lifecycleStage"];
  provider: string;
  destination?: string;
  responseSummary?: string | null;
  errorMessage?: string | null;
  attemptNumber: number;
}) {
  await upsertTenantScopedRecord({
    tenantId: input.context.tenantId,
    environmentId: input.context.environmentId,
    objectKey: getSystemObjectKey("deliveryAttempt"),
    actor: input.context.actor,
    data: {
      deliveryId: input.deliveryId,
      eventId: input.eventId,
      channelKey: input.channelKey,
      status: input.status,
      provider: input.provider,
      destination: input.destination,
      responseSummary: input.responseSummary,
      errorMessage: input.errorMessage,
      attemptNumber: input.attemptNumber,
      lifecycleStage: input.lifecycleStage,
    },
  });
}

async function updateNotificationChannelHealth(input: {
  context: PlatformContext;
  channelKey: string;
  success?: boolean;
  disabled?: boolean;
}) {
  const existing = await getNotificationChannelHealthRecord({
    context: input.context,
    channelKey: input.channelKey,
  });

  const successCount = (existing?.successCount ?? 0) + (input.success ? 1 : 0);
  const failureCount = (existing?.failureCount ?? 0) + (input.success === false ? 1 : 0);
  const consecutiveFailures =
    input.success === true ? 0 : input.success === false ? (existing?.consecutiveFailures ?? 0) + 1 : existing?.consecutiveFailures ?? 0;
  const totalAttempts = successCount + failureCount;
  const successRate = totalAttempts > 0 ? Number((successCount / totalAttempts).toFixed(4)) : 0;
  const disabledAt =
    typeof input.disabled === "boolean"
      ? input.disabled
        ? new Date().toISOString()
        : null
      : existing?.disabledAt ?? null;
  const status =
    disabledAt != null
      ? "disabled"
      : consecutiveFailures >= 3
        ? "degraded"
        : "healthy";

  await upsertTenantScopedRecord({
    tenantId: input.context.tenantId,
    environmentId: input.context.environmentId,
    objectKey: getSystemObjectKey("channelHealth"),
    recordId: existing?.id,
    actor: input.context.actor,
    data: {
      channelKey: input.channelKey,
      status,
      successCount,
      failureCount,
      consecutiveFailures,
      successRate,
      lastDeliveredAt: input.success ? new Date().toISOString() : existing?.lastDeliveredAt ?? null,
      lastFailedAt: input.success === false ? new Date().toISOString() : existing?.lastFailedAt ?? null,
      disabledAt,
      updatedAt: new Date().toISOString(),
    },
  });
}

export function validateAgentOutputSchema(outputSchema: string | undefined, outputText: string): {
  passed: boolean;
  summary: string;
  parsedOutput?: Record<string, unknown>;
} {
  if (!outputSchema?.trim()) {
    return {
      passed: true,
      summary: "No output schema configured.",
    };
  }

  let parsedSchema: Record<string, unknown>;
  try {
    parsedSchema = JSON.parse(outputSchema) as Record<string, unknown>;
  } catch {
    return {
      passed: false,
      summary: "Output schema is not valid JSON.",
    };
  }

  let parsedOutput: Record<string, unknown>;
  try {
    const candidate = JSON.parse(outputText) as unknown;
    if (!candidate || typeof candidate !== "object" || Array.isArray(candidate)) {
      return {
        passed: false,
        summary: "Model output is not a JSON object.",
      };
    }
    parsedOutput = candidate as Record<string, unknown>;
  } catch {
    return {
      passed: false,
      summary: "Model output is not valid JSON.",
    };
  }

  const requiredKeys = Array.isArray(parsedSchema.required)
    ? parsedSchema.required.map((entry) => String(entry))
    : [];
  const missing = requiredKeys.filter((key) => !(key in parsedOutput));
  if (missing.length > 0) {
    return {
      passed: false,
      summary: `Missing required output keys: ${missing.join(", ")}.`,
      parsedOutput,
    };
  }

  return {
    passed: true,
    summary: "Output satisfied the configured schema contract.",
    parsedOutput,
  };
}

export function buildAgentTrace(input: {
  agent: AgentDefinition;
  provider: ModelProviderDefinition;
  outputValidationPassed?: boolean;
  handoffWorkflowKey?: string;
  policyDecisions: string[];
  stages: AgentTrace["stages"];
}): AgentTrace {
  return {
    stages: input.stages,
    promptBlockSummary: input.agent.promptBlocks.map((block) => ({
      id: block.id,
      label: block.label,
      kind: block.kind,
    })),
    policyDecisions: input.policyDecisions,
    providerKey: input.provider.key,
    providerModel: input.provider.model,
    outputValidationPassed: input.outputValidationPassed,
    handoffWorkflowKey: input.handoffWorkflowKey,
  };
}

export async function emitPlatformEvent(input: {
  tenantSlug: string;
  type: string;
  source: EventEnvelope["source"];
  resourceType: string;
  resourceId: string;
  payload?: Record<string, unknown>;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<EventEnvelope> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });

  const record = await upsertTenantScopedRecord({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("outbox"),
    actor: context.actor,
    data: {
      type: input.type,
      source: input.source,
      resourceType: input.resourceType,
      resourceId: input.resourceId,
      tenantSlug: context.tenant.slug,
      environmentSlug: context.environment.slug,
      emittedAt: new Date().toISOString(),
      enqueuedAt: new Date().toISOString(),
      payload: input.payload ?? {},
    },
  });

  await enqueueOutboxEvent({
    tenantSlug: context.tenant.slug,
    environmentSlug: context.environment.slug,
    eventId: record.id,
  }).catch(() => false);

  return toEventEnvelope(record);
}

async function attemptNotificationDelivery(input: {
  context: PlatformContext;
  manifest: PlatformManifest;
  event: EventEnvelope;
  channel: NotificationCenterDefinition["channels"][number];
  delivery: NotificationDeliveryRecord;
  ruleName: string;
}): Promise<{
  delivery: NotificationDeliveryRecord;
  alertRaised: boolean;
}> {
  const attemptNumber = (input.delivery.attemptCount ?? 0) + 1;
  const maxAttempts = input.delivery.maxAttempts ?? getNotificationMaxAttempts(input.channel.kind);
  const nextRetryAt = getNextNotificationRetryAt(input.channel.kind, attemptNumber, maxAttempts);
  const disabledHealth = await getNotificationChannelHealthRecord({
    context: input.context,
    channelKey: input.channel.key,
  });

  if (disabledHealth?.disabledAt) {
    const saved = await upsertTenantScopedRecord({
      tenantId: input.context.tenantId,
      environmentId: input.context.environmentId,
      objectKey: getSystemObjectKey("delivery"),
      recordId: input.delivery.id,
      actor: input.context.actor,
      data: {
        ...input.delivery,
        status: "exhausted",
        disabledAt: disabledHealth.disabledAt,
        errorMessage: "Notification channel is disabled.",
        attemptCount: input.delivery.attemptCount ?? 0,
        maxAttempts,
        exhaustedAt: new Date().toISOString(),
      },
    });
    return {
      delivery: toNotificationDeliveryRecord(saved),
      alertRaised: false,
    };
  }

  await recordNotificationAttempt({
    context: input.context,
    deliveryId: input.delivery.id,
    eventId: input.event.id,
    channelKey: input.channel.key,
    status: "pending",
    lifecycleStage: "attempt_started",
    provider: input.channel.kind,
    destination: input.delivery.destination,
    attemptNumber,
  });

  try {
    const result = await deliverNotification({
      channel: input.channel,
      severity: input.delivery.severity,
      subject: input.delivery.subject,
      body: input.delivery.body,
      event: input.event,
      supportEmail: input.manifest.appShell.supportEmail,
    });
    const saved = await upsertTenantScopedRecord({
      tenantId: input.context.tenantId,
      environmentId: input.context.environmentId,
      objectKey: getSystemObjectKey("delivery"),
      recordId: input.delivery.id,
      actor: input.context.actor,
      data: {
        ...input.delivery,
        status: "sent",
        provider: result.provider,
        providerResponseSummary: result.responseSummary ?? null,
        destination: result.destination ?? input.delivery.destination ?? input.channel.destination,
        deliveredAt: result.deliveredAt,
        lastAttemptAt: result.deliveredAt,
        attemptCount: attemptNumber,
        maxAttempts,
        nextRetryAt: null,
        exhaustedAt: null,
        errorMessage: null,
      },
    });
    await recordNotificationAttempt({
      context: input.context,
      deliveryId: input.delivery.id,
      eventId: input.event.id,
      channelKey: input.channel.key,
      status: "sent",
      lifecycleStage: "attempt_succeeded",
      provider: result.provider,
      destination: result.destination ?? input.delivery.destination ?? input.channel.destination,
      responseSummary: result.responseSummary ?? null,
      attemptNumber,
    });
    await updateNotificationChannelHealth({
      context: input.context,
      channelKey: input.channel.key,
      success: true,
    });
    return {
      delivery: toNotificationDeliveryRecord(saved),
      alertRaised: false,
    };
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : "Notification delivery failed.";
    const exhausted = !nextRetryAt;
    const saved = await upsertTenantScopedRecord({
      tenantId: input.context.tenantId,
      environmentId: input.context.environmentId,
      objectKey: getSystemObjectKey("delivery"),
      recordId: input.delivery.id,
      actor: input.context.actor,
      data: {
        ...input.delivery,
        status: exhausted ? "exhausted" : "retrying",
        errorMessage,
        provider: input.channel.kind,
        providerResponseSummary: null,
        attemptCount: attemptNumber,
        maxAttempts,
        lastAttemptAt: new Date().toISOString(),
        nextRetryAt,
        exhaustedAt: exhausted ? new Date().toISOString() : null,
      },
    });
    await recordNotificationAttempt({
      context: input.context,
      deliveryId: input.delivery.id,
      eventId: input.event.id,
      channelKey: input.channel.key,
      status: exhausted ? "exhausted" : "failed",
      lifecycleStage: exhausted ? "exhausted" : "attempt_failed",
      provider: input.channel.kind,
      destination: input.delivery.destination ?? input.channel.destination,
      errorMessage,
      attemptNumber,
    });
    await updateNotificationChannelHealth({
      context: input.context,
      channelKey: input.channel.key,
      success: false,
    });

    if (exhausted) {
      await upsertTenantScopedRecord({
        tenantId: input.context.tenantId,
        environmentId: input.context.environmentId,
        objectKey: getSystemObjectKey("alert"),
        actor: input.context.actor,
        data: {
          category: "delivery",
          severity: "warning",
          title: `Delivery exhausted for ${input.ruleName}`,
          summary: errorMessage,
          sourceId: input.delivery.id,
        },
      });
      await upsertTenantScopedRecord({
        tenantId: input.context.tenantId,
        environmentId: input.context.environmentId,
        objectKey: getSystemObjectKey("deadLetter"),
        actor: input.context.actor,
        data: {
          eventId: input.event.id,
          type: "notification.delivery.exhausted",
          reason: errorMessage,
          payload: {
            deliveryId: input.delivery.id,
            channelKey: input.channel.key,
            ruleKey: input.delivery.ruleKey,
            attemptNumber,
          },
        },
      });
    }

    return {
      delivery: toNotificationDeliveryRecord(saved),
      alertRaised: exhausted,
    };
  }
}

async function processDueNotificationRetries(input: {
  context: PlatformContext;
  manifest: PlatformManifest;
  request?: Request | Headers;
}): Promise<{
  deliveries: number;
  alerts: number;
}> {
  const now = Date.now();
  const deliveries = (await listTenantScopedRecords({
    tenantId: input.context.tenantId,
    environmentId: input.context.environmentId,
    objectKey: getSystemObjectKey("delivery"),
  }))
    .map(toNotificationDeliveryRecord)
    .filter(
      (delivery) =>
        delivery.status === "retrying" &&
        typeof delivery.nextRetryAt === "string" &&
        new Date(delivery.nextRetryAt).getTime() <= now,
    );

  let processedDeliveries = 0;
  let raisedAlerts = 0;

  for (const delivery of deliveries) {
    const eventRecord = (
      await listTenantScopedRecords({
        tenantId: input.context.tenantId,
        environmentId: input.context.environmentId,
        objectKey: getSystemObjectKey("outbox"),
      })
    ).find((record) => record.id === delivery.eventId);
    if (!eventRecord) {
      continue;
    }

    const event = toEventEnvelope(eventRecord);
    const channel = input.manifest.notifications.channels.find((candidate) => candidate.key === delivery.channelKey);
    if (!channel || !channel.enabled) {
      continue;
    }

    const result = await attemptNotificationDelivery({
      context: input.context,
      manifest: input.manifest,
      event,
      channel,
      delivery,
      ruleName: delivery.ruleKey,
    });
    processedDeliveries += 1;
    if (result.alertRaised) {
      raisedAlerts += 1;
    }
  }

  return {
    deliveries: processedDeliveries,
    alerts: raisedAlerts,
  };
}

export async function processPendingOutboxEvents(input: {
  tenantSlug: string;
  environmentSlug?: string;
  eventId?: string;
  request?: Request | Headers;
}): Promise<{
  processed: number;
  deliveries: number;
  alerts: number;
}> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });

  const manifest =
    (await getRuntimeManifest({
      tenantSlug: input.tenantSlug,
      environmentSlug: input.environmentSlug,
      request: input.request,
    })) ?? context.draftManifest;

  const outbox = (await listTenantScopedRecords({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("outbox"),
  }))
    .filter((record) => !record.data.processedAt && (!input.eventId || record.id === input.eventId))
    .sort((left, right) => left.createdAt.localeCompare(right.createdAt));

  let processed = 0;
  let deliveries = 0;
  let alerts = 0;

  for (const record of outbox) {
    const event = toEventEnvelope(record);
    const matchingRules = manifest.notifications.rules.filter((rule) => rule.active && rule.eventType === event.type);

    for (const rule of matchingRules) {
      const template = manifest.notifications.templates.find((candidate) => candidate.key === rule.templateKey);
      if (!template) {
        continue;
      }

      const notificationPayload = {
        ...event.payload,
        eventType: event.type,
        resourceId: event.resourceId,
        resourceType: event.resourceType,
      };

      await upsertTenantScopedRecord({
        tenantId: context.tenantId,
        environmentId: context.environmentId,
        objectKey: getSystemObjectKey("ruleMatch"),
        actor: context.actor,
        data: {
          eventId: event.id,
          ruleKey: rule.key,
          templateKey: template.key,
          matchedAt: new Date().toISOString(),
          channelKeys: rule.channelKeys,
          payload: notificationPayload,
        },
      });

      for (const channelKey of rule.channelKeys) {
        const channel = manifest.notifications.channels.find((candidate) => candidate.key === channelKey);
        if (!channel || !channel.enabled) {
          continue;
        }
        const subject = template.subject ? interpolateTemplate(template.subject, notificationPayload) : undefined;
        const body = interpolateTemplate(template.body, notificationPayload);
        const maxAttempts = getNotificationMaxAttempts(channel.kind);
        const pendingDelivery = await upsertTenantScopedRecord({
          tenantId: context.tenantId,
          environmentId: context.environmentId,
          objectKey: getSystemObjectKey("delivery"),
          actor: context.actor,
          data: {
            eventId: event.id,
            ruleKey: rule.key,
            channelKey: channel.key,
            templateKey: template.key,
            status: "pending",
            severity: rule.severity,
            subject,
            body,
            destination: channel.destination,
            provider: channel.kind,
            attemptCount: 0,
            maxAttempts,
            resolvedPayload: notificationPayload,
          },
        });

        const result = await attemptNotificationDelivery({
          context,
          manifest,
          event,
          channel,
          delivery: toNotificationDeliveryRecord(pendingDelivery),
          ruleName: rule.name,
        });
        deliveries += 1;

        if (result.delivery.status === "sent" && (rule.severity === "warning" || rule.severity === "critical")) {
          await upsertTenantScopedRecord({
            tenantId: context.tenantId,
            environmentId: context.environmentId,
            objectKey: getSystemObjectKey("alert"),
            actor: context.actor,
            data: {
              category: event.type.startsWith("agent") ? "budget" : "runtime",
              severity: rule.severity,
              title: subject ?? rule.name,
              summary: body,
              sourceId: event.id,
            },
          });
          alerts += 1;
        }

        if (result.alertRaised) {
          alerts += 1;
        }
      }
    }

    await upsertTenantScopedRecord({
      tenantId: context.tenantId,
      environmentId: context.environmentId,
      objectKey: getSystemObjectKey("outbox"),
      recordId: record.id,
      actor: context.actor,
      data: {
        ...record.data,
        processedAt: new Date().toISOString(),
      },
    });
    processed += 1;
  }

  const retryResults = await processDueNotificationRetries({
    context,
    manifest,
    request: input.request,
  });

  return {
    processed,
    deliveries: deliveries + retryResults.deliveries,
    alerts: alerts + retryResults.alerts,
  };
}

export async function runNotificationMaintenanceCycle(): Promise<{
  scopes: number;
  deliveries: number;
  alerts: number;
}> {
  const scopes = await listNotificationMaintenanceScopes();
  let deliveries = 0;
  let alerts = 0;

  for (const scope of scopes) {
    const result = await processDueNotificationRetries({
      context: scope.context,
      manifest: scope.manifest,
    });
    deliveries += result.deliveries;
    alerts += result.alerts;
  }

  return {
    scopes: scopes.length,
    deliveries,
    alerts,
  };
}

export async function listNotificationDeliveries(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<NotificationDeliveryRecord[]> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const records = await listTenantScopedRecords({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("delivery"),
  });
  return records.map(toNotificationDeliveryRecord);
}

export async function listNotificationAttempts(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<NotificationAttemptRecord[]> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const records = await listTenantScopedRecords({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("deliveryAttempt"),
  });
  return records.map(toNotificationAttemptRecord);
}

export async function listNotificationChannelHealth(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<NotificationChannelHealth[]> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const records = await listTenantScopedRecords({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("channelHealth"),
  });
  return records.map(toNotificationChannelHealth);
}

export async function listNotificationRuleMatches(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<NotificationRuleMatchRecord[]> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const records = await listTenantScopedRecords({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("ruleMatch"),
  });
  return records.map(toNotificationRuleMatchRecord);
}

export async function listPlatformAlerts(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformAlertRecord[]> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const records = await listTenantScopedRecords({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("alert"),
  });
  return records.map(toPlatformAlertRecord);
}

export async function listAgentRuns(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<AgentRunRecord[]> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const records = await listTenantScopedRecords({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("agentRun"),
  });
  return records.map(toAgentRunRecord);
}

export async function listCostLedgerRecords(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<CostLedgerRecord[]> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const records = await listTenantScopedRecords({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("costLedger"),
  });
  return records.map(toCostLedgerRecord);
}

export async function listDeadLetterRecords(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<DeadLetterRecord[]> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const records = await listTenantScopedRecords({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("deadLetter"),
  });
  return records.map(toDeadLetterRecord);
}

export async function listApprovalTasks(input: {
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformApprovalTaskRecord[]> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const records = await listTenantScopedRecords({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("approvalTask"),
  });
  return records.map(toApprovalTaskRecord);
}

export async function getWorkflowRunDetail(input: {
  tenantSlug: string;
  runId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<WorkflowRunDetail> {
  const [runs, costs, alerts, deliveries, approvals, deadLetters, agentRuns] = await Promise.all([
    listAllWorkflowRuns(input),
    listCostLedgerRecords(input),
    listPlatformAlerts(input),
    listNotificationDeliveries(input),
    listApprovalTasks(input),
    listDeadLetterRecords(input),
    listAgentRuns(input),
  ]);
  const run = runs.find((candidate) => candidate.id === input.runId);
  assertExists(run, "Workflow run not found.");
  return {
    run,
    parentRun: run.parentWorkflowRunId ? runs.find((candidate) => candidate.id === run.parentWorkflowRunId) ?? null : null,
    childRuns: runs.filter((candidate) => candidate.parentWorkflowRunId === run.id),
    costs: costs.filter((entry) => entry.referenceId === run.id),
    alerts: alerts.filter((entry) => entry.sourceId === run.id),
    deliveries: deliveries.filter((entry) => entry.eventId === run.id || entry.id === run.id),
    approvals: approvals.filter((entry) => entry.workflowRunId === run.id),
    deadLetters: deadLetters.filter((entry) => entry.eventId === run.id),
    relatedAgentRuns: agentRuns.filter((entry) => entry.parentWorkflowRunId === run.id || entry.input.workflowRunId === run.id),
  };
}

export async function getAgentRunDetail(input: {
  tenantSlug: string;
  runId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<AgentRunDetail> {
  const [runs, costs, alerts, deliveries, approvals, workflowRuns] = await Promise.all([
    listAgentRuns(input),
    listCostLedgerRecords(input),
    listPlatformAlerts(input),
    listNotificationDeliveries(input),
    listApprovalTasks(input),
    listAllWorkflowRuns(input),
  ]);
  const run = runs.find((candidate) => candidate.id === input.runId);
  assertExists(run, "Agent run not found.");
  return {
    run,
    costs: costs.filter((entry) => entry.referenceId === run.id),
    alerts: alerts.filter((entry) => entry.sourceId === run.id),
    deliveries: deliveries.filter((entry) => entry.eventId === run.id || entry.id === run.id),
    approvals: approvals.filter((entry) => entry.agentRunId === run.id),
    parentWorkflowRun:
      workflowRuns.find((candidate) => candidate.id === run.parentWorkflowRunId || candidate.id === run.input?.workflowRunId) ?? null,
    handoffWorkflowRun:
      workflowRuns.find((candidate) => candidate.id === run.handoffWorkflowRunId || candidate.input?.parentAgentRunId === run.id) ?? null,
  };
}

export async function getNotificationDeliveryDetail(input: {
  tenantSlug: string;
  deliveryId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<{
  delivery: NotificationDeliveryRecord;
  attempts: NotificationAttemptRecord[];
  channelHealth: NotificationChannelHealth | null;
}> {
  const [deliveries, attempts, health] = await Promise.all([
    listNotificationDeliveries(input),
    listNotificationAttempts(input),
    listNotificationChannelHealth(input),
  ]);
  const delivery = deliveries.find((candidate) => candidate.id === input.deliveryId);
  assertExists(delivery, "Notification delivery not found.");
  return {
    delivery,
    attempts: attempts.filter((candidate) => candidate.deliveryId === delivery.id),
    channelHealth: health.find((candidate) => candidate.channelKey === delivery.channelKey) ?? null,
  };
}

export async function getPlatformAlertDetail(input: {
  tenantSlug: string;
  alertId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<{
  alert: PlatformAlertRecord;
  relatedDeliveries: NotificationDeliveryRecord[];
  relatedWorkflowRuns: PlatformWorkflowRunRecord[];
  relatedAgentRuns: AgentRunRecord[];
}> {
  const [alerts, deliveries, workflowRuns, agentRuns] = await Promise.all([
    listPlatformAlerts(input),
    listNotificationDeliveries(input),
    listAllWorkflowRuns(input),
    listAgentRuns(input),
  ]);
  const alert = alerts.find((candidate) => candidate.id === input.alertId);
  assertExists(alert, "Alert not found.");
  return {
    alert,
    relatedDeliveries: deliveries.filter((candidate) => candidate.eventId === alert.sourceId || candidate.id === alert.sourceId),
    relatedWorkflowRuns: workflowRuns.filter((candidate) => candidate.id === alert.sourceId),
    relatedAgentRuns: agentRuns.filter((candidate) => candidate.id === alert.sourceId),
  };
}

export async function getApprovalTaskDetail(input: {
  tenantSlug: string;
  taskId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<{
  task: PlatformApprovalTaskRecord;
  workflowRun: PlatformWorkflowRunRecord | null;
  agentRun: AgentRunRecord | null;
}> {
  const [tasks, workflowRuns, agentRuns] = await Promise.all([
    listApprovalTasks(input),
    listAllWorkflowRuns(input),
    listAgentRuns(input),
  ]);
  const task = tasks.find((candidate) => candidate.id === input.taskId);
  assertExists(task, "Approval task not found.");
  return {
    task,
    workflowRun: workflowRuns.find((candidate) => candidate.id === task.workflowRunId) ?? null,
    agentRun: agentRuns.find((candidate) => candidate.id === task.agentRunId) ?? null,
  };
}

export async function getDeadLetterDetail(input: {
  tenantSlug: string;
  recordId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<DeadLetterRecord> {
  const records = await listDeadLetterRecords(input);
  const detail = records.find((candidate) => candidate.id === input.recordId);
  assertExists(detail, "Dead letter not found.");
  return detail;
}

export async function resolveApprovalTask(input: {
  tenantSlug: string;
  taskId: string;
  resolution: "approved" | "rejected";
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformApprovalTaskRecord> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "BUILDER_ADMIN");

  const tasks = await listApprovalTasks({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const task = tasks.find((candidate) => candidate.id === input.taskId);
  assertExists(task, "Approval task not found.");

  const updated = await upsertTenantScopedRecord({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("approvalTask"),
    recordId: task.id,
    actor: context.actor,
    data: {
      ...task,
      status: input.resolution,
      resolvedAt: new Date().toISOString(),
    },
  });

  await appendAuditEvent({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    actor: context.actor,
    action: `${task.taskType === "agent_execution" ? "agent" : "workflow"}.approval.${input.resolution}`,
    resourceType: "approval_task",
    resourceId: task.id,
    summary: `${input.resolution === "approved" ? "Approved" : "Rejected"} ${task.taskType === "agent_execution" ? "agent" : "workflow"} task ${task.nodeLabel}.`,
    payload: {
      workflowRunId: task.workflowRunId,
      workflowKey: task.workflowKey,
      agentRunId: task.agentRunId,
      agentKey: task.agentKey,
    },
  });

  if (task.taskType === "agent_execution" && task.agentRunId) {
    if (input.resolution === "approved") {
      await resumeBlockedAgentRun({
        tenantSlug: input.tenantSlug,
        runId: task.agentRunId,
        environmentSlug: input.environmentSlug,
        request: input.request,
      }).catch(() => undefined);
    } else {
      await upsertTenantScopedRecord({
        tenantId: context.tenantId,
        environmentId: context.environmentId,
        objectKey: getSystemObjectKey("agentRun"),
        recordId: task.agentRunId,
        actor: context.actor,
        data: {
          ...((
            await listAgentRuns({
              tenantSlug: input.tenantSlug,
              environmentSlug: input.environmentSlug,
              request: input.request,
            })
          ).find((candidate) => candidate.id === task.agentRunId) ?? {}),
          status: "failed",
          approvalStatus: "rejected",
          completedAt: new Date().toISOString(),
        },
      }).catch(() => undefined);
    }
    return toApprovalTaskRecord(updated);
  }

  const allRuns = await listAllWorkflowRuns({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const workflowRun = allRuns.find((candidate) => candidate.id === task.workflowRunId);
  if (workflowRun && input.resolution === "approved") {
    await withLocalFallback(
      async () => {
        const prisma = getPrisma();
        await prisma.platformWorkflowRun.update({
          where: { id: workflowRun.id },
          data: {
            status: "QUEUED",
            logs: toJsonValue([
              ...workflowRun.logs,
              {
                level: "info",
                message: `Approval task ${task.nodeLabel} resolved by ${context.actor.email}.`,
                at: new Date().toISOString(),
              },
            ]),
            output: toJsonValue({
              ...(workflowRun.output ?? {}),
              runtimeState: {
                ...((((workflowRun.output as Record<string, unknown> | null) ?? {}).runtimeState as Record<string, unknown> | undefined) ?? {}),
                approvalTaskId: task.id,
              },
            }),
          },
        });
      },
      async () => {
        await updateLocalWorkflowRun({
          runId: workflowRun.id,
          status: "QUEUED",
          output: {
            ...(workflowRun.output ?? {}),
            runtimeState: {
              ...((((workflowRun.output as Record<string, unknown> | null) ?? {}).runtimeState as Record<string, unknown> | undefined) ?? {}),
              approvalTaskId: task.id,
            },
          },
          appendLog: {
            level: "info",
            message: `Approval task ${task.nodeLabel} resolved by ${context.actor.email}.`,
            at: new Date().toISOString(),
          },
        });
      },
    ).catch(() => undefined);

    await enqueueWorkflowRun(workflowRun.id).catch(() => undefined);
  }

  return toApprovalTaskRecord(updated);
}

export async function acknowledgePlatformAlert(input: {
  tenantSlug: string;
  alertId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformAlertRecord> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "BUILDER_ADMIN");

  const alerts = await listPlatformAlerts({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const alert = alerts.find((candidate) => candidate.id === input.alertId);
  assertExists(alert, "Alert not found.");

  const saved = await upsertTenantScopedRecord({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("alert"),
    recordId: alert.id,
    actor: context.actor,
    data: {
      ...alert,
      acknowledgedAt: new Date().toISOString(),
    },
  });

  return toPlatformAlertRecord(saved);
}

export async function retryNotificationDelivery(input: {
  tenantSlug: string;
  deliveryId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<NotificationDeliveryRecord> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "BUILDER_ADMIN");

  const manifest =
    (await getRuntimeManifest({
      tenantSlug: input.tenantSlug,
      environmentSlug: input.environmentSlug,
      request: input.request,
    })) ?? context.draftManifest;

  const deliveries = await listNotificationDeliveries({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const delivery = deliveries.find((candidate) => candidate.id === input.deliveryId);
  assertExists(delivery, "Notification delivery not found.");

  const channel = manifest.notifications.channels.find((candidate) => candidate.key === delivery.channelKey);
  assertExists(channel, "Notification channel is invalid.");
  const event: EventEnvelope = {
    id: delivery.eventId,
    type: "notification.retry",
    tenantSlug: context.tenant.slug,
    environmentSlug: context.environment.slug,
    emittedAt: new Date().toISOString(),
    source: "system",
    resourceType: "notification_delivery",
    resourceId: delivery.id,
    payload: delivery.resolvedPayload ?? {},
  };
  const result = await attemptNotificationDelivery({
    context,
    manifest,
    event,
    channel,
    delivery,
    ruleName: delivery.ruleKey,
  });
  return result.delivery;
}

export async function testNotificationChannel(input: {
  tenantSlug: string;
  channelKey: string;
  body: string;
  subject?: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<NotificationDeliveryRecord> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const manifest =
    (await getRuntimeManifest({
      tenantSlug: input.tenantSlug,
      environmentSlug: input.environmentSlug,
      request: input.request,
    })) ?? context.draftManifest;
  const channel = manifest.notifications.channels.find((candidate) => candidate.key === input.channelKey);
  assertExists(channel, "Notification channel not found.");
  const delivery = await upsertTenantScopedRecord({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("delivery"),
    actor: context.actor,
    data: {
      eventId: nextId("test-event"),
      ruleKey: "channel_test",
      channelKey: channel.key,
      templateKey: "channel_test",
      status: "pending",
      severity: "info",
      subject: input.subject,
      body: input.body,
      destination: channel.destination,
      attemptCount: 0,
      maxAttempts: getNotificationMaxAttempts(channel.kind),
      resolvedPayload: {
        test: true,
      },
    },
  });
  return retryNotificationDelivery({
    tenantSlug: input.tenantSlug,
    deliveryId: delivery.id,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
}

export async function toggleNotificationChannel(input: {
  tenantSlug: string;
  channelKey: string;
  enabled: boolean;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<NotificationChannelHealth> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "BUILDER_ADMIN");

  await updateNotificationChannelHealth({
    context,
    channelKey: input.channelKey,
    disabled: !input.enabled,
  });

  const health = await getNotificationChannelHealthRecord({
    context,
    channelKey: input.channelKey,
  });
  assertExists(health, "Notification channel health could not be updated.");
  return health;
}

async function resolveAgentExecutionManifest(input: {
  context: PlatformContext;
  runMode: AgentRunRecord["runMode"];
  tenantSlug: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<PlatformManifest> {
  if (input.runMode === "runtime") {
    const runtimeManifest = await getRuntimeManifest({
      tenantSlug: input.tenantSlug,
      environmentSlug: input.environmentSlug,
      request: input.request,
    });
    assertExists(runtimeManifest, "No published runtime is active for agent execution.");
    return runtimeManifest;
  }

  return input.context.draftManifest;
}

export async function resumeBlockedAgentRun(input: {
  tenantSlug: string;
  runId: string;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<AgentRunRecord> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "BUILDER_ADMIN");

  const runs = await listAgentRuns({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const run = runs.find((candidate) => candidate.id === input.runId);
  assertExists(run, "Agent run not found.");
  if (run.status !== "blocked") {
    throw new PlatformError("Only blocked agent runs can be resumed.", 409);
  }

  const approvalTask = run.approvalTaskId
    ? (
        await listApprovalTasks({
          tenantSlug: input.tenantSlug,
          environmentSlug: input.environmentSlug,
          request: input.request,
        })
      ).find((candidate) => candidate.id === run.approvalTaskId)
    : null;

  if (approvalTask?.status === "pending") {
    throw new PlatformError("Agent run is still pending approval.", 409);
  }
  if (approvalTask?.status === "rejected") {
    const rejected = await upsertTenantScopedRecord({
      tenantId: context.tenantId,
      environmentId: context.environmentId,
      objectKey: getSystemObjectKey("agentRun"),
      recordId: run.id,
      actor: context.actor,
      data: {
        ...run,
        status: "failed",
        approvalStatus: "rejected",
        completedAt: new Date().toISOString(),
      },
    });
    return toAgentRunRecord(rejected);
  }

  const manifest = await resolveAgentExecutionManifest({
    context,
    runMode: run.runMode,
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  const agent = manifest.agents.find((candidate) => candidate.id === run.agentId || candidate.key === run.agentKey);
  assertExists(agent, "Agent definition not found.");
  const provider = manifest.modelProviders.find((candidate) => candidate.id === agent.modelProviderId || candidate.key === agent.modelProviderId);
  assertExists(provider, "Model provider not found.");

  if (provider.status !== "active") {
    throw new PlatformError("Model provider is disabled.", 409);
  }
  if (!manifest.securityPolicy.allowedModelProviderKeys.includes(provider.key)) {
    throw new PlatformForbiddenError("Model provider is not allowed by tenant security policy.");
  }
  if (agent.zeroRetentionRequired && !provider.supportsZeroRetention) {
    throw new PlatformForbiddenError("Selected model provider does not support zero retention.");
  }

  const prompt = typeof run.input.prompt === "string" ? run.input.prompt : agent.name;
  const maskedRecords = Array.isArray(run.input.maskedRecords)
    ? (run.input.maskedRecords as Array<Record<string, unknown>>)
    : [];
  const policyDecisions = [
    `provider:${provider.key}`,
    `zeroRetention:${String(agent.zeroRetentionRequired)}`,
    `approval:${approvalTask ? approvalTask.status : "not_required"}`,
  ];
  const traceStages: AgentTrace["stages"] = [
    ...(run.trace?.stages ?? []).filter((stage) => stage.stage !== "audit"),
  ];

  let execution;
  try {
    execution = await executeAgentWithProvider({
      agent,
      provider,
      prompt,
      maskedRecords,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Agent execution failed.";
    const failed = await upsertTenantScopedRecord({
      tenantId: context.tenantId,
      environmentId: context.environmentId,
      objectKey: getSystemObjectKey("agentRun"),
      recordId: run.id,
      actor: context.actor,
      data: {
        ...run,
        status: "failed",
        approvalStatus: approvalTask ? "approved" : "not_required",
        logs: [
          ...run.logs,
          {
            level: "error",
            message,
            at: new Date().toISOString(),
          },
        ],
        completedAt: new Date().toISOString(),
      },
    });
    return toAgentRunRecord(failed);
  }

  const schemaValidation = validateAgentOutputSchema(agent.outputSchema, execution.outputText);
  let handoffWorkflowRunId: string | null = null;
  if (schemaValidation.passed && agent.handoffWorkflowKeys.length > 0) {
    const runtimeManifest =
      run.runMode === "runtime"
        ? manifest
        : await getRuntimeManifest({
            tenantSlug: input.tenantSlug,
            environmentSlug: input.environmentSlug,
            request: input.request,
          });
    const handoffWorkflow = runtimeManifest?.workflows.find((candidate) => agent.handoffWorkflowKeys.includes(candidate.key));
    if (handoffWorkflow) {
      const handoffRun = await queueWorkflowRun({
        tenantSlug: input.tenantSlug,
        workflowId: handoffWorkflow.id,
        environmentSlug: input.environmentSlug,
        request: input.request,
        payload: {
          parentAgentKey: agent.key,
          parentAgentRunId: run.id,
          outputSummary: execution.outputText,
        },
        meta: {
          parentAgentRunId: run.id,
        },
      }).catch(() => null);
      handoffWorkflowRunId = handoffRun?.id ?? null;
    }
  }

  const costUsd = summarizeAgentRunCost(provider.model, execution.tokensIn, execution.tokensOut);
  traceStages.push(
    {
      stage: "provider_execution",
      status: "succeeded",
      summary: `Provider responded using ${provider.model}.`,
      at: new Date().toISOString(),
      meta: {
        tokensIn: execution.tokensIn,
        tokensOut: execution.tokensOut,
      },
    },
    {
      stage: "output_validation",
      status: schemaValidation.passed ? "succeeded" : "failed",
      summary: schemaValidation.summary,
      at: new Date().toISOString(),
    },
    {
      stage: "cost_evaluation",
      status: agent.costBudgetUsd != null && costUsd >= agent.costBudgetUsd ? "blocked" : "succeeded",
      summary:
        agent.costBudgetUsd != null && costUsd >= agent.costBudgetUsd
          ? `Budget threshold breached at ${costUsd.toFixed(4)} USD.`
          : `Cost ${costUsd.toFixed(4)} USD within budget.`,
      at: new Date().toISOString(),
    },
    {
      stage: "audit",
      status: "succeeded",
      summary: "Persisted resumed agent run.",
      at: new Date().toISOString(),
    },
  );

  const saved = await upsertTenantScopedRecord({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("agentRun"),
    recordId: run.id,
    actor: context.actor,
    data: {
      ...run,
      status: schemaValidation.passed ? "succeeded" : "failed",
      approvalStatus: approvalTask ? "approved" : "not_required",
      output: {
        summary: execution.outputText,
        provider: provider.model,
      },
      logs: [
        ...run.logs,
        {
          level: schemaValidation.passed ? "info" : "error",
          message: schemaValidation.summary,
          at: new Date().toISOString(),
        },
      ],
      modelProviderKey: provider.key,
      costUsd,
      tokensIn: execution.tokensIn,
      tokensOut: execution.tokensOut,
      trace: buildAgentTrace({
        agent,
        provider,
        policyDecisions,
        outputValidationPassed: schemaValidation.passed,
        handoffWorkflowKey: agent.handoffWorkflowKeys[0],
        stages: traceStages,
      }),
      handoffWorkflowRunId,
      outputValidationPassed: schemaValidation.passed,
      schemaValidation: {
        passed: schemaValidation.passed,
        summary: schemaValidation.summary,
      },
      completedAt: new Date().toISOString(),
    },
  });

  await upsertTenantScopedRecord({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("costLedger"),
    actor: context.actor,
    data: {
      category: "agent_run",
      referenceId: run.id,
      providerKey: provider.key,
      amountUsd: costUsd,
      tokensIn: execution.tokensIn,
      tokensOut: execution.tokensOut,
      summary: `Resumed agent run cost for ${agent.key}.`,
    },
  });

  return toAgentRunRecord(saved);
}

export async function simulateAgentRun(input: {
  tenantSlug: string;
  agentId: string;
  prompt?: string;
  objectKey?: string;
  sampleSize?: number;
  environmentSlug?: string;
  request?: Request | Headers;
}): Promise<{
  preview: PlatformAgentPreview;
  run: AgentRunRecord;
  cost: CostLedgerRecord;
}> {
  const context = await getPlatformContextFromRequest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  assertRole(context.actor, "BUILDER_ADMIN");

  const preview = await previewAgentInvocation({
    tenantSlug: input.tenantSlug,
    agentId: input.agentId,
    objectKey: input.objectKey,
    sampleSize: input.sampleSize,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });

  const providerKey = preview.provider.key;
  const prompt = input.prompt?.trim() || preview.agent.name;
  const agent =
    context.draftManifest.agents.find((candidate) => candidate.id === preview.agent.id || candidate.key === preview.agent.key) ??
    context.draftManifest.agents[0];
  const provider =
    context.draftManifest.modelProviders.find((candidate) => candidate.id === providerKey || candidate.key === providerKey) ??
    context.draftManifest.modelProviders[0];
  assertExists(agent, "Agent definition not found.");
  assertExists(provider, "Model provider not found.");

  const traceStages: AgentTrace["stages"] = [];
  const policyDecisions: string[] = [
    `provider:${provider.key}`,
    `zeroRetention:${String(agent.zeroRetentionRequired)}`,
    `masked:${String(preview.metadata.masked)}`,
  ];
  const logs: Array<Record<string, unknown>> = [
    {
      level: "info",
      message: `Prepared ${preview.sampleSize} masked records for ${preview.agent.name}.`,
      at: new Date().toISOString(),
    },
  ];
  traceStages.push({
    stage: "policy_preflight",
    status: "succeeded",
    summary: `Preflight passed for ${provider.name}.`,
    at: new Date().toISOString(),
    meta: {
      providerKey,
      zeroRetentionRequired: agent.zeroRetentionRequired,
      allowedByPolicy: preview.metadata.allowedByPolicy,
    },
  });
  traceStages.push({
    stage: "prompt_assembly",
    status: "succeeded",
    summary: `Assembled ${agent.promptBlocks.length + 1} prompt blocks.`,
    at: new Date().toISOString(),
    meta: {
      promptBlockCount: agent.promptBlocks.length,
    },
  });

  let outputSummary = `Simulated ${preview.agent.name} against ${preview.sampleSize} masked records.`;
  let estimatedTokensIn = Math.max(120, prompt.length * 4);
  let estimatedTokensOut = 220 + preview.sampleSize * 48;
  let estimatedCostUsd = Number(((estimatedTokensIn + estimatedTokensOut) / 100000).toFixed(4));
  let runStatus: AgentRunRecord["status"] = "succeeded";
  let outputValidationPassed: boolean | null = agent.outputSchema ? false : true;
  let outputValidationSummary = agent.outputSchema ? "Output schema pending validation." : "No output schema configured.";
  let approvalTaskId: string | null = null;
  let handoffWorkflowRunId: string | null = null;

  if (agent.approvalPolicy?.required) {
    runStatus = "blocked";
    const blockedRun = await upsertTenantScopedRecord({
      tenantId: context.tenantId,
      environmentId: context.environmentId,
      objectKey: getSystemObjectKey("agentRun"),
      actor: context.actor,
      data: {
        agentId: preview.agent.id,
        agentKey: preview.agent.key,
        status: runStatus,
        runMode: "simulation",
        approvalStatus: "pending",
        input: {
          prompt,
          objectKey: preview.objectKey,
          sampleSize: preview.sampleSize,
          maskedRecords: preview.maskedRecords,
        },
        output: null,
        logs: [
          ...logs,
          {
            level: "warning",
            message: "Execution blocked pending approval.",
            at: new Date().toISOString(),
          },
        ],
        modelProviderKey: providerKey,
        costUsd: 0,
        tokensIn: 0,
        tokensOut: 0,
        trace: buildAgentTrace({
          agent,
          provider,
          policyDecisions: [...policyDecisions, "approval:required"],
          stages: [
            ...traceStages,
            {
              stage: "provider_execution",
              status: "blocked",
              summary: "Execution blocked pending human approval.",
              at: new Date().toISOString(),
            },
          ],
        }),
        schemaValidation: null,
      },
    });
    const approvalTask = await upsertTenantScopedRecord({
      tenantId: context.tenantId,
      environmentId: context.environmentId,
      objectKey: getSystemObjectKey("approvalTask"),
      actor: context.actor,
      data: {
        workflowRunId: "",
        workflowKey: "",
        nodeId: "agent-execution",
        nodeLabel: agent.name,
        taskType: "agent_execution",
        agentRunId: blockedRun.id,
        agentKey: agent.key,
        approverRole: agent.approvalPolicy.approverRole ?? "BUILDER_ADMIN",
        status: "pending",
        instructions: agent.approvalPolicy.notes ?? "Review agent execution before provider invocation.",
      },
    });
    approvalTaskId = approvalTask.id;
    const updatedBlockedRun = await upsertTenantScopedRecord({
      tenantId: context.tenantId,
      environmentId: context.environmentId,
      objectKey: getSystemObjectKey("agentRun"),
      recordId: blockedRun.id,
      actor: context.actor,
      data: {
        ...blockedRun.data,
        approvalTaskId,
        approvalStatus: "pending",
      },
    });
    return {
      preview,
      run: toAgentRunRecord(updatedBlockedRun),
      cost: {
        id: nextId("cost-preview"),
        category: "agent_run",
        referenceId: updatedBlockedRun.id,
        providerKey,
        amountUsd: 0,
        tokensIn: 0,
        tokensOut: 0,
        createdAt: new Date().toISOString(),
        summary: `Blocked run for ${preview.agent.name}.`,
      },
    };
  }

  try {
    const execution = await executeAgentWithProvider({
      agent,
      provider,
      prompt,
      maskedRecords: preview.maskedRecords,
    });
    outputSummary = execution.outputText;
    estimatedTokensIn = execution.tokensIn;
    estimatedTokensOut = execution.tokensOut;
    estimatedCostUsd = summarizeAgentRunCost(provider.model, execution.tokensIn, execution.tokensOut);
    const outputValidation = validateAgentOutputSchema(agent.outputSchema, execution.outputText);
    outputValidationPassed = outputValidation.passed;
    outputValidationSummary = outputValidation.summary;
    if (!outputValidation.passed) {
      runStatus = "failed";
      outputSummary = execution.outputText;
      logs.push({
        level: "error",
        message: outputValidation.summary,
        at: new Date().toISOString(),
      });
    }
    logs.push({
      level: "info",
      message: `Completed provider-backed run using ${provider.model}.`,
      at: new Date().toISOString(),
    });
    traceStages.push({
      stage: "provider_execution",
      status: "succeeded",
      summary: `Provider responded using ${provider.model}.`,
      at: new Date().toISOString(),
      meta: {
        tokensIn: execution.tokensIn,
        tokensOut: execution.tokensOut,
      },
    });
    traceStages.push({
      stage: "output_validation",
      status: outputValidation.passed ? "succeeded" : "failed",
      summary: outputValidation.summary,
      at: new Date().toISOString(),
    });
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : "Agent execution failed.";
    if (!getEnv().PLATFORM_LOCAL_DEV_MODE) {
      runStatus = "failed";
      logs.push({
        level: "error",
        message: errorMessage,
        at: new Date().toISOString(),
      });
      traceStages.push({
        stage: "provider_execution",
        status: "failed",
        summary: errorMessage,
        at: new Date().toISOString(),
      });
    } else {
      logs.push({
        level: "warning",
        message: `${errorMessage} Falling back to local simulation.`,
        at: new Date().toISOString(),
      });
      traceStages.push({
        stage: "provider_execution",
        status: "failed",
        summary: errorMessage,
        at: new Date().toISOString(),
      });
    }
  }

  if (agent.costBudgetUsd) {
    traceStages.push({
      stage: "cost_evaluation",
      status: estimatedCostUsd >= agent.costBudgetUsd ? "blocked" : "succeeded",
      summary:
        estimatedCostUsd >= agent.costBudgetUsd
          ? `Budget threshold breached at ${estimatedCostUsd.toFixed(4)} USD.`
          : `Cost ${estimatedCostUsd.toFixed(4)} USD within budget.`,
      at: new Date().toISOString(),
    });
  }

  if (runStatus === "succeeded" && agent.handoffWorkflowKeys.length > 0) {
    const activeRuntime = await getRuntimeManifest({
      tenantSlug: input.tenantSlug,
      environmentSlug: input.environmentSlug,
      request: input.request,
    });
    const handoffWorkflow = activeRuntime?.workflows.find((workflow) => agent.handoffWorkflowKeys.includes(workflow.key));
    if (handoffWorkflow) {
      const handoffRun = await queueWorkflowRun({
        tenantSlug: input.tenantSlug,
        workflowId: handoffWorkflow.id,
        payload: {
          parentAgentKey: agent.key,
          prompt,
          outputSummary,
        },
        environmentSlug: input.environmentSlug,
        request: input.request,
      }).catch(() => null);
      handoffWorkflowRunId = handoffRun?.id ?? null;
      traceStages.push({
        stage: "workflow_handoff",
        status: handoffRun ? "succeeded" : "failed",
        summary: handoffRun ? `Queued handoff to ${handoffWorkflow.key}.` : `Unable to queue handoff to ${handoffWorkflow.key}.`,
        at: new Date().toISOString(),
      });
    }
  }

  traceStages.push({
    stage: "audit",
    status: "succeeded",
    summary: "Persisted run trace and audit activity.",
    at: new Date().toISOString(),
  });

  const runRecord = await upsertTenantScopedRecord({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("agentRun"),
    actor: context.actor,
    data: {
      agentId: preview.agent.id,
      agentKey: preview.agent.key,
      status: runStatus,
      runMode: "simulation",
      approvalStatus: approvalTaskId ? "pending" : "not_required",
      input: {
        prompt,
        objectKey: preview.objectKey,
        sampleSize: preview.sampleSize,
      },
      output: {
        summary: outputSummary,
        provider: preview.provider.model,
      },
      logs: [
        ...logs,
        {
          level: runStatus === "succeeded" ? "info" : "error",
          message: runStatus === "succeeded" ? "Simulation completed in Control Tower." : "Simulation finished with a provider failure.",
          at: new Date().toISOString(),
        },
      ],
      modelProviderKey: providerKey,
      costUsd: estimatedCostUsd,
      tokensIn: estimatedTokensIn,
      tokensOut: estimatedTokensOut,
      trace: buildAgentTrace({
        agent,
        provider,
        policyDecisions,
        outputValidationPassed: outputValidationPassed ?? undefined,
        handoffWorkflowKey: agent.handoffWorkflowKeys[0],
        stages: traceStages,
      }),
      approvalTaskId,
      handoffWorkflowRunId,
      outputValidationPassed,
      schemaValidation:
        outputValidationPassed == null
          ? null
          : {
              passed: outputValidationPassed,
              summary: outputValidationSummary,
            },
      completedAt: new Date().toISOString(),
    },
  });

  const costRecord = await upsertTenantScopedRecord({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    objectKey: getSystemObjectKey("costLedger"),
    actor: context.actor,
    data: {
      category: "agent_run",
      referenceId: runRecord.id,
      providerKey,
      amountUsd: estimatedCostUsd,
      tokensIn: estimatedTokensIn,
      tokensOut: estimatedTokensOut,
      summary: `Simulation cost for ${preview.agent.name}.`,
    },
  });

  if (agent.costBudgetUsd && estimatedCostUsd >= agent.costBudgetUsd) {
    await emitPlatformEvent({
      tenantSlug: input.tenantSlug,
      environmentSlug: input.environmentSlug,
      request: input.request,
      type: "agent.budget.threshold",
      source: "agent",
      resourceType: "agent_run",
      resourceId: agent.id,
      payload: {
        agentKey: agent.key,
        costUsd: estimatedCostUsd,
        budgetUsd: agent.costBudgetUsd,
      },
    });
  } else {
    await emitPlatformEvent({
      tenantSlug: input.tenantSlug,
      environmentSlug: input.environmentSlug,
      request: input.request,
      type: runStatus === "succeeded" ? "agent.run.simulated" : "agent.run.failed",
      source: "agent",
      resourceType: "agent_run",
      resourceId: preview.agent.id,
      payload: {
        agentKey: preview.agent.key,
        costUsd: estimatedCostUsd,
      },
    });
  }

  await appendAuditEvent({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
    actor: context.actor,
    action: "agent.simulated",
    resourceType: "agent",
    resourceId: preview.agent.id,
    summary: `Ran Control Tower simulation for ${preview.agent.name}.`,
    payload: {
      providerKey,
      costUsd: estimatedCostUsd,
      objectKey: preview.objectKey,
      status: runStatus,
    },
  });

  return {
    preview,
    run: toAgentRunRecord(runRecord),
    cost: toCostLedgerRecord(costRecord),
  };
}
