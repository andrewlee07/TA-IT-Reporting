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
  listLocalInvites,
  listLocalPlatformRecords,
  listLocalTenantMembershipSummariesForUser,
  listLocalWorkflowRuns,
  listLocalVersions,
  setLocalActiveVersion,
  upsertLocalDraft,
  upsertLocalPlatformRecord,
} from "@/lib/platform/local-store";
import { ensureManifestConsistency, findObjectDefinition, touchManifest } from "@/lib/platform/manifest";
import { buildPublishedArtifacts } from "@/lib/platform/publish";
import { enqueueWorkflowRun } from "@/lib/platform/execution-bus";
import {
  PlatformError,
  PlatformForbiddenError,
  PlatformNotFoundError,
} from "@/lib/platform/errors";
import { assertRole } from "@/lib/platform/rbac";
import { validateRecordInput } from "@/lib/platform/records";
import { createDefaultBranding } from "@/lib/platform/theme";
import { getObjectStorage } from "@/lib/storage";
import type {
  AgentDefinition,
  FieldDefinition,
  FormDefinition,
  FormFieldDefinition,
  FormStepDefinition,
  LayoutDefinition,
  MenuItemDefinition,
  ModelProviderDefinition,
  ObjectDefinition,
  PageDefinition,
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
  PlatformSessionSummary,
  PlatformTenantSummary,
  SecurityPolicyDefinition,
  TenantBrandingDefinition,
  WorkflowDefinition,
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
  return {
    id: record.id,
    workflowId: record.workflowId,
    workflowKey: record.workflowKey,
    status: record.status as PlatformWorkflowRunRecord["status"],
    input: (record.input as Record<string, unknown> | null) ?? null,
    output: (record.output as Record<string, unknown> | null) ?? null,
    logs: (record.logs as Array<Record<string, unknown>> | null) ?? [],
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
        icon: input.menu.icon ?? "dot",
        pageKey: page.key,
        order: input.menu.order ?? manifest.menus.length,
        group: normalizeOptionalText(input.menu.group) ?? "Workspace",
      };

      manifest.menus = upsertById(manifest.menus, menu);
      return {
        manifest,
        result: menu,
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

      const agent: AgentDefinition = {
        id: agentId,
        key,
        name,
        description: normalizeOptionalText(input.agent.description),
        scope: input.agent.scope,
        modelProviderId: provider.id,
        prompt,
        allowedToolIds,
        objectKeys,
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

  return toFormSubmissionRecord(saved);
}

export async function runWorkflowTest(input: {
  tenantSlug: string;
  workflowId: string;
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
      payload: input.payload ?? null,
    },
  });

  return {
    id: nextId("wf-test"),
    workflowId: workflow.id,
    workflowKey: workflow.key,
    status: "SUCCEEDED",
    input: input.payload ?? null,
    output: {
      completedNodes: workflow.nodes.length,
      mode: "draft-test",
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

export async function queueWorkflowRun(input: {
  tenantSlug: string;
  workflowId: string;
  payload?: Record<string, unknown>;
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
          input: input.payload ? toJsonValue(input.payload) : toJsonValue(null),
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
        input: input.payload,
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

  await enqueueWorkflowRun(run.id).catch(() => {
    // Polling worker remains the fallback path when Redis/BullMQ is unavailable.
  });

  return run;
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

  const manifest = await getRuntimeManifest({
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    request: input.request,
  });
  if (!manifest) {
    throw new PlatformError("No published runtime is active for this tenant.", 409);
  }

  const agent = manifest.agents.find((candidate) => candidate.id === input.agentId || candidate.key === input.agentId);
  assertExists(agent, "Agent definition not found in the active runtime.");

  const objectKey = input.objectKey ?? agent.objectKeys[0];
  if (!objectKey) {
    throw new PlatformError("Agent preview requires an object scope.");
  }
  if (agent.objectKeys.length > 0 && !agent.objectKeys.includes(objectKey)) {
    throw new PlatformError("Agent preview object is outside the agent scope.");
  }

  const records = await listPlatformRecords({
    tenantSlug: input.tenantSlug,
    objectKey,
    environmentSlug: input.environmentSlug,
    request: input.request,
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
