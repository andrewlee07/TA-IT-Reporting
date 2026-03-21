import { promises as fs } from "node:fs";
import path from "node:path";

import { nanoid } from "nanoid";

import { getEnv } from "@/lib/env";
import type {
  PlatformActor,
  PlatformAuditEventRecord,
  PlatformEnvironmentSummary,
  PlatformInviteRecord,
  PlatformManifest,
  PlatformPublishedVersionRecord,
  PlatformRecord,
  PlatformRole,
  PlatformTenantMembershipSummary,
  PlatformTenantSummary,
  PlatformVersionStatus,
  PlatformWorkflowRunRecord,
  WorkflowRunStatus,
} from "@/lib/platform/types";

interface LocalTenantRecord extends PlatformTenantSummary {
  createdAt: string;
  updatedAt: string;
}

interface LocalEnvironmentRecord extends PlatformEnvironmentSummary {
  tenantId: string;
  createdAt: string;
  updatedAt: string;
}

interface LocalUserRecord {
  id: string;
  email: string;
  displayName: string;
  createdAt: string;
  updatedAt: string;
}

interface LocalMembershipRecord {
  id: string;
  tenantId: string;
  userId: string;
  role: PlatformRole;
  createdAt: string;
  updatedAt: string;
}

interface LocalInviteRecord extends PlatformInviteRecord {
  acceptedByUserId?: string | null;
}

interface LocalDraftRecord {
  id: string;
  tenantId: string;
  environmentId: string;
  manifest: PlatformManifest;
  createdAt: string;
  updatedAt: string;
}

interface LocalVersionRecord extends PlatformPublishedVersionRecord {
  tenantId: string;
  environmentId: string;
  manifest: PlatformManifest;
}

interface LocalAuditRecord extends PlatformAuditEventRecord {
  tenantId: string;
  environmentId: string | null;
}

interface LocalPlatformRecord extends PlatformRecord {
  tenantId: string;
  environmentId: string;
  createdByEmail?: string | null;
  updatedByEmail?: string | null;
}

interface LocalWorkflowRunRecord extends PlatformWorkflowRunRecord {
  tenantId: string;
  environmentId: string;
}

function getRootDir(): string {
  return path.resolve(process.cwd(), getEnv().LOCAL_STORAGE_DIR, "platform-store");
}

function getPath(fileName: string): string {
  return path.join(getRootDir(), fileName);
}

async function ensureDir(): Promise<void> {
  await fs.mkdir(getRootDir(), { recursive: true });
}

async function readJsonFile<T>(filePath: string, fallback: T): Promise<T> {
  try {
    const raw = await fs.readFile(filePath, "utf8");
    return JSON.parse(raw) as T;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") {
      return fallback;
    }

    throw error;
  }
}

async function writeJsonFile(filePath: string, value: unknown): Promise<void> {
  await ensureDir();
  await fs.writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

async function listTenants(): Promise<LocalTenantRecord[]> {
  return readJsonFile(getPath("tenants.json"), []);
}

async function listEnvironments(): Promise<LocalEnvironmentRecord[]> {
  return readJsonFile(getPath("environments.json"), []);
}

async function listUsers(): Promise<LocalUserRecord[]> {
  return readJsonFile(getPath("users.json"), []);
}

async function listMemberships(): Promise<LocalMembershipRecord[]> {
  return readJsonFile(getPath("memberships.json"), []);
}

async function listDrafts(): Promise<LocalDraftRecord[]> {
  return readJsonFile(getPath("drafts.json"), []);
}

async function listVersions(): Promise<LocalVersionRecord[]> {
  return readJsonFile(getPath("versions.json"), []);
}

async function listAudit(): Promise<LocalAuditRecord[]> {
  return readJsonFile(getPath("audit.json"), []);
}

async function listRecords(): Promise<LocalPlatformRecord[]> {
  return readJsonFile(getPath("records.json"), []);
}

async function listWorkflowRuns(): Promise<LocalWorkflowRunRecord[]> {
  return readJsonFile(getPath("workflow-runs.json"), []);
}

async function listInvites(): Promise<LocalInviteRecord[]> {
  return readJsonFile(getPath("invites.json"), []);
}

function sortByUpdatedDesc<T extends { updatedAt?: string; createdAt?: string }>(items: T[]): T[] {
  return [...items].sort((left, right) => (right.updatedAt ?? right.createdAt ?? "").localeCompare(left.updatedAt ?? left.createdAt ?? ""));
}

export async function ensureLocalActorMembership(input: {
  actor: PlatformActor;
  tenantId: string;
}): Promise<void> {
  const users = await listUsers();
  const existingUser = users.find((user) => user.email === input.actor.email);
  const now = new Date().toISOString();

  const userId = existingUser?.id ?? nanoid();
  if (!existingUser) {
    await writeJsonFile(getPath("users.json"), [
      {
        id: userId,
        email: input.actor.email,
        displayName: input.actor.name,
        createdAt: now,
        updatedAt: now,
      },
      ...users,
    ]);
  }

  const memberships = await listMemberships();
  const existingMembership = memberships.find((membership) => membership.userId === userId && membership.tenantId === input.tenantId);

  if (existingMembership) {
    return;
  }

  await writeJsonFile(getPath("memberships.json"), [
    {
      id: nanoid(),
      tenantId: input.tenantId,
      userId,
      role: input.actor.role,
      createdAt: now,
      updatedAt: now,
    },
    ...memberships,
  ]);
}

export async function getLocalUserByEmail(email: string): Promise<LocalUserRecord | null> {
  const users = await listUsers();
  return users.find((user) => user.email === email) ?? null;
}

export async function createLocalUser(input: {
  email: string;
  displayName: string;
}): Promise<LocalUserRecord> {
  const users = await listUsers();
  const now = new Date().toISOString();
  const user: LocalUserRecord = {
    id: nanoid(),
    email: input.email,
    displayName: input.displayName,
    createdAt: now,
    updatedAt: now,
  };

  await writeJsonFile(getPath("users.json"), [user, ...users]);
  return user;
}

export async function getLocalMembership(input: {
  tenantId: string;
  userId: string;
}): Promise<LocalMembershipRecord | null> {
  const memberships = await listMemberships();
  return memberships.find((membership) => membership.tenantId === input.tenantId && membership.userId === input.userId) ?? null;
}

export async function createLocalMembership(input: {
  tenantId: string;
  userId: string;
  role: PlatformRole;
}): Promise<LocalMembershipRecord> {
  const memberships = await listMemberships();
  const now = new Date().toISOString();
  const membership: LocalMembershipRecord = {
    id: nanoid(),
    tenantId: input.tenantId,
    userId: input.userId,
    role: input.role,
    createdAt: now,
    updatedAt: now,
  };

  await writeJsonFile(getPath("memberships.json"), [membership, ...memberships]);
  return membership;
}

export async function listLocalMembershipsForUser(userId: string): Promise<LocalMembershipRecord[]> {
  const memberships = await listMemberships();
  return memberships.filter((membership) => membership.userId === userId);
}

export async function getLocalTenantBySlug(slug: string): Promise<LocalTenantRecord | null> {
  const tenants = await listTenants();
  return tenants.find((tenant) => tenant.slug === slug) ?? null;
}

export async function getLocalTenantById(id: string): Promise<LocalTenantRecord | null> {
  const tenants = await listTenants();
  return tenants.find((tenant) => tenant.id === id) ?? null;
}

export async function createLocalTenant(input: {
  slug: string;
  name: string;
  description?: string;
  defaultEnvironmentSlug: string;
}): Promise<LocalTenantRecord> {
  const now = new Date().toISOString();
  const tenants = await listTenants();
  const tenant: LocalTenantRecord = {
    id: nanoid(),
    slug: input.slug,
    name: input.name,
    description: input.description,
    defaultEnvironmentSlug: input.defaultEnvironmentSlug,
    createdAt: now,
    updatedAt: now,
  };

  await writeJsonFile(getPath("tenants.json"), [tenant, ...tenants]);
  return tenant;
}

export async function getLocalEnvironmentBySlug(input: {
  tenantId: string;
  slug: string;
}): Promise<LocalEnvironmentRecord | null> {
  const environments = await listEnvironments();
  return environments.find((environment) => environment.tenantId === input.tenantId && environment.slug === input.slug) ?? null;
}

export async function getLocalEnvironmentById(id: string): Promise<LocalEnvironmentRecord | null> {
  const environments = await listEnvironments();
  return environments.find((environment) => environment.id === id) ?? null;
}

export async function createLocalEnvironment(input: {
  tenantId: string;
  slug: string;
  name: string;
  isDefault: boolean;
}): Promise<LocalEnvironmentRecord> {
  const now = new Date().toISOString();
  const environments = await listEnvironments();
  const environment: LocalEnvironmentRecord = {
    id: nanoid(),
    tenantId: input.tenantId,
    slug: input.slug,
    name: input.name,
    isDefault: input.isDefault,
    createdAt: now,
    updatedAt: now,
  };

  await writeJsonFile(getPath("environments.json"), [environment, ...environments]);
  return environment;
}

export async function listLocalTenantMembershipSummariesForUser(userId: string): Promise<PlatformTenantMembershipSummary[]> {
  const memberships = await listLocalMembershipsForUser(userId);
  const tenants = await listTenants();
  return memberships.flatMap((membership) => {
    const tenant = tenants.find((candidate) => candidate.id === membership.tenantId);
    if (!tenant) {
      return [];
    }

    return [
      {
        tenantId: tenant.id,
        tenantSlug: tenant.slug,
        tenantName: tenant.name,
        defaultEnvironmentSlug: tenant.defaultEnvironmentSlug,
        role: membership.role,
      },
    ];
  });
}

export async function getLocalDraft(input: {
  tenantId: string;
  environmentId: string;
}): Promise<LocalDraftRecord | null> {
  const drafts = await listDrafts();
  return drafts.find((draft) => draft.tenantId === input.tenantId && draft.environmentId === input.environmentId) ?? null;
}

export async function upsertLocalDraft(input: {
  tenantId: string;
  environmentId: string;
  manifest: PlatformManifest;
}): Promise<LocalDraftRecord> {
  const drafts = await listDrafts();
  const now = new Date().toISOString();
  const existing = drafts.find((draft) => draft.tenantId === input.tenantId && draft.environmentId === input.environmentId);

  const nextDraft: LocalDraftRecord = existing
    ? {
        ...existing,
        manifest: input.manifest,
        updatedAt: now,
      }
    : {
        id: nanoid(),
        tenantId: input.tenantId,
        environmentId: input.environmentId,
        manifest: input.manifest,
        createdAt: now,
        updatedAt: now,
      };

  const filtered = drafts.filter((draft) => !(draft.tenantId === input.tenantId && draft.environmentId === input.environmentId));
  await writeJsonFile(getPath("drafts.json"), [nextDraft, ...filtered]);
  return nextDraft;
}

export async function listLocalVersions(input: {
  tenantId: string;
  environmentId: string;
}): Promise<LocalVersionRecord[]> {
  const versions = await listVersions();
  return sortByUpdatedDesc(
    versions.filter((version) => version.tenantId === input.tenantId && version.environmentId === input.environmentId),
  );
}

export async function createLocalVersion(input: {
  id?: string;
  tenantId: string;
  environmentId: string;
  versionNumber: number;
  manifest: PlatformManifest;
  status: PlatformVersionStatus;
  notes?: string | null;
  manifestPath?: string | null;
  gitCommitSha?: string | null;
  activatedAt?: string | null;
}): Promise<LocalVersionRecord> {
  const versions = await listVersions();
  const record: LocalVersionRecord = {
    id: input.id ?? nanoid(),
    tenantId: input.tenantId,
    environmentId: input.environmentId,
    versionNumber: input.versionNumber,
    manifest: input.manifest,
    status: input.status,
    notes: input.notes,
    manifestPath: input.manifestPath,
    gitCommitSha: input.gitCommitSha,
    activatedAt: input.activatedAt,
    createdAt: new Date().toISOString(),
  };

  await writeJsonFile(getPath("versions.json"), [record, ...versions]);
  return record;
}

export async function setLocalActiveVersion(input: {
  tenantId: string;
  environmentId: string;
  versionId: string;
}): Promise<void> {
  const versions = await listVersions();
  const updated = versions.map((version) => {
    if (version.tenantId !== input.tenantId || version.environmentId !== input.environmentId) {
      return version;
    }

    if (version.id === input.versionId) {
      return {
        ...version,
        status: "ACTIVE" satisfies PlatformVersionStatus,
        activatedAt: new Date().toISOString(),
      };
    }

    if (version.status === "ACTIVE") {
      return {
        ...version,
        status: "ROLLED_BACK" satisfies PlatformVersionStatus,
      };
    }

    return version;
  });

  await writeJsonFile(getPath("versions.json"), updated);
}

export async function getLocalActiveVersion(input: {
  tenantId: string;
  environmentId: string;
}): Promise<LocalVersionRecord | null> {
  const versions = await listLocalVersions(input);
  return versions.find((version) => version.status === "ACTIVE") ?? null;
}

export async function createLocalAuditEvent(input: {
  tenantId: string;
  environmentId?: string | null;
  actor?: PlatformActor | null;
  action: string;
  resourceType: string;
  resourceId: string;
  summary: string;
  payload?: Record<string, unknown>;
}): Promise<LocalAuditRecord> {
  const auditEvents = await listAudit();
  const record: LocalAuditRecord = {
    id: nanoid(),
    tenantId: input.tenantId,
    environmentId: input.environmentId ?? null,
    action: input.action,
    resourceType: input.resourceType,
    resourceId: input.resourceId,
    summary: input.summary,
    actorEmail: input.actor?.email ?? null,
    actorRole: input.actor?.role ?? null,
    createdAt: new Date().toISOString(),
    payload: input.payload ?? null,
  };

  await writeJsonFile(getPath("audit.json"), [record, ...auditEvents]);
  return record;
}

export async function listLocalAuditEvents(input: {
  tenantId: string;
  environmentId?: string;
}): Promise<LocalAuditRecord[]> {
  const auditEvents = await listAudit();
  return sortByUpdatedDesc(
    auditEvents.filter(
      (record) =>
        record.tenantId === input.tenantId &&
        (input.environmentId ? record.environmentId === input.environmentId : true),
    ),
  );
}

export async function listLocalPlatformRecords(input: {
  tenantId: string;
  environmentId: string;
  objectKey: string;
}): Promise<LocalPlatformRecord[]> {
  const records = await listRecords();
  return sortByUpdatedDesc(
    records.filter(
      (record) =>
        record.tenantId === input.tenantId &&
        record.environmentId === input.environmentId &&
        record.objectKey === input.objectKey,
    ),
  );
}

export async function upsertLocalPlatformRecord(input: {
  tenantId: string;
  environmentId: string;
  objectKey: string;
  recordId?: string;
  data: Record<string, unknown>;
  actor: PlatformActor;
}): Promise<LocalPlatformRecord> {
  const records = await listRecords();
  const existing = input.recordId
    ? records.find(
        (record) =>
          record.id === input.recordId &&
          record.tenantId === input.tenantId &&
          record.environmentId === input.environmentId &&
          record.objectKey === input.objectKey,
      )
    : null;
  const now = new Date().toISOString();

  const record: LocalPlatformRecord = existing
    ? {
        ...existing,
        data: input.data,
        updatedAt: now,
        updatedByEmail: input.actor.email,
      }
    : {
        id: nanoid(),
        tenantId: input.tenantId,
        environmentId: input.environmentId,
        objectKey: input.objectKey,
        data: input.data,
        createdAt: now,
        updatedAt: now,
        createdByEmail: input.actor.email,
        updatedByEmail: input.actor.email,
      };

  const filtered = records.filter((candidate) => candidate.id !== record.id);
  await writeJsonFile(getPath("records.json"), [record, ...filtered]);
  return record;
}

export async function deleteLocalPlatformRecord(input: {
  tenantId: string;
  environmentId: string;
  objectKey: string;
  recordId: string;
}): Promise<void> {
  const records = await listRecords();
  await writeJsonFile(
    getPath("records.json"),
    records.filter(
      (record) =>
        !(
          record.id === input.recordId &&
          record.tenantId === input.tenantId &&
          record.environmentId === input.environmentId &&
          record.objectKey === input.objectKey
        ),
    ),
  );
}

export async function createLocalWorkflowRun(input: {
  tenantId: string;
  environmentId: string;
  workflowId: string;
  workflowKey: string;
  input?: Record<string, unknown>;
}): Promise<LocalWorkflowRunRecord> {
  const runs = await listWorkflowRuns();
  const now = new Date().toISOString();
  const run: LocalWorkflowRunRecord = {
    id: nanoid(),
    tenantId: input.tenantId,
    environmentId: input.environmentId,
    workflowId: input.workflowId,
    workflowKey: input.workflowKey,
    status: "QUEUED" satisfies WorkflowRunStatus,
    input: input.input ?? null,
    output: null,
    logs: [],
    createdAt: now,
    updatedAt: now,
  };

  await writeJsonFile(getPath("workflow-runs.json"), [run, ...runs]);
  return run;
}

export async function listLocalQueuedWorkflowRuns(): Promise<LocalWorkflowRunRecord[]> {
  const runs = await listWorkflowRuns();
  return runs.filter((run) => run.status === "QUEUED").sort((left, right) => left.createdAt.localeCompare(right.createdAt));
}

export async function updateLocalWorkflowRun(input: {
  runId: string;
  status: WorkflowRunStatus;
  output?: Record<string, unknown> | null;
  appendLog?: Record<string, unknown>;
}): Promise<void> {
  const runs = await listWorkflowRuns();
  const nextRuns = runs.map((run) => {
    if (run.id !== input.runId) {
      return run;
    }

    return {
      ...run,
      status: input.status,
      output: input.output ?? run.output ?? null,
      logs: input.appendLog ? [...run.logs, input.appendLog] : run.logs,
      updatedAt: new Date().toISOString(),
    };
  });

  await writeJsonFile(getPath("workflow-runs.json"), nextRuns);
}

export async function listLocalWorkflowRuns(input: {
  tenantId: string;
  environmentId: string;
  workflowId?: string;
  workflowKey?: string;
}): Promise<LocalWorkflowRunRecord[]> {
  const runs = await listWorkflowRuns();
  return sortByUpdatedDesc(
    runs.filter((run) => {
      if (run.tenantId !== input.tenantId || run.environmentId !== input.environmentId) {
        return false;
      }

      if (input.workflowId && run.workflowId !== input.workflowId) {
        return false;
      }

      if (input.workflowKey && run.workflowKey !== input.workflowKey) {
        return false;
      }

      return true;
    }),
  );
}

export async function listLocalInvites(tenantId: string): Promise<LocalInviteRecord[]> {
  const invites = await listInvites();
  return sortByUpdatedDesc(
    invites.filter((invite) => invite.tenantId === tenantId),
  );
}

export async function createLocalInvite(input: {
  tenantId: string;
  tenantSlug: string;
  email: string;
  role: PlatformRole;
  token: string;
  inviteUrl: string;
  createdByEmail?: string | null;
  expiresAt: string;
}): Promise<LocalInviteRecord> {
  const invites = await listInvites();
  const record: LocalInviteRecord = {
    id: nanoid(),
    tenantId: input.tenantId,
    tenantSlug: input.tenantSlug,
    email: input.email,
    role: input.role,
    status: "pending",
    token: input.token,
    inviteUrl: input.inviteUrl,
    createdByEmail: input.createdByEmail ?? null,
    createdAt: new Date().toISOString(),
    expiresAt: input.expiresAt,
    acceptedAt: null,
    acceptedByUserId: null,
  };

  await writeJsonFile(getPath("invites.json"), [record, ...invites]);
  return record;
}

export async function getLocalInviteByToken(token: string): Promise<LocalInviteRecord | null> {
  const invites = await listInvites();
  return invites.find((invite) => invite.token === token) ?? null;
}

export async function acceptLocalInvite(input: {
  inviteId: string;
  userId: string;
}): Promise<LocalInviteRecord> {
  const invites = await listInvites();
  let acceptedInvite: LocalInviteRecord | null = null;

  const nextInvites = invites.map((invite) => {
    if (invite.id !== input.inviteId) {
      return invite;
    }

    acceptedInvite = {
      ...invite,
      status: new Date(invite.expiresAt).getTime() < Date.now() ? "expired" : "accepted",
      acceptedAt: new Date().toISOString(),
      acceptedByUserId: input.userId,
    };
    return acceptedInvite;
  });

  if (!acceptedInvite) {
    throw new Error("Invite not found.");
  }

  await writeJsonFile(getPath("invites.json"), nextInvites);
  return acceptedInvite;
}
