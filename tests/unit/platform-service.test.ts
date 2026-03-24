import path from "node:path";
import { rm } from "node:fs/promises";

import { beforeEach, describe, expect, it } from "vitest";

import { resetEnvCache } from "@/lib/env";
import { createSignedPlatformSessionValue } from "@/lib/platform/auth";
import { createStarterManifest } from "@/lib/platform/defaults";
import {
  createLocalEnvironment,
  createLocalMembership,
  createLocalTenant,
  createLocalUser,
  createLocalVersion,
  listLocalPlatformRecords,
  setLocalActiveVersion,
  upsertLocalDraft,
  upsertLocalPlatformRecord,
} from "@/lib/platform/local-store";
import {
  acceptPlatformInvite,
  createTenantInvite,
  deletePlatformRecord,
  getCurrentPlatformSession,
  getDraftPreviewManifest,
  getPlatformBootstrap,
  listWorkflowRuns,
  previewAgentInvocation,
  publishDraftManifest,
  queueWorkflowRun,
  saveAgentDefinition,
  saveLayoutDefinition,
  savePageDefinition,
  savePlatformRecord,
  saveWorkflowDefinition,
  runNotificationMaintenanceCycle,
} from "@/lib/platform/service";
import { runWorkflowWorkerCycle } from "@/lib/platform/worker";

const STORAGE_DIR = ".storage-test-platform-service";

function devRequest(): Request {
  return new Request("http://localhost/platform");
}

function sessionRequest(email: string, name: string, headers?: Record<string, string>): Request {
  const session = createSignedPlatformSessionValue({ email, name });
  return new Request("http://localhost/platform", {
    headers: {
      cookie: `ta_platform_session=${encodeURIComponent(session)}`,
      ...headers,
    },
  });
}

async function seedTenant(options: {
  tenantSlug: string;
  environmentSlug?: string;
  activeVersion?: boolean;
}) {
  const tenant = await createLocalTenant({
    slug: options.tenantSlug,
    name: options.tenantSlug,
    description: `${options.tenantSlug} tenant`,
    defaultEnvironmentSlug: options.environmentSlug ?? "development",
  });
  const environment = await createLocalEnvironment({
    tenantId: tenant.id,
    slug: options.environmentSlug ?? "development",
    name: "Development",
    isDefault: true,
  });
  const manifest = createStarterManifest(options.tenantSlug, options.tenantSlug);
  manifest.environment = {
    slug: environment.slug,
    name: environment.name,
  };

  await upsertLocalDraft({
    tenantId: tenant.id,
    environmentId: environment.id,
    manifest,
  });

  if (options.activeVersion) {
    const version = await createLocalVersion({
      tenantId: tenant.id,
      environmentId: environment.id,
      versionNumber: 1,
      manifest,
      status: "ACTIVE",
      notes: "Seeded version",
      activatedAt: new Date().toISOString(),
    });
    await setLocalActiveVersion({
      tenantId: tenant.id,
      environmentId: environment.id,
      versionId: version.id,
    });
  }

  return { tenant, environment, manifest };
}

beforeEach(async () => {
  delete process.env.DATABASE_URL;
  process.env.LOCAL_STORAGE_DIR = STORAGE_DIR;
  process.env.PLATFORM_LOCAL_DEV_MODE = "true";
  process.env.PLATFORM_SESSION_SECRET = "platform-test-secret";
  process.env.PLATFORM_DEV_ACTOR_EMAIL = "builder@teacheractive.local";
  process.env.PLATFORM_DEV_ACTOR_NAME = "Local Builder";
  process.env.PLATFORM_DEV_ACTOR_ROLE = "SUPER_ADMIN";
  resetEnvCache();
  await rm(path.resolve(process.cwd(), STORAGE_DIR), { recursive: true, force: true });
});

describe("platform hardening and workflow depth", () => {
  it("derives the effective role from membership instead of forged request headers", async () => {
    const { tenant } = await seedTenant({ tenantSlug: "teacheractive", activeVersion: true });
    const user = await createLocalUser({
      email: "member@teacheractive.local",
      displayName: "Member User",
    });
    await createLocalMembership({
      tenantId: tenant.id,
      userId: user.id,
      role: "USER",
    });

    const bootstrap = await getPlatformBootstrap({
      tenantSlug: "teacheractive",
      request: sessionRequest("member@teacheractive.local", "Member User", {
        "x-platform-role": "SUPER_ADMIN",
        "x-platform-user-email": "admin@forged.local",
      }),
    });

    expect(bootstrap.actor.email).toBe("member@teacheractive.local");
    expect(bootstrap.actor.role).toBe("USER");
  });

  it("rejects cross-tenant record updates and deletes by leaked id", async () => {
    const alpha = await seedTenant({ tenantSlug: "alpha", activeVersion: true });
    await seedTenant({ tenantSlug: "bravo", activeVersion: true });

    const record = await savePlatformRecord({
      tenantSlug: "alpha",
      objectKey: "booking_request",
      request: devRequest(),
      data: {
        request_title: "Alpha request",
        school_name: "Alpha School",
        contact_email: "alpha@example.com",
        vacancies: 1,
        daily_rate: 220,
        days_requested: 4,
        status: "Open",
      },
    });

    await expect(
      savePlatformRecord({
        tenantSlug: "bravo",
        objectKey: "booking_request",
        recordId: record.id,
        request: devRequest(),
        data: {
          request_title: "Hijack",
          school_name: "Bravo School",
          contact_email: "bravo@example.com",
          vacancies: 1,
          daily_rate: 200,
          days_requested: 3,
          status: "Review",
        },
      }),
    ).rejects.toThrow(/not found/i);

    await expect(
      deletePlatformRecord({
        tenantSlug: "bravo",
        objectKey: "booking_request",
        recordId: record.id,
        request: devRequest(),
      }),
    ).rejects.toThrow(/not found/i);

    const remaining = await listLocalPlatformRecords({
      tenantId: alpha.tenant.id,
      environmentId: alpha.environment.id,
      objectKey: "booking_request",
    });
    expect(remaining.some((candidate) => candidate.id === record.id)).toBe(true);
  });

  it("allows builder draft edits without publish but blocks runtime execution until a version is active", async () => {
    await seedTenant({ tenantSlug: "draft-only", activeVersion: false });

    await expect(
      savePageDefinition({
        tenantSlug: "draft-only",
        request: devRequest(),
        page: {
          title: "Builder workspace",
          key: "builder_workspace",
          route: "builder-workspace",
        },
      }),
    ).resolves.toMatchObject({
      key: "builder_workspace",
    });

    await expect(
      queueWorkflowRun({
        tenantSlug: "draft-only",
        workflowId: "booking_triage",
        request: devRequest(),
      }),
    ).rejects.toThrow(/published runtime/i);
  });

  it("rejects invalid empty workflow and agent definitions before they enter the draft manifest", async () => {
    await seedTenant({ tenantSlug: "validation", activeVersion: false });

    await expect(
      saveWorkflowDefinition({
        tenantSlug: "validation",
        request: devRequest(),
        workflow: {
          name: "   ",
        },
      }),
    ).rejects.toThrow(/workflow name is required/i);

    await expect(
      saveAgentDefinition({
        tenantSlug: "validation",
        request: devRequest(),
        agent: {
          name: "   ",
          scope: "workspace",
          modelProviderId: "provider-openai-gpt5mini",
          prompt: "   ",
        },
      }),
    ).rejects.toThrow(/agent name is required/i);
  });

  it("persists workflow runs through queue and worker execution using the published manifest", async () => {
    await seedTenant({ tenantSlug: "workflow-runner", activeVersion: true });

    const queued = await queueWorkflowRun({
      tenantSlug: "workflow-runner",
      workflowId: "booking_triage",
      request: devRequest(),
      payload: {
        launchedFrom: "test",
      },
    });

    expect(queued.status).toBe("QUEUED");

    const processed = await runWorkflowWorkerCycle();
    expect(processed).toBeGreaterThan(0);

    const runs = await listWorkflowRuns({
      tenantSlug: "workflow-runner",
      workflowId: "booking_triage",
      request: devRequest(),
    });

    expect(runs[0]?.status).toBe("SUCCEEDED");
    expect(runs[0]?.logs.length).toBeGreaterThan(0);
  });

  it("processes due notification retries from the maintenance cycle without a fresh outbox event", async () => {
    const { tenant, environment } = await seedTenant({ tenantSlug: "notification-maintenance", activeVersion: true });

    const eventRecord = await upsertLocalPlatformRecord({
      tenantId: tenant.id,
      environmentId: environment.id,
      objectKey: "__system_outbox_event",
      actor: {
        email: "system@test.local",
        name: "System",
        role: "SUPER_ADMIN",
      },
      data: {
        type: "workflow.run.updated",
        source: "workflow",
        resourceType: "workflow_run",
        resourceId: "run_123",
        tenantSlug: tenant.slug,
        environmentSlug: environment.slug,
        emittedAt: new Date().toISOString(),
        processedAt: new Date().toISOString(),
        payload: {
          workflowKey: "booking_triage",
          status: "FAILED",
        },
      },
    });

    await upsertLocalPlatformRecord({
      tenantId: tenant.id,
      environmentId: environment.id,
      objectKey: "__system_notification_delivery",
      actor: {
        email: "system@test.local",
        name: "System",
        role: "SUPER_ADMIN",
      },
      data: {
        eventId: eventRecord.id,
        ruleKey: "workflow_run_changed",
        channelKey: "in_app_primary",
        templateKey: "workflow_run_status",
        status: "retrying",
        severity: "info",
        subject: "Workflow update",
        body: "booking_triage changed to FAILED.",
        provider: "in_app",
        attemptCount: 0,
        maxAttempts: 1,
        nextRetryAt: new Date(Date.now() - 60_000).toISOString(),
        resolvedPayload: {
          workflowKey: "booking_triage",
          status: "FAILED",
        },
      },
    });

    const result = await runNotificationMaintenanceCycle();
    expect(result.scopes).toBeGreaterThan(0);
    expect(result.deliveries).toBe(1);

    const deliveries = await listLocalPlatformRecords({
      tenantId: tenant.id,
      environmentId: environment.id,
      objectKey: "__system_notification_delivery",
    });
    expect(deliveries[0]?.data.status).toBe("sent");

    const attempts = await listLocalPlatformRecords({
      tenantId: tenant.id,
      environmentId: environment.id,
      objectKey: "__system_notification_delivery_attempt",
    });
    expect(attempts.some((attempt) => attempt.data.lifecycleStage === "attempt_succeeded")).toBe(true);

    const channelHealth = await listLocalPlatformRecords({
      tenantId: tenant.id,
      environmentId: environment.id,
      objectKey: "__system_notification_channel_health",
    });
    expect(channelHealth[0]?.data.successCount).toBe(1);
  });

  it("allows builder-side agent preview on draft definitions and preserves the published flow", async () => {
    await seedTenant({ tenantSlug: "agent-preview", activeVersion: false });

    await saveAgentDefinition({
      tenantSlug: "agent-preview",
      request: devRequest(),
      agent: {
        name: "Coverage Analyst",
        key: "coverage_analyst",
        scope: "workspace",
        modelProviderId: "provider-openai-gpt5mini",
        prompt: "Summarize booking demand with masked values only.",
        objectKeys: ["booking_request"],
        allowedToolIds: ["tool-booking-triage"],
      },
    });

    const draftPreview = await previewAgentInvocation({
      tenantSlug: "agent-preview",
      agentId: "coverage_analyst",
      objectKey: "booking_request",
      request: devRequest(),
    });

    expect(draftPreview.agent.key).toBe("coverage_analyst");
    expect(draftPreview.metadata.allowedByPolicy).toBe(true);

    await publishDraftManifest({
      tenantSlug: "agent-preview",
      request: devRequest(),
      notes: "Publish draft for agent preview",
    });

    const preview = await previewAgentInvocation({
      tenantSlug: "agent-preview",
      agentId: "coverage_analyst",
      objectKey: "booking_request",
      request: devRequest(),
    });

    expect(preview.agent.key).toBe(draftPreview.agent.key);
    expect(preview.metadata.allowedByPolicy).toBe(true);
  });

  it("accepts invite-based builder sessions and exposes tenant memberships", async () => {
    await seedTenant({ tenantSlug: "invite-beta", activeVersion: false });

    const invite = await createTenantInvite({
      tenantSlug: "invite-beta",
      request: devRequest(),
      email: "builder@example.com",
      role: "BUILDER_ADMIN",
    });

    const accepted = await acceptPlatformInvite({
      token: invite.token,
      email: "builder@example.com",
      name: "Beta Builder",
    });

    const session = await getCurrentPlatformSession({
      request: new Headers({
        cookie: `ta_platform_session=${encodeURIComponent(accepted.sessionValue)}`,
      }),
    });

    expect(session.actor?.email).toBe("builder@example.com");
    expect(session.memberships.some((membership) => membership.tenantSlug === "invite-beta")).toBe(true);

    const bootstrap = await getPlatformBootstrap({
      tenantSlug: "invite-beta",
      request: new Headers({
        cookie: `ta_platform_session=${encodeURIComponent(accepted.sessionValue)}`,
      }),
    });

    expect(bootstrap.actor.role).toBe("BUILDER_ADMIN");
    expect(bootstrap.session.memberships[0]?.tenantSlug).toBe("invite-beta");
  });

  it("allows builder draft preview without publish and records preview audit", async () => {
    await seedTenant({ tenantSlug: "draft-preview", activeVersion: false });

    const manifest = await getDraftPreviewManifest({
      tenantSlug: "draft-preview",
      request: devRequest(),
      route: "operations-overview",
    });

    expect(manifest.pages.some((page) => page.route === "operations-overview")).toBe(true);

    const bootstrap = await getPlatformBootstrap({
      tenantSlug: "draft-preview",
      request: devRequest(),
    });

    expect(bootstrap.auditEvents[0]?.action).toBe("preview.opened");
  });

  it("rejects invalid designer bindings before layout metadata is stored", async () => {
    await seedTenant({ tenantSlug: "designer-validation", activeVersion: false });
    const bootstrap = await getPlatformBootstrap({
      tenantSlug: "designer-validation",
      request: devRequest(),
    });
    const layout = structuredClone(bootstrap.draftManifest.layouts[0]!);
    layout.sections[0]!.components[0] = {
      ...layout.sections[0]!.components[0]!,
      kind: "workflow_launcher",
      workflowKey: undefined,
      binding: {},
    };

    await expect(
      saveLayoutDefinition({
        tenantSlug: "designer-validation",
        request: devRequest(),
        layout,
      }),
    ).rejects.toThrow(/workflow binding/i);
  });
});
