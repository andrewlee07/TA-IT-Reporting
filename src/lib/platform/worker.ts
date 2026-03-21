import { Prisma } from "@/generated/prisma/client";

import { getEnv } from "@/lib/env";
import {
  getLocalEnvironmentById,
  getLocalTenantById,
  listLocalQueuedWorkflowRuns,
  updateLocalWorkflowRun,
} from "@/lib/platform/local-store";
import { getRuntimeManifest } from "@/lib/platform/service";
import { getPrisma } from "@/lib/prisma";

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

async function processLocalQueuedRuns(): Promise<number> {
  const runs = await listLocalQueuedWorkflowRuns();
  let processed = 0;

  for (const run of runs) {
    const tenant = await getLocalTenantById(run.tenantId);
    const environment = await getLocalEnvironmentById(run.environmentId);
    if (!tenant || !environment) {
      await updateLocalWorkflowRun({
        runId: run.id,
        status: "FAILED",
        appendLog: {
          level: "error",
          message: "Tenant or environment not found for workflow run.",
          at: new Date().toISOString(),
        },
      });
      processed += 1;
      continue;
    }

    const manifest = await getRuntimeManifest({
      tenantSlug: tenant.slug,
      environmentSlug: environment.slug,
    });
    const workflow = manifest?.workflows.find((candidate) => candidate.id === run.workflowId || candidate.key === run.workflowKey);

    if (!workflow) {
      await updateLocalWorkflowRun({
        runId: run.id,
        status: "FAILED",
        appendLog: {
          level: "error",
          message: "Workflow definition not found.",
          at: new Date().toISOString(),
        },
      });
      processed += 1;
      continue;
    }

    await updateLocalWorkflowRun({
      runId: run.id,
      status: "RUNNING",
      appendLog: {
        level: "info",
        message: `Starting workflow ${workflow.name}.`,
        at: new Date().toISOString(),
      },
    });

    for (const node of workflow.nodes.sort((left, right) => left.position.y - right.position.y || left.position.x - right.position.x)) {
      await updateLocalWorkflowRun({
        runId: run.id,
        status: "RUNNING",
        appendLog: {
          level: "info",
          message: `Visited ${node.type} node "${node.label}".`,
          at: new Date().toISOString(),
        },
      });
    }

    await updateLocalWorkflowRun({
      runId: run.id,
      status: "SUCCEEDED",
      output: {
        completedNodes: workflow.nodes.length,
        workflowKey: workflow.key,
      },
      appendLog: {
        level: "info",
        message: `Completed workflow ${workflow.name}.`,
        at: new Date().toISOString(),
      },
    });
    processed += 1;
  }

  return processed;
}

async function processDatabaseQueuedRuns(): Promise<number> {
  const prisma = getPrisma();
  const runs = await prisma.platformWorkflowRun.findMany({
    where: { status: "QUEUED" },
    include: {
      tenant: true,
      environment: true,
    },
    orderBy: { createdAt: "asc" },
    take: 20,
  });

  let processed = 0;

  for (const run of runs) {
    const manifest = await getRuntimeManifest({
      tenantSlug: run.tenant.slug,
      environmentSlug: run.environment.slug,
    });
    const workflow = manifest?.workflows.find((candidate) => candidate.id === run.workflowId || candidate.key === run.workflowKey);

    if (!workflow) {
      await prisma.platformWorkflowRun.update({
        where: { id: run.id },
        data: {
          status: "FAILED",
          logs: toJsonValue([
            {
              level: "error",
              message: "Workflow definition not found.",
              at: new Date().toISOString(),
            },
          ]),
        },
      });
      processed += 1;
      continue;
    }

    await prisma.platformWorkflowRun.update({
      where: { id: run.id },
      data: {
        status: "RUNNING",
        startedAt: run.startedAt ?? new Date(),
        logs: toJsonValue([
          {
            level: "info",
            message: `Starting workflow ${workflow.name}.`,
            at: new Date().toISOString(),
          },
        ]),
      },
    });

    const logs: Array<Record<string, unknown>> = [
      {
        level: "info",
        message: `Starting workflow ${workflow.name}.`,
        at: new Date().toISOString(),
      },
      ...workflow.nodes
        .sort((left, right) => left.position.y - right.position.y || left.position.x - right.position.x)
        .map((node) => ({
          level: "info",
          message: `Visited ${node.type} node "${node.label}".`,
          at: new Date().toISOString(),
        })),
      {
        level: "info",
        message: `Completed workflow ${workflow.name}.`,
        at: new Date().toISOString(),
      },
    ];

    await prisma.platformWorkflowRun.update({
      where: { id: run.id },
      data: {
        status: "SUCCEEDED",
        startedAt: run.startedAt ?? new Date(),
        finishedAt: new Date(),
        output: toJsonValue({
          completedNodes: workflow.nodes.length,
          workflowKey: workflow.key,
        }),
        logs: toJsonValue(logs),
      },
    });
    processed += 1;
  }

  return processed;
}

export async function runWorkflowWorkerCycle(): Promise<number> {
  return withLocalFallback(processDatabaseQueuedRuns, processLocalQueuedRuns);
}

export async function startPlatformWorker(intervalMs = 10_000): Promise<void> {
  console.log(`[platform-worker] started with interval ${intervalMs}ms`);

  while (true) {
    try {
      const processed = await runWorkflowWorkerCycle();
      console.log(`[platform-worker] processed ${processed} queued run(s)`);
    } catch (error) {
      console.error("[platform-worker] cycle failed", error);
    }

    await new Promise((resolve) => setTimeout(resolve, intervalMs));
  }
}
