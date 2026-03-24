import { Prisma } from "@/generated/prisma/client";

import { prepareAgentInvocation } from "@/lib/platform/agent-gateway";
import { evaluateRuleAsBoolean, evaluateRuleExpression } from "@/lib/platform/rule-engine";
import { deliverNotification, executeAgentWithProvider, summarizeAgentRunCost } from "@/lib/platform/runtime-adapters";
import { getEnv } from "@/lib/env";
import { createOutboxWorker, createWorkflowRunWorker, enqueueWorkflowRun } from "@/lib/platform/execution-bus";
import {
  getLocalEnvironmentById,
  getLocalTenantById,
  listLocalEnvironments,
  listLocalPlatformRecords,
  listLocalQueuedWorkflowRuns,
  listLocalTenants,
  listLocalWorkflowRuns,
  updateLocalWorkflowRun,
  upsertLocalPlatformRecord,
} from "@/lib/platform/local-store";
import type {
  AgentTrace,
  PlatformWorkflowRunRecord as WorkflowRunRecord,
  PlatformWorkflowRunRecord as WorkflowRunShape,
  WorkflowDefinition,
  WorkflowEdgeDefinition,
  WorkflowNodeDefinition,
} from "@/lib/platform/types";
import {
  buildAgentTrace,
  getRuntimeManifest,
  processPendingOutboxEvents,
  queueWorkflowRun,
  validateAgentOutputSchema,
} from "@/lib/platform/service";
import { getPrisma } from "@/lib/prisma";

const SYSTEM_OBJECT_KEYS = {
  delivery: "__system_notification_delivery",
  alert: "__system_alert",
  agentRun: "__system_agent_run",
  costLedger: "__system_cost_ledger",
  deadLetter: "__system_dead_letter",
  approvalTask: "__system_approval_task",
} as const;

interface RuntimeState {
  cursor: number;
  context: Record<string, unknown>;
  retryCounts: Record<string, number>;
  waitUntil?: string;
  waitNodeId?: string;
  approvalTaskId?: string;
  pauseReason?: "wait" | "approval" | "subflow";
  nextRetryAt?: string;
  subflowRunId?: string;
  pendingAgentRunId?: string;
}

interface WorkflowRuntimeContext {
  tenantId: string;
  environmentId: string;
  tenantSlug: string;
  environmentSlug: string;
}

function toJsonValue(value: unknown): Prisma.InputJsonValue {
  return JSON.parse(JSON.stringify(value)) as Prisma.InputJsonValue;
}

function nowIso(): string {
  return new Date().toISOString();
}

function systemActor() {
  return {
    email: "platform-worker@local.test",
    name: "Platform Worker",
    role: "SUPER_ADMIN" as const,
  };
}

function systemObjectKey(key: keyof typeof SYSTEM_OBJECT_KEYS): string {
  return SYSTEM_OBJECT_KEYS[key];
}

function getRuntimeState(run: {
  input?: Record<string, unknown> | null;
  output?: Record<string, unknown> | null;
}): RuntimeState {
  const output = (run.output ?? {}) as Record<string, unknown>;
  const raw = output.runtimeState as Record<string, unknown> | undefined;
  return {
    cursor: typeof raw?.cursor === "number" ? raw.cursor : 0,
    context: typeof raw?.context === "object" && raw?.context ? (raw.context as Record<string, unknown>) : { ...(run.input ?? {}) },
    retryCounts: typeof raw?.retryCounts === "object" && raw?.retryCounts ? (raw.retryCounts as Record<string, number>) : {},
    waitUntil: typeof raw?.waitUntil === "string" ? raw.waitUntil : undefined,
    waitNodeId: typeof raw?.waitNodeId === "string" ? raw.waitNodeId : undefined,
    approvalTaskId: typeof raw?.approvalTaskId === "string" ? raw.approvalTaskId : undefined,
    pauseReason: typeof raw?.pauseReason === "string" ? (raw.pauseReason as RuntimeState["pauseReason"]) : undefined,
    nextRetryAt: typeof raw?.nextRetryAt === "string" ? raw.nextRetryAt : undefined,
    subflowRunId: typeof raw?.subflowRunId === "string" ? raw.subflowRunId : undefined,
    pendingAgentRunId: typeof raw?.pendingAgentRunId === "string" ? raw.pendingAgentRunId : undefined,
  };
}

function withRuntimeState(output: Record<string, unknown> | null | undefined, state: RuntimeState, extra?: Record<string, unknown>) {
  return {
    ...(output ?? {}),
    ...(extra ?? {}),
    runtimeState: state,
  };
}

function orderedNodes(workflow: WorkflowDefinition): WorkflowNodeDefinition[] {
  return [...workflow.nodes].sort((left, right) => left.position.y - right.position.y || left.position.x - right.position.x);
}

function edgeLabelMatches(edge: WorkflowEdgeDefinition, expected: boolean): boolean {
  const label = edge.label?.trim().toLowerCase();
  if (!label) {
    return expected;
  }
  if (expected) {
    return ["true", "yes", "success", "match"].includes(label);
  }
  return ["false", "no", "failure", "else"].includes(label);
}

function findNextNodeIndex(workflow: WorkflowDefinition, currentNode: WorkflowNodeDefinition, ordered: WorkflowNodeDefinition[], expectedCondition?: boolean): number {
  const outgoing = workflow.edges.filter((edge) => edge.sourceId === currentNode.id);
  const matching = typeof expectedCondition === "boolean" ? outgoing.find((edge) => edgeLabelMatches(edge, expectedCondition)) : outgoing[0];
  if (matching) {
    const targetIndex = ordered.findIndex((node) => node.id === matching.targetId);
    if (targetIndex >= 0) {
      return targetIndex;
    }
  }

  return ordered.findIndex((node) => node.id === currentNode.id) + 1;
}

function appendLog(logs: Array<Record<string, unknown>>, level: string, message: string, extra?: Record<string, unknown>) {
  return [
    ...logs,
    {
      level,
      message,
      at: nowIso(),
      ...(extra ?? {}),
    },
  ];
}

async function listSystemRecords(input: WorkflowRuntimeContext & { objectKey: string }) {
  return listLocalPlatformRecords({
    tenantId: input.tenantId,
    environmentId: input.environmentId,
    objectKey: input.objectKey,
  });
}

async function createSystemRecord(input: WorkflowRuntimeContext & { objectKey: string; data: Record<string, unknown> }) {
  return upsertLocalPlatformRecord({
    tenantId: input.tenantId,
    environmentId: input.environmentId,
    objectKey: input.objectKey,
    data: input.data,
    actor: systemActor(),
  });
}

async function updateSystemRecord(input: WorkflowRuntimeContext & { objectKey: string; recordId: string; data: Record<string, unknown> }) {
  return upsertLocalPlatformRecord({
    tenantId: input.tenantId,
    environmentId: input.environmentId,
    objectKey: input.objectKey,
    recordId: input.recordId,
    data: input.data,
    actor: systemActor(),
  });
}

function getExponentialRetryAt(retryCount: number): string {
  const schedule = [60_000, 5 * 60_000, 15 * 60_000];
  const delay = schedule[Math.max(0, Math.min(retryCount - 1, schedule.length - 1))] ?? schedule[schedule.length - 1]!;
  return new Date(Date.now() + delay).toISOString();
}

async function markDuePausedRunsQueued(context: WorkflowRuntimeContext, useDatabase: boolean): Promise<void> {
  const now = Date.now();
  if (useDatabase) {
    const prisma = getPrisma();
    const paused = await prisma.platformWorkflowRun.findMany({
      where: {
        tenantId: context.tenantId,
        environmentId: context.environmentId,
        status: "PAUSED",
      },
      orderBy: { updatedAt: "asc" },
      take: 25,
    });

    for (const run of paused) {
      const output = (run.output as Record<string, unknown> | null) ?? {};
      const state = getRuntimeState({
        input: (run.input as Record<string, unknown> | null) ?? null,
        output,
      });
      if (state.waitUntil && new Date(state.waitUntil).getTime() <= now) {
        await prisma.platformWorkflowRun.update({
          where: { id: run.id },
          data: {
            status: "QUEUED",
            output: toJsonValue(withRuntimeState(output, { ...state, waitUntil: undefined, waitNodeId: undefined, pauseReason: undefined })),
            logs: toJsonValue(
              appendLog((run.logs as Array<Record<string, unknown>> | null) ?? [], "info", "Wait duration elapsed. Re-queued workflow run."),
            ),
          },
        });
        await enqueueWorkflowRun(run.id).catch(() => undefined);
      }
    }
    return;
  }

  const paused = (await listLocalWorkflowRuns({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
  })).filter((run) => run.status === "PAUSED");

  for (const run of paused) {
    const state = getRuntimeState({
      input: (run.input as Record<string, unknown> | null) ?? null,
      output: (run.output as Record<string, unknown> | null) ?? null,
    });
    if (state.waitUntil && new Date(state.waitUntil).getTime() <= now) {
      await updateLocalWorkflowRun({
        runId: run.id,
        status: "QUEUED",
        output: withRuntimeState((run.output as Record<string, unknown> | null) ?? null, {
          ...state,
          waitUntil: undefined,
          waitNodeId: undefined,
          pauseReason: undefined,
        }),
        appendLog: {
          level: "info",
          message: "Wait duration elapsed. Re-queued workflow run.",
          at: nowIso(),
        },
      });
    }
  }
}

async function createApprovalTask(context: WorkflowRuntimeContext, run: WorkflowRunShape, node: WorkflowNodeDefinition, approverRole: "SUPER_ADMIN" | "BUILDER_ADMIN" | "USER", instructions?: string) {
  return createSystemRecord({
    ...context,
    objectKey: systemObjectKey("approvalTask"),
    data: {
      workflowRunId: run.id,
      workflowKey: run.workflowKey,
      nodeId: node.id,
      nodeLabel: node.label,
      approverRole,
      instructions,
      status: "pending",
    },
  });
}

async function listApprovalTasksForRun(context: WorkflowRuntimeContext, workflowRunId: string) {
  const records = await listSystemRecords({
    ...context,
    objectKey: systemObjectKey("approvalTask"),
  });
  return records.filter((record) => record.data.workflowRunId === workflowRunId);
}

async function getWorkflowRunById(
  context: WorkflowRuntimeContext,
  runId: string,
  useDatabase: boolean,
): Promise<WorkflowRunShape | null> {
  if (useDatabase) {
    const prisma = getPrisma();
    const run = await prisma.platformWorkflowRun.findUnique({
      where: { id: runId },
    });
    if (!run) {
      return null;
    }

    return {
      ...run,
      workflowId: run.workflowId,
      workflowKey: run.workflowKey,
      status: run.status,
      input: (run.input as Record<string, unknown> | null) ?? null,
      output: (run.output as Record<string, unknown> | null) ?? null,
      logs: (run.logs as Array<Record<string, unknown>> | null) ?? [],
      startedAt: run.startedAt?.toISOString() ?? null,
      finishedAt: run.finishedAt?.toISOString() ?? null,
      createdAt: run.createdAt.toISOString(),
      updatedAt: run.updatedAt.toISOString(),
    };
  }

  const runs = await listLocalWorkflowRuns({
    tenantId: context.tenantId,
    environmentId: context.environmentId,
  });
  return runs.find((run) => run.id === runId) ?? null;
}

async function persistDbWorkflowRun(input: {
  runId: string;
  status: WorkflowRunRecord["status"];
  logs: Array<Record<string, unknown>>;
  output?: Record<string, unknown> | null;
  startedAt?: boolean;
  finishedAt?: boolean;
}) {
  const prisma = getPrisma();
  await prisma.platformWorkflowRun.update({
    where: { id: input.runId },
    data: {
      status: input.status,
      logs: toJsonValue(input.logs),
      output: toJsonValue(input.output ?? null),
      startedAt: input.startedAt ? new Date() : undefined,
      finishedAt: input.finishedAt ? new Date() : undefined,
    },
  });
}

async function persistLocalWorkflowRun(input: {
  runId: string;
  status: WorkflowRunRecord["status"];
  logs: Array<Record<string, unknown>>;
  output?: Record<string, unknown> | null;
}) {
  await updateLocalWorkflowRun({
    runId: input.runId,
    status: input.status,
    output: input.output ?? null,
    appendLog: input.logs[input.logs.length - 1],
  });
}

async function recordDeadLetter(context: WorkflowRuntimeContext, eventId: string, type: string, reason: string, payload: Record<string, unknown>) {
  await createSystemRecord({
    ...context,
    objectKey: systemObjectKey("deadLetter"),
    data: {
      eventId,
      type,
      reason,
      payload,
    },
  });
}

async function recordRuntimeAlert(context: WorkflowRuntimeContext, title: string, summary: string, sourceId: string) {
  await createSystemRecord({
    ...context,
    objectKey: systemObjectKey("alert"),
    data: {
      category: "runtime",
      severity: "warning",
      title,
      summary,
      sourceId,
    },
  });
}

async function recordNotificationDelivery(context: WorkflowRuntimeContext, data: Record<string, unknown>) {
  await createSystemRecord({
    ...context,
    objectKey: systemObjectKey("delivery"),
    data,
  });
}

async function recordAgentRun(context: WorkflowRuntimeContext, data: Record<string, unknown>, recordId?: string) {
  const runRecord = recordId
    ? await updateSystemRecord({
        ...context,
        objectKey: systemObjectKey("agentRun"),
        recordId,
        data,
      })
    : await createSystemRecord({
        ...context,
        objectKey: systemObjectKey("agentRun"),
        data,
      });

  await createSystemRecord({
    ...context,
    objectKey: systemObjectKey("costLedger"),
    data: {
      category: "agent_run",
      referenceId: runRecord.id,
      providerKey: data.modelProviderKey,
      amountUsd: data.costUsd,
      tokensIn: data.tokensIn,
      tokensOut: data.tokensOut,
      summary: `Workflow model-call cost for ${data.agentKey}.`,
    },
  });

  return runRecord;
}

async function executeWorkflowNode(input: {
  context: WorkflowRuntimeContext;
  manifest: Awaited<ReturnType<typeof getRuntimeManifest>>;
  workflow: WorkflowDefinition;
  node: WorkflowNodeDefinition;
  ordered: WorkflowNodeDefinition[];
  run: WorkflowRunShape;
  state: RuntimeState;
  logs: Array<Record<string, unknown>>;
  useDatabase: boolean;
}): Promise<{
  status: WorkflowRunRecord["status"];
  state: RuntimeState;
  logs: Array<Record<string, unknown>>;
  advance: boolean;
  finished?: boolean;
}> {
  const { context, manifest, workflow, node, ordered, run, state } = input;
  const logs = [...input.logs];

  if (!manifest) {
    throw new Error("Published runtime manifest is missing.");
  }

  if (node.type === "condition") {
    const expression = "expression" in node.config ? node.config.expression : "true";
    let result = false;
    try {
      result = evaluateRuleAsBoolean({ mode: "text", expression }, state.context);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "Condition evaluation failed.";
      logs.push({
        level: "warning",
        message: `Condition ${node.label} fell back to false: ${errorMessage}`,
        at: nowIso(),
      });
    }
    logs.push({ level: "info", message: `Condition ${node.label} resolved ${result ? "true" : "false"}.`, at: nowIso() });
    return {
      status: "RUNNING",
      state: {
        ...state,
        cursor: findNextNodeIndex(workflow, node, ordered, result),
      },
      logs,
      advance: false,
    };
  }

  if (node.type === "formula" && "expression" in node.config && "outputKey" in node.config) {
    const result = evaluateRuleExpression({ mode: "text", expression: node.config.expression }, state.context);
    logs.push({ level: "info", message: `Formula ${node.label} stored ${node.config.outputKey}.`, at: nowIso() });
    return {
      status: "RUNNING",
      state: {
        ...state,
        cursor: state.cursor + 1,
        context: {
          ...state.context,
          [node.config.outputKey]: result,
        },
      },
      logs,
      advance: false,
    };
  }

  if (node.type === "webhook" && "url" in node.config) {
    const response = await fetch(node.config.url, {
      method: node.config.method,
      headers: {
        "content-type": "application/json",
      },
      body: node.config.bodyTemplate ? JSON.stringify({ template: node.config.bodyTemplate, context: state.context }) : undefined,
    });
    if (!response.ok) {
      throw new Error(`Webhook ${node.label} failed with status ${response.status}.`);
    }
    logs.push({ level: "info", message: `Webhook ${node.label} delivered.`, at: nowIso() });
    return {
      status: "RUNNING",
      state: { ...state, cursor: state.cursor + 1 },
      logs,
      advance: false,
    };
  }

  if (node.type === "notification" && "channel" in node.config) {
    const configuredChannel = node.config.channel;
    const kind: "email" | "slack_style" | "in_app" =
      configuredChannel === "email" ? "email" : configuredChannel === "slack" ? "slack_style" : "in_app";
    const manifestChannel =
      manifest.notifications.channels.find((channel) => channel.key === configuredChannel) ??
      manifest.notifications.channels.find((channel) => channel.kind === kind && channel.enabled);
    const destination = node.config.recipient?.trim() || manifestChannel?.destination;
    const deliveryChannel = manifestChannel ?? {
      id: `workflow-${node.id}`,
      key: `workflow-${node.id}`,
      name: `${node.label} delivery`,
      kind,
      enabled: true,
      destination,
    };
    const body = node.config.message || `Workflow ${workflow.name} executed ${node.label}.`;
    try {
      const result = await deliverNotification({
        channel: deliveryChannel,
        severity: "info",
        subject: `${workflow.name} notification`,
        body,
        event: {
          id: run.id,
          type: "workflow.node.notification",
          tenantSlug: context.tenantSlug,
          environmentSlug: context.environmentSlug,
          emittedAt: nowIso(),
          source: "workflow",
          resourceType: "workflow_run",
          resourceId: run.id,
          payload: {
            workflowKey: workflow.key,
            nodeId: node.id,
          },
        },
      });
      await recordNotificationDelivery(context, {
        eventId: run.id,
        ruleKey: `workflow-node-${node.id}`,
        channelKey: deliveryChannel.key,
        templateKey: "workflow_node",
        status: "sent",
        severity: "info",
        body,
        subject: `${workflow.name} notification`,
        destination: result.destination,
        deliveredAt: result.deliveredAt,
      });
      logs.push({ level: "info", message: `Notification ${node.label} delivered via ${kind}.`, at: nowIso() });
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : "Notification delivery failed.";
      await recordNotificationDelivery(context, {
        eventId: run.id,
        ruleKey: `workflow-node-${node.id}`,
        channelKey: deliveryChannel.key,
        templateKey: "workflow_node",
        status: "failed",
        severity: "warning",
        body,
        subject: `${workflow.name} notification`,
        destination,
        errorMessage,
      });
      throw new Error(errorMessage);
    }

    return {
      status: "RUNNING",
      state: { ...state, cursor: state.cursor + 1 },
      logs,
      advance: false,
    };
  }

  if (node.type === "wait" && "durationMinutes" in node.config) {
    const waitUntil = state.waitNodeId === node.id && state.waitUntil ? new Date(state.waitUntil).getTime() : null;
    if (waitUntil && waitUntil <= Date.now()) {
      logs.push({ level: "info", message: `Wait ${node.label} elapsed.`, at: nowIso() });
      return {
        status: "RUNNING",
        state: { ...state, cursor: state.cursor + 1, waitUntil: undefined, waitNodeId: undefined, pauseReason: undefined },
        logs,
        advance: false,
      };
    }

    const nextWaitUntil = new Date(Date.now() + node.config.durationMinutes * 60 * 1000).toISOString();
    logs.push({ level: "info", message: `Paused at wait node ${node.label} until ${nextWaitUntil}.`, at: nowIso() });
    return {
      status: "PAUSED",
      state: { ...state, waitUntil: nextWaitUntil, waitNodeId: node.id, pauseReason: "wait" },
      logs,
      advance: false,
    };
  }

  if (node.type === "approval" && "approverRole" in node.config) {
    const existingTasks = await listApprovalTasksForRun(context, run.id);
    const task = existingTasks.find((candidate) => candidate.data.nodeId === node.id);

    if (!task) {
      const created = await createApprovalTask(context, run, node, node.config.approverRole, node.config.instructions);
      logs.push({ level: "info", message: `Approval task ${node.label} created.`, at: nowIso() });
      return {
        status: "PAUSED",
        state: { ...state, approvalTaskId: created.id, pauseReason: "approval" },
        logs,
        advance: false,
      };
    }

    if (task.data.status === "approved") {
      logs.push({ level: "info", message: `Approval task ${node.label} approved.`, at: nowIso() });
      return {
        status: "RUNNING",
        state: { ...state, cursor: state.cursor + 1, approvalTaskId: task.id, pauseReason: undefined },
        logs,
        advance: false,
      };
    }

    if (task.data.status === "rejected") {
      throw new Error(`Approval task ${node.label} was rejected.`);
    }

    logs.push({ level: "info", message: `Workflow waiting on approval for ${node.label}.`, at: nowIso() });
    return {
      status: "PAUSED",
      state: { ...state, approvalTaskId: task.id, pauseReason: "approval" },
      logs,
      advance: false,
    };
  }

  if (node.type === "subflow" && "workflowKey" in node.config) {
    const childWorkflowKey = node.config.workflowKey;
    const childWorkflow = manifest.workflows.find((candidate) => candidate.key === childWorkflowKey);
    if (!childWorkflow) {
      throw new Error(`Subflow ${node.label} references missing workflow ${childWorkflowKey}.`);
    }

    if (!state.subflowRunId) {
      const childRun = await queueWorkflowRun({
        tenantSlug: context.tenantSlug,
        environmentSlug: context.environmentSlug,
        workflowId: childWorkflow.id,
        payload: {
          parentWorkflowKey: workflow.key,
          parentNodeId: node.id,
          parentContext: state.context,
        },
        meta: {
          parentWorkflowRunId: run.id,
        },
      });
      logs.push({ level: "info", message: `Queued subflow ${childWorkflow.key} from ${node.label}.`, at: nowIso() });
      return {
        status: "PAUSED",
        state: {
          ...state,
          subflowRunId: childRun.id,
          pauseReason: "subflow",
        },
        logs,
        advance: false,
      };
    }

    const childRun = await getWorkflowRunById(context, state.subflowRunId, input.useDatabase);
    if (!childRun) {
      throw new Error(`Subflow run ${state.subflowRunId} could not be found.`);
    }

    if (childRun.status === "SUCCEEDED") {
      logs.push({ level: "info", message: `Subflow ${childWorkflow.key} completed.`, at: nowIso() });
      return {
        status: "RUNNING",
        state: {
          ...state,
          cursor: state.cursor + 1,
          pauseReason: undefined,
          subflowRunId: undefined,
          context: {
            ...state.context,
            [`${node.id}_subflowOutput`]: childRun.output ?? {},
          },
        },
        logs,
        advance: false,
      };
    }

    if (childRun.status === "FAILED") {
      throw new Error(`Subflow ${childWorkflow.key} failed.`);
    }

    logs.push({ level: "info", message: `Waiting for subflow ${childWorkflow.key}.`, at: nowIso() });
    return {
      status: "PAUSED",
      state: {
        ...state,
        pauseReason: "subflow",
      },
      logs,
      advance: false,
    };
  }

  if (node.type === "model_call" && "agentId" in node.config) {
    const agentId = node.config.agentId;
    const objectKey = node.config.objectKey;
    const records = objectKey
      ? await listSystemRecords({
          ...context,
          objectKey,
        })
      : [];
    const prepared = prepareAgentInvocation({
      manifest,
      agentId,
      objectKey: objectKey ?? manifest.agents.find((candidate) => candidate.id === agentId || candidate.key === agentId)?.objectKeys[0] ?? "",
      records: records.map((record) => ({
        id: record.id,
        objectKey: objectKey ?? "",
        data: record.data,
        createdAt: record.createdAt,
        updatedAt: record.updatedAt,
      })),
    });
    const prompt = typeof state.context.prompt === "string" ? state.context.prompt : node.label;
    const approvalTasks = await listApprovalTasksForRun(context, run.id);
    const agentApprovalTask = approvalTasks.find(
      (candidate) => candidate.data.nodeId === node.id && candidate.data.taskType === "agent_execution",
    );

    if (prepared.agent.approvalPolicy?.required && (!agentApprovalTask || agentApprovalTask.data.status === "pending")) {
      let agentRunId = state.pendingAgentRunId;
      if (!agentRunId) {
        const blockedRun = await recordAgentRun(context, {
          agentId: prepared.agent.id,
          agentKey: prepared.agent.key,
          status: "blocked",
          runMode: "runtime",
          approvalStatus: "pending",
          input: {
            workflowRunId: run.id,
            workflowKey: workflow.key,
            prompt,
            objectKey: objectKey ?? prepared.agent.objectKeys[0] ?? "",
            maskedRecords: prepared.inputRecords,
          },
          output: null,
          logs: [
            {
              level: "warning",
              message: `Execution blocked pending approval for ${prepared.agent.key}.`,
              at: nowIso(),
            },
          ],
          modelProviderKey: prepared.provider.key,
          costUsd: 0,
          tokensIn: 0,
          tokensOut: 0,
          trace: buildAgentTrace({
            agent: prepared.agent,
            provider: prepared.provider,
            policyDecisions: ["approval:required", `provider:${prepared.provider.key}`],
            stages: [
              {
                stage: "policy_preflight",
                status: "succeeded",
                summary: `Preflight passed for ${prepared.provider.name}.`,
                at: nowIso(),
              },
              {
                stage: "provider_execution",
                status: "blocked",
                summary: "Execution blocked pending approval.",
                at: nowIso(),
              },
            ] satisfies AgentTrace["stages"],
          }),
          completedAt: null,
        });
        agentRunId = blockedRun.id;
      }

      const approvalTask =
        agentApprovalTask ??
        (await createSystemRecord({
          ...context,
          objectKey: systemObjectKey("approvalTask"),
          data: {
            workflowRunId: run.id,
            workflowKey: workflow.key,
            nodeId: node.id,
            nodeLabel: node.label,
            taskType: "agent_execution",
            agentRunId,
            agentKey: prepared.agent.key,
            approverRole: prepared.agent.approvalPolicy.approverRole ?? "BUILDER_ADMIN",
            status: "pending",
            instructions: prepared.agent.approvalPolicy.notes ?? "Approve agent execution before provider invocation.",
          },
        }));

      logs.push({ level: "info", message: `Agent ${prepared.agent.key} blocked pending approval.`, at: nowIso() });
      return {
        status: "PAUSED",
        state: {
          ...state,
          approvalTaskId: approvalTask.id,
          pauseReason: "approval",
          pendingAgentRunId: agentRunId,
        },
        logs,
        advance: false,
      };
    }

    if (agentApprovalTask?.data.status === "rejected") {
      throw new Error(`Agent execution ${prepared.agent.key} was rejected.`);
    }

    const traceStages: AgentTrace["stages"] = [
      {
        stage: "policy_preflight",
        status: "succeeded",
        summary: `Preflight passed for ${prepared.provider.name}.`,
        at: nowIso(),
      },
      {
        stage: "prompt_assembly",
        status: "succeeded",
        summary: `Prepared ${prepared.agent.promptBlocks.length + 1} prompt blocks.`,
        at: nowIso(),
      },
    ];
    const policyDecisions = [
      `provider:${prepared.provider.key}`,
      `zeroRetention:${String(prepared.agent.zeroRetentionRequired)}`,
      `approval:${agentApprovalTask ? "approved" : "not_required"}`,
    ];

    let agentOutput;
    try {
      agentOutput = await executeAgentWithProvider({
        agent: prepared.agent,
        provider: prepared.provider,
        prompt,
        maskedRecords: prepared.inputRecords,
      });
      traceStages.push({
        stage: "provider_execution",
        status: "succeeded",
        summary: `Provider responded using ${prepared.provider.model}.`,
        at: nowIso(),
        meta: {
          tokensIn: agentOutput.tokensIn,
          tokensOut: agentOutput.tokensOut,
        },
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Agent execution failed.";
      traceStages.push({
        stage: "provider_execution",
        status: "failed",
        summary: message,
        at: nowIso(),
      });
      await recordAgentRun(
        context,
        {
          agentId: prepared.agent.id,
          agentKey: prepared.agent.key,
          status: "failed",
          runMode: "runtime",
          approvalStatus: agentApprovalTask ? "approved" : "not_required",
          input: {
            workflowRunId: run.id,
            workflowKey: workflow.key,
            prompt,
          },
          output: null,
          logs: [
            {
              level: "error",
              message,
              at: nowIso(),
            },
          ],
          modelProviderKey: prepared.provider.key,
          costUsd: 0,
          tokensIn: 0,
          tokensOut: 0,
          trace: buildAgentTrace({
            agent: prepared.agent,
            provider: prepared.provider,
            policyDecisions,
            stages: traceStages,
          }),
          parentWorkflowRunId: run.id,
          approvalTaskId: typeof agentApprovalTask?.id === "string" ? agentApprovalTask.id : null,
          completedAt: nowIso(),
        },
        state.pendingAgentRunId,
      );
      throw error instanceof Error ? error : new Error(message);
    }

    const costUsd = summarizeAgentRunCost(prepared.provider.model, agentOutput.tokensIn, agentOutput.tokensOut);
    const schemaValidation = validateAgentOutputSchema(prepared.agent.outputSchema, agentOutput.outputText);
    traceStages.push({
      stage: "output_validation",
      status: schemaValidation.passed ? "succeeded" : "failed",
      summary: schemaValidation.summary,
      at: nowIso(),
    });

    let handoffWorkflowRunId: string | null = null;
    if (schemaValidation.passed && prepared.agent.handoffWorkflowKeys.length > 0) {
      const handoffWorkflow = manifest.workflows.find((candidate) => prepared.agent.handoffWorkflowKeys.includes(candidate.key));
      if (handoffWorkflow) {
        const handoffRun = await queueWorkflowRun({
          tenantSlug: context.tenantSlug,
          environmentSlug: context.environmentSlug,
          workflowId: handoffWorkflow.id,
          payload: {
            parentAgentKey: prepared.agent.key,
            parentWorkflowRunId: run.id,
            outputSummary: agentOutput.outputText,
          },
          meta: {
            parentWorkflowRunId: run.id,
          },
        }).catch(() => null);
        handoffWorkflowRunId = handoffRun?.id ?? null;
        traceStages.push({
          stage: "workflow_handoff",
          status: handoffRun ? "succeeded" : "failed",
          summary: handoffRun ? `Queued handoff to ${handoffWorkflow.key}.` : `Failed to queue handoff to ${handoffWorkflow.key}.`,
          at: nowIso(),
        });
      }
    }

    traceStages.push({
      stage: "cost_evaluation",
      status:
        prepared.agent.costBudgetUsd != null && costUsd >= prepared.agent.costBudgetUsd ? "blocked" : "succeeded",
      summary:
        prepared.agent.costBudgetUsd != null && costUsd >= prepared.agent.costBudgetUsd
          ? `Budget threshold breached at ${costUsd.toFixed(4)} USD.`
          : `Cost ${costUsd.toFixed(4)} USD within budget.`,
      at: nowIso(),
    });
    traceStages.push({
      stage: "audit",
      status: "succeeded",
      summary: "Persisted runtime trace.",
      at: nowIso(),
    });

    await recordAgentRun(
      context,
      {
        agentId: prepared.agent.id,
        agentKey: prepared.agent.key,
        status: schemaValidation.passed ? "succeeded" : "failed",
        runMode: "runtime",
        approvalStatus: agentApprovalTask ? "approved" : "not_required",
        input: {
          workflowRunId: run.id,
          workflowKey: workflow.key,
          prompt,
        },
        output: {
          summary: agentOutput.outputText,
        },
        logs: [
          {
            level: schemaValidation.passed ? "info" : "error",
            message: schemaValidation.passed
              ? `Workflow ${workflow.key} invoked ${prepared.agent.key}.`
              : schemaValidation.summary,
            at: nowIso(),
          },
        ],
        modelProviderKey: prepared.provider.key,
        costUsd,
        tokensIn: agentOutput.tokensIn,
        tokensOut: agentOutput.tokensOut,
        trace: buildAgentTrace({
          agent: prepared.agent,
          provider: prepared.provider,
          policyDecisions,
          outputValidationPassed: schemaValidation.passed,
          handoffWorkflowKey: prepared.agent.handoffWorkflowKeys[0],
          stages: traceStages,
        }),
        parentWorkflowRunId: run.id,
        approvalTaskId: typeof agentApprovalTask?.id === "string" ? agentApprovalTask.id : null,
        handoffWorkflowRunId,
        outputValidationPassed: schemaValidation.passed,
        schemaValidation: {
          passed: schemaValidation.passed,
          summary: schemaValidation.summary,
        },
        completedAt: nowIso(),
      },
      state.pendingAgentRunId,
    );

    if (!schemaValidation.passed) {
      await recordRuntimeAlert(context, `Agent output validation failed for ${prepared.agent.key}`, schemaValidation.summary, run.id);
      throw new Error(schemaValidation.summary);
    }

    logs.push({ level: "info", message: `Model call ${node.label} completed.`, at: nowIso() });
    return {
      status: "RUNNING",
      state: {
        ...state,
        cursor: state.cursor + 1,
        pauseReason: undefined,
        approvalTaskId: undefined,
        pendingAgentRunId: undefined,
        context: {
          ...state.context,
          [`${node.id}_output`]: agentOutput.outputText,
        },
      },
      logs,
      advance: false,
    };
  }

  logs.push({ level: "info", message: `Visited ${node.type} node "${node.label}".`, at: nowIso() });
  return {
    status: "RUNNING",
    state: { ...state, cursor: state.cursor + 1 },
    logs,
    advance: false,
  };
}

async function processRun(input: {
  context: WorkflowRuntimeContext;
  run: WorkflowRunShape;
  workflow: WorkflowDefinition;
  manifest: Awaited<ReturnType<typeof getRuntimeManifest>>;
  useDatabase: boolean;
}): Promise<void> {
  const ordered = orderedNodes(input.workflow);
  let logs = [...((input.run.logs as Array<Record<string, unknown>> | null) ?? [])];
  let state = getRuntimeState({
    input: (input.run.input as Record<string, unknown> | null) ?? null,
    output: (input.run.output as Record<string, unknown> | null) ?? null,
  });
  state = {
    ...state,
    nextRetryAt: undefined,
  };

  if (logs.length === 0) {
    logs = appendLog(logs, "info", `Starting workflow ${input.workflow.name}.`);
  }

  const persist = async (status: WorkflowRunRecord["status"], finished = false) => {
    const output = withRuntimeState((input.run.output as Record<string, unknown> | null) ?? null, state, {
      completedNodes: Math.min(state.cursor, ordered.length),
      workflowKey: input.workflow.key,
    });
    if (input.useDatabase) {
      await persistDbWorkflowRun({
        runId: input.run.id,
        status,
        logs,
        output,
        startedAt: status === "RUNNING" && !input.run.startedAt,
        finishedAt: finished,
      });
      return;
    }

    await persistLocalWorkflowRun({
      runId: input.run.id,
      status,
      logs,
      output,
    });
  };

  await persist("RUNNING");

  while (state.cursor < ordered.length) {
    const node = ordered[state.cursor]!;
    try {
      const result = await executeWorkflowNode({
        context: input.context,
        manifest: input.manifest,
        workflow: input.workflow,
        node,
        ordered,
        run: input.run,
        state,
        logs,
        useDatabase: input.useDatabase,
      });
      logs = result.logs;
      state = result.state;
      if (result.status === "PAUSED") {
        await persist("PAUSED");
        return;
      }
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : `Failed at node ${node.label}.`;
      const retries = state.retryCounts[node.id] ?? 0;
      const nextRetryCount = retries + 1;
      state = {
        ...state,
        retryCounts: {
          ...state.retryCounts,
          [node.id]: nextRetryCount,
        },
        nextRetryAt: getExponentialRetryAt(nextRetryCount),
        pauseReason: undefined,
      };
      logs = appendLog(logs, "error", errorMessage, { nodeId: node.id });

      if (retries < 2) {
        await persist("QUEUED");
        await recordRuntimeAlert(input.context, `Retry scheduled for ${node.label}`, errorMessage, input.run.id);
        await enqueueWorkflowRun(input.run.id).catch(() => undefined);
        return;
      }

      await recordDeadLetter(input.context, input.run.id, input.workflow.key, errorMessage, {
        nodeId: node.id,
        nodeLabel: node.label,
        workflowKey: input.workflow.key,
      });
      await recordRuntimeAlert(input.context, `Workflow failed at ${node.label}`, errorMessage, input.run.id);
      await persist("FAILED", true);
      return;
    }
  }

  logs = appendLog(logs, "info", `Completed workflow ${input.workflow.name}.`);
  await persist("SUCCEEDED", true);
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

async function processLocalQueuedRuns(targetRunId?: string): Promise<number> {
  const runs = (await listLocalQueuedWorkflowRuns()).filter((run) => {
    if (targetRunId && run.id !== targetRunId) {
      return false;
    }
    const state = getRuntimeState({
      input: (run.input as Record<string, unknown> | null) ?? null,
      output: (run.output as Record<string, unknown> | null) ?? null,
    });
    return !state.nextRetryAt || new Date(state.nextRetryAt).getTime() <= Date.now();
  });
  let processed = 0;

  for (const run of runs) {
    const tenant = await getLocalTenantById(run.tenantId);
    const environment = await getLocalEnvironmentById(run.environmentId);
    if (!tenant || !environment) {
      continue;
    }

    await markDuePausedRunsQueued(
      {
        tenantId: tenant.id,
        environmentId: environment.id,
        tenantSlug: tenant.slug,
        environmentSlug: environment.slug,
      },
      false,
    );

    const manifest = await getRuntimeManifest({
      tenantSlug: tenant.slug,
      environmentSlug: environment.slug,
    });
    const workflow = manifest?.workflows.find((candidate) => candidate.id === run.workflowId || candidate.key === run.workflowKey);
    if (!manifest || !workflow) {
      await updateLocalWorkflowRun({
        runId: run.id,
        status: "FAILED",
        appendLog: {
          level: "error",
          message: "Workflow definition not found.",
          at: nowIso(),
        },
      });
      processed += 1;
      continue;
    }

    await processRun({
      context: {
        tenantId: tenant.id,
        environmentId: environment.id,
        tenantSlug: tenant.slug,
        environmentSlug: environment.slug,
      },
      run,
      workflow,
      manifest,
      useDatabase: false,
    });
    processed += 1;
  }

  return processed;
}

async function processLocalOutboxEvents(): Promise<{
  processed: number;
  deliveries: number;
  alerts: number;
}> {
  const tenants = await listLocalTenants();
  const totals = {
    processed: 0,
    deliveries: 0,
    alerts: 0,
  };

  for (const tenant of tenants) {
    const environments = await listLocalEnvironments(tenant.id);
    for (const environment of environments) {
      const result = await processPendingOutboxEvents({
        tenantSlug: tenant.slug,
        environmentSlug: environment.slug,
      });
      totals.processed += result.processed;
      totals.deliveries += result.deliveries;
      totals.alerts += result.alerts;
    }
  }

  return totals;
}

async function processDatabaseQueuedRuns(targetRunId?: string): Promise<number> {
  const prisma = getPrisma();
  const queuedRuns = await prisma.platformWorkflowRun.findMany({
    where: {
      status: "QUEUED",
      ...(targetRunId ? { id: targetRunId } : {}),
    },
    include: {
      tenant: true,
      environment: true,
    },
    orderBy: { createdAt: "asc" },
    take: 20,
  });
  const runs = queuedRuns.filter((run) => {
    const state = getRuntimeState({
      input: (run.input as Record<string, unknown> | null) ?? null,
      output: (run.output as Record<string, unknown> | null) ?? null,
    });
    return !state.nextRetryAt || new Date(state.nextRetryAt).getTime() <= Date.now();
  });

  let processed = 0;

  for (const run of runs) {
    await markDuePausedRunsQueued(
      {
        tenantId: run.tenantId,
        environmentId: run.environmentId,
        tenantSlug: run.tenant.slug,
        environmentSlug: run.environment.slug,
      },
      true,
    );

    const manifest = await getRuntimeManifest({
      tenantSlug: run.tenant.slug,
      environmentSlug: run.environment.slug,
    });
    const workflow = manifest?.workflows.find((candidate) => candidate.id === run.workflowId || candidate.key === run.workflowKey);
    if (!manifest || !workflow) {
      await prisma.platformWorkflowRun.update({
        where: { id: run.id },
        data: {
          status: "FAILED",
          logs: toJsonValue(appendLog((run.logs as Array<Record<string, unknown>> | null) ?? [], "error", "Workflow definition not found.")),
        },
      });
      processed += 1;
      continue;
    }

    await processRun({
      context: {
        tenantId: run.tenantId,
        environmentId: run.environmentId,
        tenantSlug: run.tenant.slug,
        environmentSlug: run.environment.slug,
      },
      run: {
        ...run,
        workflowId: run.workflowId,
        workflowKey: run.workflowKey,
        status: run.status,
        input: (run.input as Record<string, unknown> | null) ?? null,
        output: (run.output as Record<string, unknown> | null) ?? null,
        logs: (run.logs as Array<Record<string, unknown>> | null) ?? [],
        startedAt: run.startedAt?.toISOString() ?? null,
        finishedAt: run.finishedAt?.toISOString() ?? null,
        createdAt: run.createdAt.toISOString(),
        updatedAt: run.updatedAt.toISOString(),
      },
      workflow,
      manifest,
      useDatabase: true,
    });
    processed += 1;
  }

  return processed;
}

export async function runWorkflowWorkerCycle(targetRunId?: string): Promise<number> {
  return withLocalFallback(() => processDatabaseQueuedRuns(targetRunId), () => processLocalQueuedRuns(targetRunId));
}

export async function startPlatformWorker(intervalMs = 10_000): Promise<void> {
  console.log(`[platform-worker] started with interval ${intervalMs}ms`);

  const queueWorker = createWorkflowRunWorker(async (runId) => {
    const processed = await runWorkflowWorkerCycle(runId);
    console.log(`[platform-worker] processed ${processed} queued run(s) from execution bus`);
  });
  const outboxWorker = createOutboxWorker(async ({ tenantSlug, environmentSlug, eventId }) => {
    const processed = await processPendingOutboxEvents({
      tenantSlug,
      environmentSlug,
      eventId,
    });
    console.log(
      `[platform-worker] processed ${processed.processed} outbox event(s), ${processed.deliveries} deliveries, ${processed.alerts} alerts`,
    );
  });

  if (queueWorker) {
    console.log("[platform-worker] BullMQ worker active");
  }
  if (outboxWorker) {
    console.log("[platform-worker] Outbox worker active");
  }

  while (true) {
    try {
      const processed = await runWorkflowWorkerCycle();
      if (processed > 0) {
        console.log(`[platform-worker] processed ${processed} queued run(s)`);
      }

      if (!outboxWorker && getEnv().PLATFORM_LOCAL_DEV_MODE) {
        const outboxResults = await processLocalOutboxEvents();
        if (outboxResults.processed > 0 || outboxResults.deliveries > 0 || outboxResults.alerts > 0) {
          console.log(
            `[platform-worker] processed ${outboxResults.processed} local outbox event(s), ${outboxResults.deliveries} deliveries, ${outboxResults.alerts} alerts`,
          );
        }
      }
    } catch (error) {
      console.error("[platform-worker] cycle failed", error);
    }

    await new Promise((resolve) => setTimeout(resolve, intervalMs));
  }
}
