import { Queue, Worker, type ConnectionOptions } from "bullmq";

import { getEnv } from "@/lib/env";

export interface WorkflowRunJobPayload {
  type: "workflow.run";
  runId: string;
}

export interface OutboxEventJobPayload {
  type: "outbox.event";
  tenantSlug: string;
  environmentSlug: string;
  eventId?: string;
}

const WORKFLOW_QUEUE_NAME = "ta-platform-workflow-runs";
const OUTBOX_QUEUE_NAME = "ta-platform-outbox-events";

let cachedConnection: ConnectionOptions | null = null;
let cachedWorkflowQueue: Queue | null = null;
let cachedOutboxQueue: Queue | null = null;

function getRedisConnection(): ConnectionOptions | null {
  const env = getEnv();
  if (!env.REDIS_URL) {
    return null;
  }

  if (cachedConnection) {
    return cachedConnection;
  }

  cachedConnection = {
    url: env.REDIS_URL,
    maxRetriesPerRequest: null,
  };
  return cachedConnection;
}

export function hasExecutionQueue(): boolean {
  return Boolean(getRedisConnection());
}

function getWorkflowQueue(): Queue | null {
  const connection = getRedisConnection();
  if (!connection) {
    return null;
  }

  if (cachedWorkflowQueue) {
    return cachedWorkflowQueue;
  }

  cachedWorkflowQueue = new Queue(WORKFLOW_QUEUE_NAME, { connection });
  return cachedWorkflowQueue;
}

function getOutboxQueue(): Queue | null {
  const connection = getRedisConnection();
  if (!connection) {
    return null;
  }

  if (cachedOutboxQueue) {
    return cachedOutboxQueue;
  }

  cachedOutboxQueue = new Queue(OUTBOX_QUEUE_NAME, { connection });
  return cachedOutboxQueue;
}

export async function enqueueWorkflowRun(runId: string): Promise<boolean> {
  const queue = getWorkflowQueue();
  if (!queue) {
    return false;
  }

  await queue.add(
    "workflow.run",
    {
      type: "workflow.run",
      runId,
    },
    {
      jobId: `workflow-run:${runId}`,
      removeOnComplete: 100,
      removeOnFail: 100,
    },
  );

  return true;
}

export async function enqueueOutboxEvent(input: {
  tenantSlug: string;
  environmentSlug: string;
  eventId?: string;
}): Promise<boolean> {
  const queue = getOutboxQueue();
  if (!queue) {
    return false;
  }

  await queue.add(
    "outbox.event",
    {
      type: "outbox.event",
      tenantSlug: input.tenantSlug,
      environmentSlug: input.environmentSlug,
      eventId: input.eventId,
    },
    {
      jobId: input.eventId
        ? `outbox-event:${input.eventId}`
        : `outbox-cycle:${input.tenantSlug}:${input.environmentSlug}`,
      removeOnComplete: 100,
      removeOnFail: 100,
    },
  );

  return true;
}

export function createWorkflowRunWorker(
  processor: (runId: string) => Promise<void>,
): Worker | null {
  const connection = getRedisConnection();
  if (!connection) {
    return null;
  }

  return new Worker<WorkflowRunJobPayload>(
    WORKFLOW_QUEUE_NAME,
    async (job) => {
      if (job.data.type !== "workflow.run") {
        return;
      }

      await processor(job.data.runId);
    },
    {
      connection,
    },
  );
}

export function createOutboxWorker(
  processor: (payload: { tenantSlug: string; environmentSlug: string; eventId?: string }) => Promise<void>,
): Worker | null {
  const connection = getRedisConnection();
  if (!connection) {
    return null;
  }

  return new Worker<OutboxEventJobPayload>(
    OUTBOX_QUEUE_NAME,
    async (job) => {
      if (job.data.type !== "outbox.event") {
        return;
      }

      await processor({
        tenantSlug: job.data.tenantSlug,
        environmentSlug: job.data.environmentSlug,
        eventId: job.data.eventId,
      });
    },
    {
      connection,
    },
  );
}
