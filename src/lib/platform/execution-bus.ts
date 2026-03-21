import { Queue, Worker, type ConnectionOptions } from "bullmq";

import { getEnv } from "@/lib/env";

export interface WorkflowRunJobPayload {
  type: "workflow.run";
  runId: string;
}

const WORKFLOW_QUEUE_NAME = "ta-platform-workflow-runs";

let cachedConnection: ConnectionOptions | null = null;
let cachedQueue: Queue | null = null;

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

  if (cachedQueue) {
    return cachedQueue;
  }

  cachedQueue = new Queue(WORKFLOW_QUEUE_NAME, { connection });
  return cachedQueue;
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
