import nodemailer from "nodemailer";

import { getEnv } from "@/lib/env";
import type {
  AgentDefinition,
  EventEnvelope,
  ModelProviderDefinition,
  NotificationChannelDefinition,
  NotificationSeverity,
} from "@/lib/platform/types";

interface AgentExecutionResult {
  outputText: string;
  tokensIn: number;
  tokensOut: number;
  raw: Record<string, unknown>;
}

interface NotificationDeliveryResult {
  deliveredAt: string;
  destination?: string;
  provider: string;
  responseSummary?: string;
}

let cachedTransporter: nodemailer.Transporter | null = null;

function trimTrailingSlash(value: string): string {
  return value.endsWith("/") ? value.slice(0, -1) : value;
}

export function resolveSecretReference(secretRef: string): string | null {
  if (!secretRef.trim()) {
    return null;
  }

  if (secretRef.startsWith("env:")) {
    const envKey = secretRef.slice(4).trim();
    return envKey ? process.env[envKey] ?? null : null;
  }

  return process.env[secretRef] ?? null;
}

function getSmtpTransporter(): nodemailer.Transporter | null {
  const env = getEnv();
  if (!env.SMTP_URL) {
    return null;
  }

  if (cachedTransporter) {
    return cachedTransporter;
  }

  cachedTransporter = nodemailer.createTransport(env.SMTP_URL);
  return cachedTransporter;
}

function estimateCostUsd(model: string, tokensIn: number, tokensOut: number): number {
  const total = tokensIn + tokensOut;
  const multiplier = model.includes("gpt-5") ? 1.15 : 1;
  return Number((((total / 1000) * 0.0085) * multiplier).toFixed(4));
}

function estimateFallbackTokens(text: string): number {
  return Math.max(48, Math.ceil(text.length / 3.6));
}

function extractCompletionText(payload: Record<string, unknown>): string {
  const choices = Array.isArray(payload.choices) ? payload.choices : [];
  const firstChoice = choices[0] as Record<string, unknown> | undefined;
  const message = firstChoice?.message as Record<string, unknown> | undefined;
  const content = message?.content;

  if (typeof content === "string" && content.trim()) {
    return content.trim();
  }

  if (Array.isArray(content)) {
    const text = content
      .map((entry) => (typeof entry === "object" && entry && "text" in entry ? String((entry as { text?: unknown }).text ?? "") : ""))
      .join("\n")
      .trim();
    if (text) {
      return text;
    }
  }

  return "The model returned no textual response.";
}

export async function executeAgentWithProvider(input: {
  agent: AgentDefinition;
  provider: ModelProviderDefinition;
  prompt: string;
  maskedRecords: Array<Record<string, unknown>>;
}): Promise<AgentExecutionResult> {
  if (input.provider.provider !== "openai" && input.provider.provider !== "azure_openai") {
    throw new Error(`Provider ${input.provider.provider} is not supported in this tranche.`);
  }

  const apiKey = resolveSecretReference(input.provider.apiKeySecretRef);
  if (!apiKey) {
    throw new Error(`Missing provider secret for ${input.provider.name}.`);
  }

  const endpoint = trimTrailingSlash(input.provider.endpoint || "https://api.openai.com/v1");
  const env = getEnv();
  const isAzure = input.provider.provider === "azure_openai";
  const azureApiVersion = env.AZURE_OPENAI_API_VERSION;
  const requestUrl = isAzure
    ? endpoint.includes("/openai/deployments/")
      ? `${endpoint}/chat/completions${endpoint.includes("?") ? "&" : "?"}api-version=${encodeURIComponent(azureApiVersion)}`
      : `${endpoint}/openai/deployments/${encodeURIComponent(input.provider.model)}/chat/completions?api-version=${encodeURIComponent(azureApiVersion)}`
    : `${endpoint}/chat/completions`;
  const response = await fetch(requestUrl, {
    method: "POST",
    headers: {
      ...(isAzure ? { "api-key": apiKey } : { authorization: `Bearer ${apiKey}` }),
      "content-type": "application/json",
    },
    body: JSON.stringify({
      ...(isAzure ? {} : { model: input.provider.model }),
      reasoning_effort: !isAzure && input.provider.model.includes("gpt-5") ? "low" : undefined,
      messages: [
        {
          role: "system",
          content: [
            input.agent.prompt,
            ...input.agent.promptBlocks.map((block) => `${block.kind.toUpperCase()}: ${block.label}\n${block.content}`),
            input.agent.outputSchema ? `OUTPUT_SCHEMA:\n${input.agent.outputSchema}` : "",
            input.agent.zeroRetentionRequired ? "POLICY: Zero retention is required for this invocation." : "",
          ]
            .filter(Boolean)
            .join("\n\n"),
        },
        {
          role: "user",
          content: JSON.stringify(
            {
              prompt: input.prompt,
              records: input.maskedRecords,
              objectScope: input.agent.objectKeys,
              allowedToolIds: input.agent.allowedToolIds,
              handoffWorkflowKeys: input.agent.handoffWorkflowKeys,
            },
            null,
            2,
          ),
        },
      ],
    }),
  });

  const payload = (await response.json().catch(() => ({}))) as Record<string, unknown>;
  if (!response.ok) {
    throw new Error(String(payload.error && typeof payload.error === "object" ? (payload.error as { message?: unknown }).message ?? "Model request failed." : "Model request failed."));
  }

  const outputText = extractCompletionText(payload);
  const usage = (payload.usage as Record<string, unknown> | undefined) ?? {};
  const promptTokens = typeof usage.prompt_tokens === "number" ? usage.prompt_tokens : estimateFallbackTokens(input.prompt);
  const completionTokens = typeof usage.completion_tokens === "number" ? usage.completion_tokens : estimateFallbackTokens(outputText);

  return {
    outputText,
    tokensIn: promptTokens,
    tokensOut: completionTokens,
    raw: payload,
  };
}

export function summarizeAgentRunCost(model: string, tokensIn: number, tokensOut: number): number {
  return estimateCostUsd(model, tokensIn, tokensOut);
}

export async function deliverNotification(input: {
  channel: NotificationChannelDefinition;
  severity: NotificationSeverity;
  subject?: string;
  body: string;
  event: EventEnvelope;
  supportEmail?: string;
}): Promise<NotificationDeliveryResult> {
  const destination = input.channel.destination?.trim() || undefined;

  if (input.channel.kind === "in_app") {
    return {
      deliveredAt: new Date().toISOString(),
      destination,
      provider: "in_app",
      responseSummary: "Stored in tenant inbox.",
    };
  }

  if (!destination) {
    throw new Error(`Notification channel ${input.channel.key} has no destination configured.`);
  }

  if (input.channel.kind === "webhook") {
    const response = await fetch(destination, {
      method: "POST",
      headers: {
        "content-type": "application/json",
      },
      body: JSON.stringify({
        event: input.event,
        severity: input.severity,
        subject: input.subject,
        body: input.body,
      }),
    });

    if (!response.ok) {
      throw new Error(`Webhook delivery failed with status ${response.status}.`);
    }

    return {
      deliveredAt: new Date().toISOString(),
      destination,
      provider: "webhook",
      responseSummary: `HTTP ${response.status}`,
    };
  }

  if (input.channel.kind === "slack_style") {
    const response = await fetch(destination, {
      method: "POST",
      headers: {
        "content-type": "application/json",
      },
      body: JSON.stringify({
        text: input.subject ? `${input.subject}\n${input.body}` : input.body,
        severity: input.severity,
        eventType: input.event.type,
      }),
    });

    if (!response.ok) {
      throw new Error(`Slack-style delivery failed with status ${response.status}.`);
    }

    return {
      deliveredAt: new Date().toISOString(),
      destination,
      provider: "slack_style",
      responseSummary: `HTTP ${response.status}`,
    };
  }

  const transporter = getSmtpTransporter();
  if (!transporter) {
    if (getEnv().PLATFORM_LOCAL_DEV_MODE) {
      return {
        deliveredAt: new Date().toISOString(),
        destination,
        provider: "email_simulated",
        responseSummary: "Simulated local SMTP delivery.",
      };
    }
    throw new Error("SMTP_URL is required for email delivery.");
  }

  const from = getEnv().SMTP_FROM || input.supportEmail || "platform@local.test";
  await transporter.sendMail({
    from,
    to: destination,
    subject: input.subject || `Platform ${input.severity} notification`,
    text: input.body,
  });

  return {
    deliveredAt: new Date().toISOString(),
    destination,
    provider: "email",
    responseSummary: "SMTP accepted message.",
  };
}
