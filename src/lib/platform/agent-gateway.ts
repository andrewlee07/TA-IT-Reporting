import { findObjectDefinition } from "@/lib/platform/manifest";
import { maskRecordForAgent } from "@/lib/platform/records";
import type { AgentDefinition, ModelProviderDefinition, PlatformManifest, PlatformRecord } from "@/lib/platform/types";

export interface PreparedAgentInvocation {
  agent: AgentDefinition;
  provider: ModelProviderDefinition;
  inputRecords: Array<Record<string, unknown>>;
  metadata: {
    masked: boolean;
    zeroRetentionRequired: boolean;
  };
}

export function prepareAgentInvocation(input: {
  manifest: PlatformManifest;
  agentId: string;
  objectKey: string;
  records: PlatformRecord[];
}): PreparedAgentInvocation {
  const agent = input.manifest.agents.find((candidate) => candidate.id === input.agentId || candidate.key === input.agentId);
  if (!agent) {
    throw new Error("Agent definition not found.");
  }

  const provider = input.manifest.modelProviders.find((candidate) => candidate.id === agent.modelProviderId || candidate.key === agent.modelProviderId);
  if (!provider) {
    throw new Error("Model provider not found.");
  }

  if (!input.manifest.securityPolicy.allowedModelProviderKeys.includes(provider.key)) {
    throw new Error("Model provider is not allowed by tenant security policy.");
  }

  if (agent.zeroRetentionRequired && !provider.supportsZeroRetention) {
    throw new Error("Selected model provider does not support zero retention.");
  }

  const objectDefinition = findObjectDefinition(input.manifest, input.objectKey);
  if (!objectDefinition) {
    throw new Error("Object definition not found for agent invocation.");
  }

  return {
    agent,
    provider,
    inputRecords: input.records.map((record) =>
      maskRecordForAgent({
        manifest: input.manifest,
        objectKey: input.objectKey,
        record,
      }),
    ),
    metadata: {
      masked: true,
      zeroRetentionRequired: agent.zeroRetentionRequired,
    },
  };
}
