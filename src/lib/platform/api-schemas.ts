import { z } from "zod";

export const nonEmptyStringSchema = z.string().trim().min(1);
export const optionalNonEmptyStringSchema = z.string().trim().min(1).optional();

export const recordSchema = z.object({
  data: z.record(z.string(), z.unknown()),
});

export const objectSchema = z.object({
  id: z.string().optional(),
  key: optionalNonEmptyStringSchema,
  label: nonEmptyStringSchema,
  pluralLabel: nonEmptyStringSchema,
  description: optionalNonEmptyStringSchema,
  icon: optionalNonEmptyStringSchema,
  primaryFieldKey: optionalNonEmptyStringSchema,
  allowCreate: z.boolean().optional(),
  allowUpdate: z.boolean().optional(),
  allowDelete: z.boolean().optional(),
});

export const fieldSchema = z.object({
  objectId: nonEmptyStringSchema,
  field: z.object({
    id: z.string().optional(),
    key: optionalNonEmptyStringSchema,
    label: nonEmptyStringSchema,
    type: z.enum(["text", "long_text", "number", "currency", "boolean", "date", "datetime", "select", "relationship", "computed"]),
    description: optionalNonEmptyStringSchema,
    required: z.boolean().optional(),
    unique: z.boolean().optional(),
    sensitivity: z.enum(["public", "internal", "pii", "sensitive"]).optional(),
    placeholder: optionalNonEmptyStringSchema,
    options: z.array(nonEmptyStringSchema).optional(),
    defaultValue: z.union([z.string(), z.number(), z.boolean(), z.null()]).optional(),
    validations: z
      .array(
        z.object({
          id: nonEmptyStringSchema,
          type: z.enum(["required", "unique", "min", "max", "regex"]),
          message: nonEmptyStringSchema,
          value: z.union([z.string(), z.number(), z.boolean()]).optional(),
        }),
      )
      .optional(),
    calculation: z
      .object({
        id: nonEmptyStringSchema,
        expression: nonEmptyStringSchema,
        outputType: z.enum(["text", "long_text", "number", "currency", "boolean", "date", "datetime", "select", "computed"]),
        description: optionalNonEmptyStringSchema,
      })
      .nullable()
      .optional(),
  }),
});

export const pageSchema = z.object({
  id: z.string().optional(),
  key: optionalNonEmptyStringSchema,
  title: nonEmptyStringSchema,
  route: optionalNonEmptyStringSchema,
  description: optionalNonEmptyStringSchema,
  layoutKey: optionalNonEmptyStringSchema,
  objectKey: optionalNonEmptyStringSchema,
  isHome: z.boolean().optional(),
  previewNote: optionalNonEmptyStringSchema,
});

export const menuSchema = z.object({
  id: z.string().optional(),
  key: optionalNonEmptyStringSchema,
  label: nonEmptyStringSchema,
  icon: optionalNonEmptyStringSchema,
  pageKey: nonEmptyStringSchema,
  order: z.number().optional(),
  group: optionalNonEmptyStringSchema,
});

const visibilityRuleSchema = z.object({
  expression: nonEmptyStringSchema,
  summary: optionalNonEmptyStringSchema,
});

const layoutPlacementSchema = z.object({
  zone: z.enum(["header", "main", "rail", "footer", "drawer"]),
  span: z.number().int().min(1).max(12),
  stackDirection: z.enum(["row", "column"]),
  variant: z.enum(["standard", "full_width", "rail", "tabs", "drawer"]),
  region: nonEmptyStringSchema,
  spacing: z.enum(["tight", "comfortable", "relaxed"]),
  alignment: z.enum(["start", "center", "between"]),
  minHeight: z.number().int().min(0),
  responsive: z.object({
    mobileSpan: z.number().int().min(1).max(12).optional(),
    tabletSpan: z.number().int().min(1).max(12).optional(),
    desktopSpan: z.number().int().min(1).max(12).optional(),
    hiddenOn: z.array(z.enum(["mobile", "tablet", "desktop"])).optional(),
  }),
});

const layoutComponentSchema = z.object({
  id: nonEmptyStringSchema,
  kind: z.enum([
    "hero",
    "text",
    "rich_text",
    "stat_tiles",
    "record_table",
    "record_form",
    "stats",
    "related_records",
    "workflow_launcher",
    "agent_summary",
    "agent_panel",
    "activity_feed",
    "callout",
  ]),
  title: nonEmptyStringSchema,
  description: optionalNonEmptyStringSchema,
  objectKey: optionalNonEmptyStringSchema,
  workflowKey: optionalNonEmptyStringSchema,
  agentId: optionalNonEmptyStringSchema,
  relatedObjectKey: optionalNonEmptyStringSchema,
  width: z.number().int().min(1).max(12),
  stylePreset: optionalNonEmptyStringSchema,
  placement: layoutPlacementSchema,
  visibilityRule: visibilityRuleSchema.optional(),
  binding: z
    .object({
      objectKey: optionalNonEmptyStringSchema,
      workflowKey: optionalNonEmptyStringSchema,
      agentId: optionalNonEmptyStringSchema,
      relatedObjectKey: optionalNonEmptyStringSchema,
      viewKey: optionalNonEmptyStringSchema,
      promptAsset: optionalNonEmptyStringSchema,
    })
    .optional(),
  props: z.record(z.string(), z.unknown()),
});

export const layoutSchema = z.object({
  id: nonEmptyStringSchema,
  key: nonEmptyStringSchema,
  name: nonEmptyStringSchema,
  pageKey: nonEmptyStringSchema,
  mobileColumns: z.number().int().min(1),
  tabletColumns: z.number().int().min(1),
  desktopColumns: z.number().int().min(1),
  sections: z.array(
    z.object({
      id: nonEmptyStringSchema,
      title: nonEmptyStringSchema,
      description: optionalNonEmptyStringSchema,
      kind: z.enum(["grid", "tabs", "drawer"]),
      columns: z.number().int().min(1),
      templateKey: optionalNonEmptyStringSchema,
      placement: layoutPlacementSchema,
      visibilityRule: visibilityRuleSchema.optional(),
      components: z.array(layoutComponentSchema),
    }),
  ),
});

const workflowTriggerSchema = z.discriminatedUnion("type", [
  z.object({
    id: nonEmptyStringSchema,
    type: z.literal("manual"),
    label: nonEmptyStringSchema,
    config: z.object({
      notes: optionalNonEmptyStringSchema,
    }),
  }),
  z.object({
    id: nonEmptyStringSchema,
    type: z.literal("record_created"),
    label: nonEmptyStringSchema,
    config: z.object({
      objectKey: nonEmptyStringSchema,
    }),
  }),
  z.object({
    id: nonEmptyStringSchema,
    type: z.literal("record_updated"),
    label: nonEmptyStringSchema,
    config: z.object({
      objectKey: nonEmptyStringSchema,
    }),
  }),
]);

const workflowNodeSchema = z.discriminatedUnion("type", [
  z.object({
    id: nonEmptyStringSchema,
    type: z.literal("condition"),
    label: nonEmptyStringSchema,
    config: z.object({
      expression: nonEmptyStringSchema,
    }),
    position: z.object({ x: z.number(), y: z.number() }),
  }),
  z.object({
    id: nonEmptyStringSchema,
    type: z.literal("crud"),
    label: nonEmptyStringSchema,
    config: z.object({
      operation: z.enum(["create", "update", "delete"]),
      objectKey: nonEmptyStringSchema,
      targetFieldKey: optionalNonEmptyStringSchema,
      valueExpression: optionalNonEmptyStringSchema,
    }),
    position: z.object({ x: z.number(), y: z.number() }),
  }),
  z.object({
    id: nonEmptyStringSchema,
    type: z.literal("formula"),
    label: nonEmptyStringSchema,
    config: z.object({
      expression: nonEmptyStringSchema,
      outputKey: nonEmptyStringSchema,
    }),
    position: z.object({ x: z.number(), y: z.number() }),
  }),
  z.object({
    id: nonEmptyStringSchema,
    type: z.literal("webhook"),
    label: nonEmptyStringSchema,
    config: z.object({
      method: z.enum(["GET", "POST", "PUT", "PATCH"]),
      url: nonEmptyStringSchema,
      bodyTemplate: optionalNonEmptyStringSchema,
    }),
    position: z.object({ x: z.number(), y: z.number() }),
  }),
  z.object({
    id: nonEmptyStringSchema,
    type: z.literal("notification"),
    label: nonEmptyStringSchema,
    config: z.object({
      channel: z.enum(["email", "slack", "task"]),
      recipient: optionalNonEmptyStringSchema,
      message: optionalNonEmptyStringSchema,
    }),
    position: z.object({ x: z.number(), y: z.number() }),
  }),
  z.object({
    id: nonEmptyStringSchema,
    type: z.literal("wait"),
    label: nonEmptyStringSchema,
    config: z.object({
      durationMinutes: z.number().int().min(1),
    }),
    position: z.object({ x: z.number(), y: z.number() }),
  }),
  z.object({
    id: nonEmptyStringSchema,
    type: z.literal("approval"),
    label: nonEmptyStringSchema,
    config: z.object({
      approverRole: z.enum(["SUPER_ADMIN", "BUILDER_ADMIN", "USER"]),
      instructions: optionalNonEmptyStringSchema,
    }),
    position: z.object({ x: z.number(), y: z.number() }),
  }),
  z.object({
    id: nonEmptyStringSchema,
    type: z.literal("model_call"),
    label: nonEmptyStringSchema,
    config: z.object({
      agentId: nonEmptyStringSchema,
      objectKey: optionalNonEmptyStringSchema,
      promptAsset: optionalNonEmptyStringSchema,
    }),
    position: z.object({ x: z.number(), y: z.number() }),
  }),
]);

export const workflowSchema = z.object({
  id: z.string().optional(),
  key: optionalNonEmptyStringSchema,
  name: nonEmptyStringSchema,
  description: optionalNonEmptyStringSchema,
  objectKey: optionalNonEmptyStringSchema,
  status: z.enum(["draft", "active"]).optional(),
  triggers: z.array(workflowTriggerSchema).optional(),
  nodes: z.array(workflowNodeSchema).optional(),
  edges: z
    .array(
      z.object({
        id: nonEmptyStringSchema,
        sourceId: nonEmptyStringSchema,
        targetId: nonEmptyStringSchema,
        label: optionalNonEmptyStringSchema,
      }),
    )
    .optional(),
});

export const agentSchema = z.object({
  id: z.string().optional(),
  key: optionalNonEmptyStringSchema,
  name: nonEmptyStringSchema,
  description: optionalNonEmptyStringSchema,
  scope: z.enum(["node", "workspace"]),
  modelProviderId: nonEmptyStringSchema,
  prompt: nonEmptyStringSchema,
  allowedToolIds: z.array(nonEmptyStringSchema).optional(),
  objectKeys: z.array(nonEmptyStringSchema).optional(),
  zeroRetentionRequired: z.boolean().optional(),
});

export const providerSchema = z.object({
  id: nonEmptyStringSchema,
  key: nonEmptyStringSchema,
  name: nonEmptyStringSchema,
  provider: z.enum(["openai", "azure_openai", "anthropic", "google", "custom"]),
  model: nonEmptyStringSchema,
  endpoint: optionalNonEmptyStringSchema,
  apiKeySecretRef: nonEmptyStringSchema,
  supportsZeroRetention: z.boolean(),
  allowedForSensitiveData: z.boolean(),
  status: z.enum(["active", "disabled"]),
});

export const inviteSchema = z.object({
  email: z.string().trim().email(),
  role: z.enum(["SUPER_ADMIN", "BUILDER_ADMIN", "USER"]),
  expiresInDays: z.number().int().min(1).max(30).optional(),
});

export const acceptInviteSessionSchema = z.object({
  token: nonEmptyStringSchema,
  email: z.string().trim().email(),
  name: nonEmptyStringSchema,
});

export const securityPolicySchema = z.object({
  zeroRetentionRequiredForSensitiveData: z.boolean(),
  defaultMaskingPolicyKey: nonEmptyStringSchema,
  allowedModelProviderKeys: z.array(nonEmptyStringSchema),
});
