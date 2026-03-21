export const PLATFORM_SCHEMA_VERSION = 2 as const;

export type PlatformRole = "SUPER_ADMIN" | "BUILDER_ADMIN" | "USER";
export type PlatformVersionStatus = "DRAFT" | "ACTIVE" | "ROLLED_BACK";
export type WorkflowRunStatus = "QUEUED" | "RUNNING" | "SUCCEEDED" | "FAILED" | "PAUSED";
export type RuleExpressionMode = "text" | "json_logic";

export type FieldType =
  | "text"
  | "long_text"
  | "number"
  | "currency"
  | "boolean"
  | "date"
  | "datetime"
  | "select"
  | "relationship"
  | "computed";

export type FieldSensitivity = "public" | "internal" | "pii" | "sensitive";

export type ValidationRuleType = "required" | "unique" | "min" | "max" | "regex";

export type LayoutComponentKind =
  | "hero"
  | "text"
  | "rich_text"
  | "stat_tiles"
  | "record_table"
  | "record_form"
  | "stats"
  | "related_records"
  | "workflow_launcher"
  | "agent_summary"
  | "agent_panel"
  | "activity_feed"
  | "callout";

export type LayoutSectionKind = "grid" | "tabs" | "drawer";
export type LayoutZone = "header" | "main" | "rail" | "footer" | "drawer";
export type LayoutStackDirection = "row" | "column";
export type LayoutVariant = "standard" | "full_width" | "rail" | "tabs" | "drawer";
export type LayoutSpacingPreset = "tight" | "comfortable" | "relaxed";
export type LayoutAlignmentPreset = "start" | "center" | "between";
export type ResponsiveBreakpoint = "mobile" | "tablet" | "desktop";

export type WorkflowNodeType =
  | "condition"
  | "crud"
  | "formula"
  | "webhook"
  | "notification"
  | "wait"
  | "approval"
  | "model_call";

export type TriggerType = "manual" | "record_created" | "record_updated";
export type AgentScope = "node" | "workspace";
export type MaskingMode = "mask" | "anonymize" | "block";
export type ModelProviderKind = "openai" | "azure_openai" | "anthropic" | "google" | "custom";
export type ModelProviderStatus = "active" | "disabled";
export type WorkflowCrudOperation = "create" | "update" | "delete";
export type WorkflowNotificationChannel = "email" | "slack" | "task";
export type WorkflowWebhookMethod = "GET" | "POST" | "PUT" | "PATCH";
export type FormDeliveryMode = "public" | "embedded" | "authenticated";
export type FormSubmissionStatus = "draft" | "submitted";
export type BrandingMode = "draft" | "review" | "approved";
export type BrandAssetKind = "logo" | "icon" | "brand_book" | "reference";

export interface PlatformActor {
  email: string;
  name: string;
  role: PlatformRole;
}

export interface RuleExpressionDefinition {
  mode: RuleExpressionMode;
  summary?: string;
  expression?: string;
  jsonLogic?: Record<string, unknown>;
}

export interface ValidationRuleDefinition {
  id: string;
  type: ValidationRuleType;
  message: string;
  value?: string | number | boolean;
  rule?: RuleExpressionDefinition;
}

export interface CalculationRuleDefinition {
  id: string;
  expression: string;
  outputType: Exclude<FieldType, "relationship">;
  description?: string;
  rule?: RuleExpressionDefinition;
}

export interface RelationshipDefinition {
  id: string;
  key: string;
  label: string;
  sourceObjectKey: string;
  targetObjectKey: string;
  type: "one_to_many" | "many_to_one";
}

export interface ViewDefinition {
  id: string;
  key: string;
  label: string;
  objectKey: string;
  visibleFieldKeys: string[];
  defaultSortKey?: string;
}

export interface FieldDefinition {
  id: string;
  key: string;
  label: string;
  type: FieldType;
  description?: string;
  helpText?: string;
  tooltip?: string;
  fieldGroup?: string;
  required: boolean;
  unique: boolean;
  sensitivity: FieldSensitivity;
  placeholder?: string;
  options?: string[];
  defaultValue?: string | number | boolean | null;
  mandatoryRule?: RuleExpressionDefinition;
  advancedValidation?: RuleExpressionDefinition;
  validations: ValidationRuleDefinition[];
  calculation?: CalculationRuleDefinition | null;
}

export interface ObjectDefinition {
  id: string;
  key: string;
  label: string;
  pluralLabel: string;
  description?: string;
  icon: string;
  primaryFieldKey: string;
  allowCreate: boolean;
  allowUpdate: boolean;
  allowDelete: boolean;
  fields: FieldDefinition[];
  relationships: RelationshipDefinition[];
  views: ViewDefinition[];
}

export interface VisibilityRuleDefinition {
  expression: string;
  summary?: string;
  rule?: RuleExpressionDefinition;
}

export interface LayoutResponsiveDefinition {
  mobileSpan?: number;
  tabletSpan?: number;
  desktopSpan?: number;
  hiddenOn?: ResponsiveBreakpoint[];
}

export interface LayoutPlacementDefinition {
  zone: LayoutZone;
  span: number;
  stackDirection: LayoutStackDirection;
  variant: LayoutVariant;
  region: string;
  spacing: LayoutSpacingPreset;
  alignment: LayoutAlignmentPreset;
  minHeight: number;
  responsive: LayoutResponsiveDefinition;
}

export interface LayoutBindingDefinition {
  objectKey?: string;
  workflowKey?: string;
  agentId?: string;
  relatedObjectKey?: string;
  viewKey?: string;
  promptAsset?: string;
  formKey?: string;
}

export interface LayoutComponentDefinition {
  id: string;
  kind: LayoutComponentKind;
  title: string;
  description?: string;
  objectKey?: string;
  workflowKey?: string;
  width: number;
  agentId?: string;
  relatedObjectKey?: string;
  stylePreset?: string;
  placement: LayoutPlacementDefinition;
  visibilityRule?: VisibilityRuleDefinition;
  binding?: LayoutBindingDefinition;
  props: Record<string, unknown>;
}

export interface BrandAssetDefinition {
  id: string;
  kind: BrandAssetKind;
  label: string;
  fileName: string;
  contentType: string;
  storageKey: string;
  url: string;
  uploadedAt: string;
}

export interface TenantBrandingDefinition {
  mode: BrandingMode;
  themeName: string;
  primaryColor: string;
  secondaryColor: string;
  accentColor: string;
  surfaceColor: string;
  textColor: string;
  pageBackground: string;
  fontFamily: string;
  logoAssetId?: string;
  iconAssetId?: string;
  brandBookAssetId?: string;
  notes?: string;
  assets: BrandAssetDefinition[];
}

export interface AppShellDefinition {
  productName: string;
  tagLine?: string;
  supportEmail?: string;
  menuStyle: "sidebar" | "topbar";
  profilePageKey?: string;
  settingsPageKey?: string;
}

export interface ProfileSettingsSectionDefinition {
  key: string;
  label: string;
  description?: string;
  preferenceKeys: string[];
}

export interface UserPreferenceDefinition {
  key: string;
  label: string;
  description?: string;
  type: "text" | "boolean" | "select";
  defaultValue?: string | boolean;
  options?: string[];
}

export interface ProfileConfigurationDefinition {
  objectKey: string;
  settingsObjectKey: string;
  pageTitle: string;
  profilePageKey?: string;
  settingsPageKey?: string;
  visibleFieldKeys: string[];
  preferences: UserPreferenceDefinition[];
  settingsSections: ProfileSettingsSectionDefinition[];
}

export interface FormFieldDefinition {
  id: string;
  key: string;
  label: string;
  type: FieldType;
  description?: string;
  helpText?: string;
  tooltip?: string;
  required: boolean;
  placeholder?: string;
  options?: string[];
  defaultValue?: string | number | boolean | null;
  validations: ValidationRuleDefinition[];
  calculation?: CalculationRuleDefinition | null;
  mandatoryRule?: RuleExpressionDefinition;
}

export interface FormStepDefinition {
  id: string;
  key: string;
  title: string;
  description?: string;
  fieldKeys: string[];
  visibilityRule?: VisibilityRuleDefinition;
}

export interface FormDefinition {
  id: string;
  key: string;
  title: string;
  description?: string;
  route: string;
  objectKey?: string;
  deliveryMode: FormDeliveryMode;
  submitLabel: string;
  successMessage: string;
  saveAndResume: boolean;
  requireAuthentication: boolean;
  analyticsEnabled: boolean;
  fields: FormFieldDefinition[];
  steps: FormStepDefinition[];
}

export interface LayoutSectionDefinition {
  id: string;
  title: string;
  description?: string;
  kind: LayoutSectionKind;
  columns: number;
  templateKey?: string;
  placement: LayoutPlacementDefinition;
  visibilityRule?: VisibilityRuleDefinition;
  components: LayoutComponentDefinition[];
}

export interface LayoutDefinition {
  id: string;
  key: string;
  name: string;
  pageKey: string;
  mobileColumns: number;
  tabletColumns: number;
  desktopColumns: number;
  sections: LayoutSectionDefinition[];
}

export interface PageDefinition {
  id: string;
  key: string;
  title: string;
  route: string;
  description?: string;
  layoutKey: string;
  objectKey?: string;
  isHome?: boolean;
  previewNote?: string;
}

export interface MenuItemDefinition {
  id: string;
  key: string;
  label: string;
  icon: string;
  pageKey: string;
  order: number;
  group: string;
}

export interface ManualTriggerConfig {
  notes?: string;
}

export interface RecordChangeTriggerConfig {
  objectKey: string;
}

export type TriggerConfig = ManualTriggerConfig | RecordChangeTriggerConfig;

export interface TriggerDefinition {
  id: string;
  type: TriggerType;
  label: string;
  config: TriggerConfig;
}

export interface ConditionNodeConfig {
  expression: string;
}

export interface CrudNodeConfig {
  operation: WorkflowCrudOperation;
  objectKey: string;
  targetFieldKey?: string;
  valueExpression?: string;
}

export interface FormulaNodeConfig {
  expression: string;
  outputKey: string;
}

export interface WebhookNodeConfig {
  method: WorkflowWebhookMethod;
  url: string;
  bodyTemplate?: string;
}

export interface NotificationNodeConfig {
  channel: WorkflowNotificationChannel;
  recipient?: string;
  message?: string;
}

export interface WaitNodeConfig {
  durationMinutes: number;
}

export interface ApprovalNodeConfig {
  approverRole: PlatformRole;
  instructions?: string;
}

export interface ModelCallNodeConfig {
  agentId: string;
  objectKey?: string;
  promptAsset?: string;
}

export type WorkflowNodeConfig =
  | ConditionNodeConfig
  | CrudNodeConfig
  | FormulaNodeConfig
  | WebhookNodeConfig
  | NotificationNodeConfig
  | WaitNodeConfig
  | ApprovalNodeConfig
  | ModelCallNodeConfig;

export interface WorkflowNodeDefinition {
  id: string;
  type: WorkflowNodeType;
  label: string;
  config: WorkflowNodeConfig;
  position: {
    x: number;
    y: number;
  };
}

export interface WorkflowEdgeDefinition {
  id: string;
  sourceId: string;
  targetId: string;
  label?: string;
}

export interface WorkflowDefinition {
  id: string;
  key: string;
  name: string;
  description?: string;
  objectKey?: string;
  status: "draft" | "active";
  triggers: TriggerDefinition[];
  nodes: WorkflowNodeDefinition[];
  edges: WorkflowEdgeDefinition[];
}

export interface ToolDefinition {
  id: string;
  key: string;
  name: string;
  type: "workflow" | "action" | "query";
  config: Record<string, unknown>;
}

export interface AgentDefinition {
  id: string;
  key: string;
  name: string;
  description?: string;
  scope: AgentScope;
  modelProviderId: string;
  prompt: string;
  allowedToolIds: string[];
  objectKeys: string[];
  zeroRetentionRequired: boolean;
}

export interface ModelProviderDefinition {
  id: string;
  key: string;
  name: string;
  provider: ModelProviderKind;
  model: string;
  endpoint?: string;
  apiKeySecretRef: string;
  supportsZeroRetention: boolean;
  allowedForSensitiveData: boolean;
  status: ModelProviderStatus;
}

export interface MaskingPolicyDefinition {
  id: string;
  key: string;
  name: string;
  mode: MaskingMode;
  targetSensitivities: FieldSensitivity[];
}

export interface SecurityPolicyDefinition {
  zeroRetentionRequiredForSensitiveData: boolean;
  defaultMaskingPolicyKey: string;
  allowedModelProviderKeys: string[];
}

export interface PlatformViewAsState {
  active: boolean;
  role: PlatformRole;
  personaLabel: string;
  actorEmail?: string;
}

export interface PlatformManifestMetadata {
  draftUpdatedAt: string;
  publishedAt: string | null;
  publishedVersionId: string | null;
}

export interface PlatformManifest {
  schemaVersion: typeof PLATFORM_SCHEMA_VERSION;
  tenant: {
    slug: string;
    name: string;
    description?: string;
  };
  environment: {
    slug: string;
    name: string;
  };
  roles: PlatformRole[];
  branding: TenantBrandingDefinition;
  appShell: AppShellDefinition;
  profiles: ProfileConfigurationDefinition;
  objects: ObjectDefinition[];
  layouts: LayoutDefinition[];
  pages: PageDefinition[];
  menus: MenuItemDefinition[];
  forms: FormDefinition[];
  workflows: WorkflowDefinition[];
  tools: ToolDefinition[];
  agents: AgentDefinition[];
  modelProviders: ModelProviderDefinition[];
  maskingPolicies: MaskingPolicyDefinition[];
  securityPolicy: SecurityPolicyDefinition;
  metadata: PlatformManifestMetadata;
}

export interface PlatformTenantSummary {
  id: string;
  slug: string;
  name: string;
  description?: string;
  defaultEnvironmentSlug: string;
}

export interface PlatformEnvironmentSummary {
  id: string;
  slug: string;
  name: string;
  isDefault: boolean;
}

export interface PlatformPublishedVersionRecord {
  id: string;
  versionNumber: number;
  status: PlatformVersionStatus;
  notes?: string | null;
  manifestPath?: string | null;
  gitCommitSha?: string | null;
  activatedAt?: string | null;
  createdAt: string;
}

export interface PlatformAuditEventRecord {
  id: string;
  action: string;
  resourceType: string;
  resourceId: string;
  summary: string;
  actorEmail?: string | null;
  actorRole?: PlatformRole | null;
  createdAt: string;
  payload?: Record<string, unknown> | null;
}

export interface PlatformRecord {
  id: string;
  objectKey: string;
  data: Record<string, unknown>;
  createdAt: string;
  updatedAt: string;
}

export interface PlatformWorkflowRunRecord {
  id: string;
  workflowId: string;
  workflowKey: string;
  status: WorkflowRunStatus;
  input?: Record<string, unknown> | null;
  output?: Record<string, unknown> | null;
  logs: Array<Record<string, unknown>>;
  startedAt?: string | null;
  finishedAt?: string | null;
  createdAt: string;
  updatedAt: string;
}

export interface PlatformFormSubmissionRecord {
  id: string;
  formKey: string;
  objectKey?: string;
  status: FormSubmissionStatus;
  data: Record<string, unknown>;
  createdAt: string;
  submittedAt?: string | null;
  createdByEmail?: string | null;
}

export interface PlatformAgentEvalRecord {
  id: string;
  agentId: string;
  agentKey: string;
  score: number;
  summary: string;
  createdAt: string;
  result: Record<string, unknown>;
}

export interface PublishPreviewBucket {
  added: string[];
  updated: string[];
  removed: string[];
}

export interface PlatformPublishPageImpact {
  pageKey: string;
  title: string;
  route: string;
  layoutKey: string;
  status: "added" | "updated" | "removed";
  sectionCount: number;
  componentCount: number;
}

export interface PlatformPublishRouteImpact {
  pageKey: string;
  route: string;
  status: "added" | "updated" | "removed";
}

export interface PlatformPublishPreview {
  tenantSlug: string;
  environmentSlug: string;
  activeVersionNumber: number | null;
  nextVersionNumber: number;
  summary: {
    objects: PublishPreviewBucket;
    pages: PublishPreviewBucket;
    layouts: PublishPreviewBucket;
    menus: PublishPreviewBucket;
    workflows: PublishPreviewBucket;
    agents: PublishPreviewBucket;
    modelProviders: PublishPreviewBucket;
  };
  pageImpacts: PlatformPublishPageImpact[];
  routeImpacts: PlatformPublishRouteImpact[];
}

export interface PlatformAgentPreview {
  agent: Pick<AgentDefinition, "id" | "key" | "name" | "scope" | "zeroRetentionRequired">;
  provider: Pick<ModelProviderDefinition, "id" | "key" | "name" | "provider" | "model" | "supportsZeroRetention" | "allowedForSensitiveData" | "status">;
  objectKey: string;
  sampleSize: number;
  maskedRecords: Array<Record<string, unknown>>;
  metadata: {
    masked: boolean;
    zeroRetentionRequired: boolean;
    allowedToolIds: string[];
    allowedObjectKeys: string[];
    allowedByPolicy: boolean;
  };
  recentActivity: PlatformAuditEventRecord[];
}

export interface PlatformBootstrap {
  tenant: PlatformTenantSummary;
  environment: PlatformEnvironmentSummary;
  actor: PlatformActor;
  session: PlatformSessionSummary;
  viewAs: PlatformViewAsState | null;
  draftManifest: PlatformManifest;
  activeVersion: PlatformPublishedVersionRecord | null;
  versions: PlatformPublishedVersionRecord[];
  auditEvents: PlatformAuditEventRecord[];
  invites: PlatformInviteRecord[];
  designerCatalog: PlatformDesignerCatalog;
}

export interface PlatformTenantMembershipSummary {
  tenantId: string;
  tenantSlug: string;
  tenantName: string;
  defaultEnvironmentSlug: string;
  role: PlatformRole;
}

export interface PlatformSessionSummary {
  actor: PlatformActor | null;
  memberships: PlatformTenantMembershipSummary[];
  source: "session" | "local_dev" | "none";
}

export interface PlatformInviteRecord {
  id: string;
  tenantId: string;
  tenantSlug: string;
  email: string;
  role: PlatformRole;
  status: "pending" | "accepted" | "revoked" | "expired";
  token: string;
  inviteUrl: string;
  createdByEmail?: string | null;
  createdAt: string;
  expiresAt: string;
  acceptedAt?: string | null;
}

export interface PlatformComponentPreset {
  key: string;
  label: string;
  description: string;
  kind: LayoutComponentKind;
  stylePreset: string;
  defaults: Partial<LayoutComponentDefinition>;
}

export interface PlatformSectionTemplate {
  key: string;
  label: string;
  description: string;
  section: Omit<LayoutSectionDefinition, "id" | "components"> & {
    components: Array<Omit<LayoutComponentDefinition, "id">>;
  };
}

export interface PlatformPageTemplate {
  key: string;
  label: string;
  description: string;
  page: Omit<PageDefinition, "id" | "route" | "layoutKey"> & { route?: string; layoutKey?: string };
  layout: Omit<LayoutDefinition, "id" | "pageKey" | "sections"> & {
    sections: Array<Omit<LayoutSectionDefinition, "id" | "components"> & { components: Array<Omit<LayoutComponentDefinition, "id">> }>;
  };
  source: "platform" | "tenant";
}

export interface PlatformDesignerCatalog {
  componentPresets: PlatformComponentPreset[];
  sectionTemplates: PlatformSectionTemplate[];
  pageTemplates: PlatformPageTemplate[];
}

export function cloneManifest(manifest: PlatformManifest): PlatformManifest {
  return structuredClone(manifest);
}
