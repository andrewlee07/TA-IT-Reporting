import type {
  LayoutComponentDefinition,
  LayoutComponentKind,
  LayoutPlacementDefinition,
  LayoutResponsiveDefinition,
  LayoutSectionDefinition,
  LayoutVariant,
  LayoutZone,
  ObjectDefinition,
  PageDefinition,
  PlatformComponentPreset,
  PlatformDesignerCatalog,
  PlatformManifest,
  PlatformPageTemplate,
  PlatformSectionTemplate,
  ResponsiveBreakpoint,
  VisibilityRuleDefinition,
} from "@/lib/platform/types";

const DEFAULT_HIDDEN_ON: ResponsiveBreakpoint[] = [];

function clampSpan(value: number | undefined, fallback = 12): number {
  const candidate = Number.isFinite(value) ? Number(value) : fallback;
  return Math.min(12, Math.max(1, Math.round(candidate)));
}

export function createDefaultResponsiveDefinition(span = 12): LayoutResponsiveDefinition {
  return {
    mobileSpan: 12,
    tabletSpan: clampSpan(Math.min(span, 6), 6),
    desktopSpan: clampSpan(span),
    hiddenOn: DEFAULT_HIDDEN_ON,
  };
}

export function createDefaultPlacement(input?: Partial<LayoutPlacementDefinition> & { span?: number }): LayoutPlacementDefinition {
  const span = clampSpan(input?.span, 12);
  return {
    zone: input?.zone ?? "main",
    span,
    stackDirection: input?.stackDirection ?? "column",
    variant: input?.variant ?? "standard",
    region: input?.region?.trim() || "body",
    spacing: input?.spacing ?? "comfortable",
    alignment: input?.alignment ?? "start",
    minHeight: Math.max(0, Math.round(input?.minHeight ?? 220)),
    responsive: {
      ...createDefaultResponsiveDefinition(span),
      ...input?.responsive,
      hiddenOn: input?.responsive?.hiddenOn ?? DEFAULT_HIDDEN_ON,
      mobileSpan: clampSpan(input?.responsive?.mobileSpan, 12),
      tabletSpan: clampSpan(input?.responsive?.tabletSpan, Math.min(span, 6)),
      desktopSpan: clampSpan(input?.responsive?.desktopSpan, span),
    },
  };
}

export function normalizeVisibilityRule(rule: VisibilityRuleDefinition | undefined): VisibilityRuleDefinition | undefined {
  if (!rule?.expression?.trim()) {
    return undefined;
  }

  return {
    expression: rule.expression.trim(),
    summary: rule.summary?.trim() || undefined,
  };
}

export function normalizeLayoutComponentDefinition(component: LayoutComponentDefinition): LayoutComponentDefinition {
  const binding = {
    objectKey: component.binding?.objectKey ?? component.objectKey,
    workflowKey: component.binding?.workflowKey ?? component.workflowKey,
    agentId: component.binding?.agentId ?? component.agentId,
    relatedObjectKey: component.binding?.relatedObjectKey ?? component.relatedObjectKey,
    viewKey: component.binding?.viewKey,
    promptAsset: component.binding?.promptAsset,
    formKey: component.binding?.formKey,
  };

  return {
    ...component,
    width: clampSpan(component.width, component.placement?.span ?? 12),
    stylePreset: component.stylePreset ?? defaultStylePreset(component.kind),
    objectKey: binding.objectKey,
    workflowKey: binding.workflowKey,
    agentId: binding.agentId,
    relatedObjectKey: binding.relatedObjectKey,
    binding,
    placement: createDefaultPlacement({
      ...component.placement,
      span: component.placement?.span ?? component.width ?? 12,
    }),
    visibilityRule: normalizeVisibilityRule(component.visibilityRule),
    props: component.props ?? {},
  };
}

export function normalizeLayoutSectionDefinition(section: LayoutSectionDefinition): LayoutSectionDefinition {
  return {
    ...section,
    columns: clampSpan(section.columns, 12),
    templateKey: section.templateKey?.trim() || undefined,
    placement: createDefaultPlacement({
      ...section.placement,
      variant: section.placement?.variant ?? sectionVariantFromKind(section.kind),
      span: section.placement?.span ?? 12,
    }),
    visibilityRule: normalizeVisibilityRule(section.visibilityRule),
    components: section.components.map(normalizeLayoutComponentDefinition),
  };
}

function defaultStylePreset(kind: LayoutComponentKind): string {
  switch (kind) {
    case "hero":
      return "hero-banner";
    case "callout":
      return "callout-band";
    case "agent_panel":
      return "agent-panel";
    case "workflow_launcher":
      return "workflow-stack";
    default:
      return "surface";
  }
}

function sectionVariantFromKind(kind: LayoutSectionDefinition["kind"]): LayoutVariant {
  if (kind === "tabs") {
    return "tabs";
  }

  if (kind === "drawer") {
    return "drawer";
  }

  return "standard";
}

function sectionPlacement(overrides?: Partial<LayoutPlacementDefinition>): LayoutPlacementDefinition {
  return createDefaultPlacement({
    zone: "main",
    span: 12,
    minHeight: 260,
    variant: "standard",
    region: "body",
    spacing: "comfortable",
    alignment: "start",
    ...overrides,
  });
}

function componentPlacement(zone: LayoutZone, span: number, overrides?: Partial<LayoutPlacementDefinition>): LayoutPlacementDefinition {
  return createDefaultPlacement({
    zone,
    span,
    minHeight: 220,
    variant: zone === "rail" ? "rail" : "standard",
    region: zone === "header" ? "header" : "body",
    spacing: "comfortable",
    alignment: "start",
    ...overrides,
  });
}

export const DESIGNER_COMPONENT_PRESETS: PlatformComponentPreset[] = [
  {
    key: "hero",
    label: "Hero",
    description: "A headline-led banner for landing pages and workspace overviews.",
    kind: "hero",
    stylePreset: "hero-banner",
    defaults: {
      title: "Hero headline",
      description: "Orient operators with the page goal, next action, and context.",
      width: 12,
      placement: componentPlacement("header", 12, { minHeight: 280, variant: "full_width" }),
      props: { eyebrow: "Workspace" },
    },
  },
  {
    key: "rich_text",
    label: "Rich text",
    description: "Narrative copy, instructions, or policy guidance.",
    kind: "rich_text",
    stylePreset: "editorial",
    defaults: {
      title: "Guidance",
      description: "Add narrative guidance or process instructions.",
      width: 8,
      placement: componentPlacement("main", 8, { minHeight: 220 }),
      props: { body: "Use this panel for rich text, markdown, or builder instructions." },
    },
  },
  {
    key: "stat_tiles",
    label: "Stat tiles",
    description: "A compact summary strip for key platform metrics.",
    kind: "stat_tiles",
    stylePreset: "metric-strip",
    defaults: {
      title: "Key stats",
      width: 12,
      placement: componentPlacement("main", 12, { minHeight: 180 }),
      props: {
        metrics: [
          { label: "Objects", metric: "objects" },
          { label: "Pages", metric: "pages" },
          { label: "Workflows", metric: "workflows" },
          { label: "Agents", metric: "agents" },
        ],
      },
    },
  },
  {
    key: "record_table",
    label: "Record table",
    description: "A data grid bound to an object and optional view.",
    kind: "record_table",
    stylePreset: "record-grid",
    defaults: {
      title: "Records",
      width: 8,
      placement: componentPlacement("main", 8, { minHeight: 360 }),
      props: {},
    },
  },
  {
    key: "record_form",
    label: "Record form",
    description: "A create or edit form for a bound object.",
    kind: "record_form",
    stylePreset: "record-form",
    defaults: {
      title: "Create record",
      width: 4,
      placement: componentPlacement("rail", 4, { minHeight: 360, variant: "rail" }),
      props: {},
    },
  },
  {
    key: "related_records",
    label: "Related records",
    description: "Show related records or supporting object lists.",
    kind: "related_records",
    stylePreset: "linked-records",
    defaults: {
      title: "Related records",
      width: 6,
      placement: componentPlacement("main", 6, { minHeight: 260 }),
      props: { emptyState: "Choose a related object to populate this panel." },
    },
  },
  {
    key: "workflow_launcher",
    label: "Workflow launcher",
    description: "Expose a workflow as an operational action card.",
    kind: "workflow_launcher",
    stylePreset: "workflow-stack",
    defaults: {
      title: "Run workflow",
      width: 4,
      placement: componentPlacement("rail", 4, { minHeight: 220, variant: "rail" }),
      props: {},
    },
  },
  {
    key: "agent_summary",
    label: "Agent summary",
    description: "Compact AI posture summary with scoped objects and provider.",
    kind: "agent_summary",
    stylePreset: "agent-summary",
    defaults: {
      title: "Agent summary",
      width: 4,
      placement: componentPlacement("rail", 4, { minHeight: 220, variant: "rail" }),
      props: {},
    },
  },
  {
    key: "agent_panel",
    label: "Agent panel",
    description: "Expanded agent prompt surface with masked output preview.",
    kind: "agent_panel",
    stylePreset: "agent-panel",
    defaults: {
      title: "Agent panel",
      width: 8,
      placement: componentPlacement("main", 8, { minHeight: 320 }),
      props: {},
    },
  },
  {
    key: "activity_feed",
    label: "Activity feed",
    description: "Operational activity, publish events, and workflow actions.",
    kind: "activity_feed",
    stylePreset: "activity-stream",
    defaults: {
      title: "Activity feed",
      width: 6,
      placement: componentPlacement("main", 6, { minHeight: 320 }),
      props: {},
    },
  },
  {
    key: "callout",
    label: "Callout",
    description: "A highlighted guidance or exception card.",
    kind: "callout",
    stylePreset: "callout-band",
    defaults: {
      title: "Callout",
      description: "Raise a blocker, policy reminder, or critical note.",
      width: 12,
      placement: componentPlacement("main", 12, { minHeight: 180, variant: "full_width" }),
      props: { tone: "info", body: "Use this card for high-signal operator messaging." },
    },
  },
];

export const DESIGNER_SECTION_TEMPLATES: PlatformSectionTemplate[] = [
  {
    key: "overview_hero_strip",
    label: "Overview hero strip",
    description: "Hero plus stat tiles for a workspace landing page.",
    section: {
      title: "Workspace overview",
      description: "Introduce the page and frame the current operational state.",
      kind: "grid",
      columns: 12,
      templateKey: "overview_hero_strip",
      placement: sectionPlacement({ minHeight: 420 }),
      components: [DESIGNER_COMPONENT_PRESETS[0]!.defaults, DESIGNER_COMPONENT_PRESETS[2]!.defaults].map((defaults) => ({
        id: "template-component",
        kind: defaults.kind ?? "callout",
        title: defaults.title ?? "Template component",
        description: defaults.description,
        objectKey: defaults.objectKey,
        workflowKey: defaults.workflowKey,
        width: defaults.width ?? 12,
        agentId: defaults.agentId,
        relatedObjectKey: defaults.relatedObjectKey,
        stylePreset: defaults.stylePreset,
        placement: createDefaultPlacement(defaults.placement),
        visibilityRule: defaults.visibilityRule,
        binding: defaults.binding,
        props: defaults.props ?? {},
      })),
    },
  },
  {
    key: "list_and_rail",
    label: "List + rail",
    description: "Main list with a supporting rail panel for forms or agents.",
    section: {
      title: "Records and actions",
      description: "Keep primary data on the left and operator actions on the right.",
      kind: "grid",
      columns: 12,
      templateKey: "list_and_rail",
      placement: sectionPlacement(),
      components: [DESIGNER_COMPONENT_PRESETS[3]!.defaults, DESIGNER_COMPONENT_PRESETS[4]!.defaults].map((defaults) => ({
        id: "template-component",
        kind: defaults.kind ?? "callout",
        title: defaults.title ?? "Template component",
        description: defaults.description,
        objectKey: defaults.objectKey,
        workflowKey: defaults.workflowKey,
        width: defaults.width ?? 12,
        agentId: defaults.agentId,
        relatedObjectKey: defaults.relatedObjectKey,
        stylePreset: defaults.stylePreset,
        placement: createDefaultPlacement(defaults.placement),
        visibilityRule: defaults.visibilityRule,
        binding: defaults.binding,
        props: defaults.props ?? {},
      })),
    },
  },
  {
    key: "operations_feed",
    label: "Operations feed",
    description: "Workflow actions, agent posture, and activity feed.",
    section: {
      title: "Operations feed",
      description: "Surface automation and AI controls next to recent activity.",
      kind: "grid",
      columns: 12,
      templateKey: "operations_feed",
      placement: sectionPlacement({ minHeight: 320 }),
      components: [
        DESIGNER_COMPONENT_PRESETS[6]!.defaults,
        DESIGNER_COMPONENT_PRESETS[7]!.defaults,
        DESIGNER_COMPONENT_PRESETS[9]!.defaults,
      ].map((defaults) => ({
        id: "template-component",
        kind: defaults.kind ?? "callout",
        title: defaults.title ?? "Template component",
        description: defaults.description,
        objectKey: defaults.objectKey,
        workflowKey: defaults.workflowKey,
        width: defaults.width ?? 12,
        agentId: defaults.agentId,
        relatedObjectKey: defaults.relatedObjectKey,
        stylePreset: defaults.stylePreset,
        placement: createDefaultPlacement(defaults.placement),
        visibilityRule: defaults.visibilityRule,
        binding: defaults.binding,
        props: defaults.props ?? {},
      })),
    },
  },
];

export function createDesignerCatalog(manifest: PlatformManifest): PlatformDesignerCatalog {
  return {
    componentPresets: DESIGNER_COMPONENT_PRESETS,
    sectionTemplates: DESIGNER_SECTION_TEMPLATES,
    pageTemplates: getTenantPageTemplates(manifest),
  };
}

export function getTenantPageTemplates(manifest: PlatformManifest): PlatformPageTemplate[] {
  return manifest.pages.slice(0, 6).flatMap((page) => {
    const layout = manifest.layouts.find((candidate) => candidate.key === page.layoutKey);
    if (!layout) {
      return [];
    }

    return [
      {
        key: `tenant_${page.key}`,
        label: page.title,
        description: page.description ?? `Reuse the ${page.title} page structure as a template.`,
        source: "tenant",
        page: {
          key: page.key,
          title: page.title,
          description: page.description,
          objectKey: page.objectKey,
          isHome: page.isHome ?? false,
          previewNote: page.previewNote,
        },
        layout: {
          key: layout.key,
          name: layout.name,
          mobileColumns: layout.mobileColumns,
          tabletColumns: layout.tabletColumns,
          desktopColumns: layout.desktopColumns,
          sections: layout.sections.map((section) => ({
            ...section,
            id: "template-section",
            components: section.components.map((component) => ({
              ...component,
              id: "template-component",
            })),
          })),
        },
      },
    ];
  });
}

export function createComponentFromPreset(input: {
  presetKey: string;
  idFactory: (prefix: string) => string;
  objectKey?: string;
}): LayoutComponentDefinition {
  const preset = DESIGNER_COMPONENT_PRESETS.find((candidate) => candidate.key === input.presetKey) ?? DESIGNER_COMPONENT_PRESETS[0]!;
  const defaults = preset.defaults;
  const bindingObjectKey = input.objectKey ?? defaults.binding?.objectKey ?? defaults.objectKey;

  return normalizeLayoutComponentDefinition({
    id: input.idFactory("component"),
    kind: preset.kind,
    title: defaults.title ?? preset.label,
    description: defaults.description,
    objectKey: bindingObjectKey,
    workflowKey: defaults.workflowKey,
    width: defaults.width ?? 12,
    agentId: defaults.agentId,
    relatedObjectKey: defaults.relatedObjectKey,
    stylePreset: preset.stylePreset,
    placement: createDefaultPlacement(defaults.placement),
    visibilityRule: defaults.visibilityRule,
    binding: {
      ...defaults.binding,
      objectKey: bindingObjectKey,
    },
    props: structuredClone(defaults.props ?? {}),
  });
}

export function createSectionFromTemplate(input: {
  templateKey: string;
  idFactory: (prefix: string) => string;
  objectKey?: string;
}): LayoutSectionDefinition {
  const template = DESIGNER_SECTION_TEMPLATES.find((candidate) => candidate.key === input.templateKey) ?? DESIGNER_SECTION_TEMPLATES[0]!;
  return normalizeLayoutSectionDefinition({
    id: input.idFactory("section"),
    title: template.section.title,
    description: template.section.description,
    kind: template.section.kind,
    columns: template.section.columns,
    templateKey: template.key,
    placement: createDefaultPlacement(template.section.placement),
    visibilityRule: template.section.visibilityRule,
    components: template.section.components.map((component) =>
      normalizeLayoutComponentDefinition({
        ...component,
        id: input.idFactory("component"),
        objectKey: input.objectKey ?? component.objectKey,
        binding: {
          ...component.binding,
          objectKey: input.objectKey ?? component.binding?.objectKey ?? component.objectKey,
        },
      }),
    ),
  });
}

export function createPageTemplateInstance(input: {
  template: PlatformPageTemplate;
  idFactory: (prefix: string) => string;
  pageKey: string;
  title: string;
  route: string;
  objectKey?: string;
}): { page: PageDefinition; layoutSections: LayoutSectionDefinition[] } {
  return {
    page: {
      id: input.idFactory("page"),
      key: input.pageKey,
      title: input.title,
      route: input.route,
      description: input.template.page.description,
      layoutKey: input.pageKey,
      objectKey: input.objectKey ?? input.template.page.objectKey,
      isHome: input.template.page.isHome ?? false,
      previewNote: input.template.page.previewNote,
    },
    layoutSections: input.template.layout.sections.map((section) =>
      normalizeLayoutSectionDefinition({
        ...section,
        id: input.idFactory("section"),
        components: section.components.map((component) =>
          normalizeLayoutComponentDefinition({
            ...component,
            id: input.idFactory("component"),
            objectKey: input.objectKey ?? component.objectKey,
            binding: {
              ...component.binding,
              objectKey: input.objectKey ?? component.binding?.objectKey ?? component.objectKey,
            },
          }),
        ),
      }),
    ),
  };
}

export function countLayoutComponents(sections: LayoutSectionDefinition[]): number {
  return sections.reduce((total, section) => total + section.components.length, 0);
}

export function getSampleFieldValue(objectDefinition: ObjectDefinition, index: number, fieldKey?: string): string {
  const field = objectDefinition.fields.find((candidate) => candidate.key === fieldKey) ?? objectDefinition.fields[0];
  if (!field) {
    return "Sample value";
  }

  switch (field.type) {
    case "number":
    case "currency":
      return String((index + 1) * 125);
    case "boolean":
      return index % 2 === 0 ? "Yes" : "No";
    case "date":
      return `2026-04-${String(index + 4).padStart(2, "0")}`;
    case "datetime":
      return `2026-04-${String(index + 4).padStart(2, "0")}T09:00`;
    case "select":
      return field.options?.[index % (field.options.length || 1)] ?? "Open";
    default:
      return `${field.label} ${index + 1}`;
  }
}
