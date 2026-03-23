import { normalizeLayoutSectionDefinition } from "@/lib/platform/designer";
import { createStarterManifest } from "@/lib/platform/defaults";
import { normalizeRuleExpression } from "@/lib/platform/rule-engine";
import { createDefaultBranding } from "@/lib/platform/theme";
import {
  PLATFORM_SCHEMA_VERSION,
  type FormDefinition,
  type LayoutDefinition,
  type ObjectDefinition,
  type PageDefinition,
  type PlatformManifest,
  type PlatformRole,
} from "@/lib/platform/types";

function hasRequiredText(value: string | null | undefined): boolean {
  return Boolean(value?.trim());
}

function mergeByKey<T extends { key: string }>(current: T[], fallback: T[]): T[] {
  const seen = new Set(current.map((entry) => entry.key));
  return [...current, ...fallback.filter((entry) => !seen.has(entry.key))];
}

export function touchManifest(manifest: PlatformManifest): PlatformManifest {
  return {
    ...manifest,
    metadata: {
      ...manifest.metadata,
      draftUpdatedAt: new Date().toISOString(),
    },
  };
}

export function ensureManifestConsistency(manifest: PlatformManifest): PlatformManifest {
  const starterManifest = createStarterManifest(manifest.tenant.slug, manifest.tenant.name);

  const objects = mergeByKey(manifest.objects, starterManifest.objects)
    .filter((objectDefinition) => hasRequiredText(objectDefinition.key) && hasRequiredText(objectDefinition.label) && hasRequiredText(objectDefinition.pluralLabel))
    .map((objectDefinition) => ({
      ...objectDefinition,
      fields: objectDefinition.fields
        .filter((field) => hasRequiredText(field.key) && hasRequiredText(field.label))
        .map((field) => ({
          ...field,
          helpText: hasRequiredText(field.helpText) ? field.helpText?.trim() : undefined,
          tooltip: hasRequiredText(field.tooltip) ? field.tooltip?.trim() : undefined,
          fieldGroup: hasRequiredText(field.fieldGroup) ? field.fieldGroup?.trim() : undefined,
          mandatoryRule: normalizeRuleExpression(field.mandatoryRule),
          advancedValidation: normalizeRuleExpression(field.advancedValidation),
          validations: field.validations.map((rule) => ({
            ...rule,
            rule: normalizeRuleExpression(rule.rule),
          })),
          calculation: field.calculation
            ? {
                ...field.calculation,
                rule: normalizeRuleExpression(field.calculation.rule),
              }
            : null,
        })),
    }))
    .filter((objectDefinition) => objectDefinition.fields.length > 0);
  const layouts = mergeByKey(manifest.layouts, starterManifest.layouts).filter(
    (layout) => hasRequiredText(layout.key) && hasRequiredText(layout.name) && hasRequiredText(layout.pageKey),
  ).map((layout) => ({
    ...layout,
    sections: layout.sections
      .filter((section) => hasRequiredText(section.id) && hasRequiredText(section.title))
      .map(normalizeLayoutSectionDefinition),
  }));
  const pages = mergeByKey(manifest.pages, starterManifest.pages).filter(
    (page) => hasRequiredText(page.key) && hasRequiredText(page.title) && hasRequiredText(page.route) && hasRequiredText(page.layoutKey),
  ).map((page) => ({
    ...page,
    isHome: page.isHome ?? false,
    previewNote: hasRequiredText(page.previewNote) ? page.previewNote : undefined,
  }));
  const menus = mergeByKey(manifest.menus, starterManifest.menus).filter(
    (menu) => hasRequiredText(menu.key) && hasRequiredText(menu.label) && hasRequiredText(menu.pageKey),
  ).map((menu) => ({
    ...menu,
    description: hasRequiredText(menu.description) ? menu.description?.trim() : undefined,
    groupKey: hasRequiredText(menu.groupKey) ? menu.groupKey?.trim() : undefined,
    visibleToRoles: menu.visibleToRoles?.length ? menu.visibleToRoles : (["SUPER_ADMIN", "BUILDER_ADMIN", "USER"] satisfies PlatformRole[]),
    highlight: menu.highlight ?? false,
    badgeBindingKey: hasRequiredText(menu.badgeBindingKey) ? menu.badgeBindingKey?.trim() : undefined,
  }));
  const forms = mergeByKey(manifest.forms ?? [], starterManifest.forms).filter(
    (form) => hasRequiredText(form.key) && hasRequiredText(form.title) && hasRequiredText(form.route),
  ).map((form): FormDefinition => ({
    ...form,
    description: hasRequiredText(form.description) ? form.description?.trim() : undefined,
    submitLabel: hasRequiredText(form.submitLabel) ? form.submitLabel : "Submit",
    successMessage: hasRequiredText(form.successMessage) ? form.successMessage : "Submitted successfully.",
    fields: form.fields
      .filter((field) => hasRequiredText(field.key) && hasRequiredText(field.label))
      .map((field) => ({
        ...field,
        helpText: hasRequiredText(field.helpText) ? field.helpText?.trim() : undefined,
        tooltip: hasRequiredText(field.tooltip) ? field.tooltip?.trim() : undefined,
        mandatoryRule: normalizeRuleExpression(field.mandatoryRule),
        validations: field.validations.map((rule) => ({
          ...rule,
          rule: normalizeRuleExpression(rule.rule),
        })),
        calculation: field.calculation
          ? {
              ...field.calculation,
              rule: normalizeRuleExpression(field.calculation.rule),
            }
          : null,
      })),
    steps: form.steps.filter((step) => hasRequiredText(step.key) && hasRequiredText(step.title)).map((step) => ({
      ...step,
      description: hasRequiredText(step.description) ? step.description?.trim() : undefined,
      visibilityRule: step.visibilityRule
        ? {
            ...step.visibilityRule,
            summary: hasRequiredText(step.visibilityRule.summary) ? step.visibilityRule.summary?.trim() : undefined,
            rule: normalizeRuleExpression(step.visibilityRule.rule),
          }
        : undefined,
    })),
  }));
  const workflows = mergeByKey(manifest.workflows, starterManifest.workflows).filter((workflow) => hasRequiredText(workflow.key) && hasRequiredText(workflow.name));
  const agents = mergeByKey(manifest.agents, starterManifest.agents).filter(
    (agent) =>
      hasRequiredText(agent.key) &&
      hasRequiredText(agent.name) &&
      hasRequiredText(agent.modelProviderId) &&
      hasRequiredText(agent.prompt),
  ).map((agent) => ({
    ...agent,
    promptBlocks: agent.promptBlocks?.filter((block) => hasRequiredText(block.label) && hasRequiredText(block.content)) ?? [],
    handoffWorkflowKeys: agent.handoffWorkflowKeys ?? [],
    outputSchema: hasRequiredText(agent.outputSchema) ? agent.outputSchema?.trim() : undefined,
    evalPolicy: agent.evalPolicy
      ? {
          rubric: hasRequiredText(agent.evalPolicy.rubric) ? agent.evalPolicy.rubric.trim() : "Evaluate safety and operational usefulness.",
          samplePrompt: hasRequiredText(agent.evalPolicy.samplePrompt)
            ? agent.evalPolicy.samplePrompt.trim()
            : "Summarise the current workload.",
          passingScore: agent.evalPolicy.passingScore ?? 0.8,
        }
      : undefined,
    costBudgetUsd: typeof agent.costBudgetUsd === "number" ? agent.costBudgetUsd : undefined,
    approvalPolicy: agent.approvalPolicy
      ? {
          required: agent.approvalPolicy.required ?? false,
          approverRole: agent.approvalPolicy.approverRole,
          notes: hasRequiredText(agent.approvalPolicy.notes) ? agent.approvalPolicy.notes?.trim() : undefined,
        }
      : undefined,
  }));
  const modelProviders = mergeByKey(manifest.modelProviders, starterManifest.modelProviders).filter(
    (provider) =>
      hasRequiredText(provider.key) &&
      hasRequiredText(provider.name) &&
      hasRequiredText(provider.model) &&
      hasRequiredText(provider.apiKeySecretRef),
  );
  const tools = mergeByKey(manifest.tools, starterManifest.tools).filter((tool) => hasRequiredText(tool.key) && hasRequiredText(tool.name));

  const layoutKeys = new Set(layouts.map((layout) => layout.key));

  const normalizedPages = pages.filter((page) => layoutKeys.has(page.layoutKey));
  if (normalizedPages.length > 0 && !normalizedPages.some((page) => page.isHome)) {
    normalizedPages[0] = {
      ...normalizedPages[0],
      isHome: true,
    };
  }
  const pageKeys = new Set(normalizedPages.map((page) => page.key));

  return {
    ...manifest,
    schemaVersion: PLATFORM_SCHEMA_VERSION,
    branding: {
      ...createDefaultBranding(manifest.tenant.name),
      ...(manifest.branding ?? {}),
      assets: manifest.branding?.assets ?? [],
    },
    appShell: {
      productName: manifest.appShell?.productName?.trim() || `${manifest.tenant.name} Adaptive Platform`,
      tagLine: hasRequiredText(manifest.appShell?.tagLine) ? manifest.appShell?.tagLine?.trim() : undefined,
      supportEmail: hasRequiredText(manifest.appShell?.supportEmail) ? manifest.appShell?.supportEmail?.trim() : undefined,
      menuStyle: manifest.appShell?.menuStyle ?? "sidebar",
      navigationMode: manifest.appShell?.navigationMode ?? manifest.appShell?.menuStyle ?? "sidebar",
      menuGroups: manifest.appShell?.menuGroups?.length ? manifest.appShell.menuGroups : starterManifest.appShell.menuGroups,
      quickActions: manifest.appShell?.quickActions ?? starterManifest.appShell.quickActions,
      defaultLandingPageKey: hasRequiredText(manifest.appShell?.defaultLandingPageKey)
        ? manifest.appShell?.defaultLandingPageKey?.trim()
        : undefined,
      announcementSlots: manifest.appShell?.announcementSlots ?? starterManifest.appShell.announcementSlots,
      badgeBindings: manifest.appShell?.badgeBindings ?? starterManifest.appShell.badgeBindings,
      visibilityRules: manifest.appShell?.visibilityRules ?? starterManifest.appShell.visibilityRules,
      profilePageKey: hasRequiredText(manifest.appShell?.profilePageKey) ? manifest.appShell?.profilePageKey?.trim() : undefined,
      settingsPageKey: hasRequiredText(manifest.appShell?.settingsPageKey) ? manifest.appShell?.settingsPageKey?.trim() : undefined,
    },
    notifications: {
      channels: manifest.notifications?.channels ?? starterManifest.notifications.channels,
      templates: manifest.notifications?.templates ?? starterManifest.notifications.templates,
      rules: manifest.notifications?.rules ?? starterManifest.notifications.rules,
    },
    profiles: {
      objectKey: manifest.profiles?.objectKey ?? "user_profile",
      settingsObjectKey: manifest.profiles?.settingsObjectKey ?? "user_setting",
      pageTitle: manifest.profiles?.pageTitle?.trim() || "Profile",
      profilePageKey: hasRequiredText(manifest.profiles?.profilePageKey) ? manifest.profiles?.profilePageKey?.trim() : undefined,
      settingsPageKey: hasRequiredText(manifest.profiles?.settingsPageKey) ? manifest.profiles?.settingsPageKey?.trim() : undefined,
      visibleFieldKeys: manifest.profiles?.visibleFieldKeys ?? [],
      preferences: manifest.profiles?.preferences ?? [],
      settingsSections: manifest.profiles?.settingsSections ?? [],
    },
    objects,
    layouts,
    pages: normalizedPages,
    menus: menus
      .filter((menu) => pageKeys.has(menu.pageKey))
      .sort((left, right) => left.order - right.order),
    forms,
    workflows,
    tools,
    agents,
    modelProviders,
  };
}

export function findObjectDefinition(manifest: PlatformManifest, objectKeyOrId: string): ObjectDefinition | undefined {
  return manifest.objects.find((objectDefinition) => objectDefinition.key === objectKeyOrId || objectDefinition.id === objectKeyOrId);
}

export function findPageDefinition(manifest: PlatformManifest, pageKeyOrId: string): PageDefinition | undefined {
  return manifest.pages.find((pageDefinition) => pageDefinition.key === pageKeyOrId || pageDefinition.id === pageKeyOrId);
}

export function findLayoutDefinition(manifest: PlatformManifest, layoutKeyOrId: string): LayoutDefinition | undefined {
  return manifest.layouts.find((layout) => layout.key === layoutKeyOrId || layout.id === layoutKeyOrId);
}

export function findLayoutForPage(manifest: PlatformManifest, page: PageDefinition): LayoutDefinition | undefined {
  return findLayoutDefinition(manifest, page.layoutKey);
}
