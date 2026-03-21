import { normalizeLayoutSectionDefinition } from "@/lib/platform/designer";
import type {
  LayoutDefinition,
  ObjectDefinition,
  PageDefinition,
  PlatformManifest,
} from "@/lib/platform/types";

function hasRequiredText(value: string | null | undefined): boolean {
  return Boolean(value?.trim());
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
  const objects = manifest.objects
    .filter((objectDefinition) => hasRequiredText(objectDefinition.key) && hasRequiredText(objectDefinition.label) && hasRequiredText(objectDefinition.pluralLabel))
    .map((objectDefinition) => ({
      ...objectDefinition,
      fields: objectDefinition.fields.filter((field) => hasRequiredText(field.key) && hasRequiredText(field.label)),
    }))
    .filter((objectDefinition) => objectDefinition.fields.length > 0);
  const layouts = manifest.layouts.filter(
    (layout) => hasRequiredText(layout.key) && hasRequiredText(layout.name) && hasRequiredText(layout.pageKey),
  ).map((layout) => ({
    ...layout,
    sections: layout.sections
      .filter((section) => hasRequiredText(section.id) && hasRequiredText(section.title))
      .map(normalizeLayoutSectionDefinition),
  }));
  const pages = manifest.pages.filter(
    (page) => hasRequiredText(page.key) && hasRequiredText(page.title) && hasRequiredText(page.route) && hasRequiredText(page.layoutKey),
  ).map((page) => ({
    ...page,
    isHome: page.isHome ?? false,
    previewNote: hasRequiredText(page.previewNote) ? page.previewNote : undefined,
  }));
  const menus = manifest.menus.filter(
    (menu) => hasRequiredText(menu.key) && hasRequiredText(menu.label) && hasRequiredText(menu.pageKey),
  );
  const workflows = manifest.workflows.filter((workflow) => hasRequiredText(workflow.key) && hasRequiredText(workflow.name));
  const agents = manifest.agents.filter(
    (agent) =>
      hasRequiredText(agent.key) &&
      hasRequiredText(agent.name) &&
      hasRequiredText(agent.modelProviderId) &&
      hasRequiredText(agent.prompt),
  );
  const modelProviders = manifest.modelProviders.filter(
    (provider) =>
      hasRequiredText(provider.key) &&
      hasRequiredText(provider.name) &&
      hasRequiredText(provider.model) &&
      hasRequiredText(provider.apiKeySecretRef),
  );
  const tools = manifest.tools.filter((tool) => hasRequiredText(tool.key) && hasRequiredText(tool.name));

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
    objects,
    layouts,
    pages: normalizedPages,
    menus: menus
      .filter((menu) => pageKeys.has(menu.pageKey))
      .sort((left, right) => left.order - right.order),
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
