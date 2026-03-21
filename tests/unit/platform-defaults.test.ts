import { describe, expect, it } from "vitest";

import { createStarterManifest } from "@/lib/platform/defaults";

describe("createStarterManifest", () => {
  it("creates a coherent starter manifest with runtime navigation", () => {
    const manifest = createStarterManifest("teacheractive");

    expect(manifest.schemaVersion).toBe(2);
    expect(manifest.objects.length).toBeGreaterThan(0);
    expect(manifest.pages.length).toBeGreaterThan(0);
    expect(manifest.menus.length).toBeGreaterThan(0);

    for (const menu of manifest.menus) {
      expect(manifest.pages.some((page) => page.key === menu.pageKey)).toBe(true);
    }

    for (const page of manifest.pages) {
      expect(manifest.layouts.some((layout) => layout.key === page.layoutKey)).toBe(true);
    }
  });
});
