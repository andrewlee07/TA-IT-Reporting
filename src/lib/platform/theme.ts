import type { CSSProperties } from "react";

import type { PlatformManifest, TenantBrandingDefinition } from "@/lib/platform/types";

function normalizeHex(input: string | undefined, fallback: string): string {
  const value = input?.trim() || fallback;
  if (/^#[0-9a-f]{6}$/i.test(value)) {
    return value;
  }

  if (/^#[0-9a-f]{3}$/i.test(value)) {
    return `#${value[1]}${value[1]}${value[2]}${value[2]}${value[3]}${value[3]}`;
  }

  return fallback;
}

function hexToRgb(hex: string): { r: number; g: number; b: number } {
  const normalized = normalizeHex(hex, "#005292");
  return {
    r: Number.parseInt(normalized.slice(1, 3), 16),
    g: Number.parseInt(normalized.slice(3, 5), 16),
    b: Number.parseInt(normalized.slice(5, 7), 16),
  };
}

function rgbToHex(r: number, g: number, b: number): string {
  const toHex = (value: number) => Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, "0");
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

function mix(hex: string, target: string, weight: number): string {
  const left = hexToRgb(hex);
  const right = hexToRgb(target);
  return rgbToHex(
    left.r + (right.r - left.r) * weight,
    left.g + (right.g - left.g) * weight,
    left.b + (right.b - left.b) * weight,
  );
}

export function createDefaultBranding(tenantName: string): TenantBrandingDefinition {
  return {
    mode: "draft",
    themeName: `${tenantName} Default`,
    primaryColor: "#005292",
    secondaryColor: "#003d6e",
    accentColor: "#f57d00",
    surfaceColor: "#ffffff",
    textColor: "#111827",
    pageBackground: "#d8dde3",
    fontFamily: "var(--font-inter, 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif)",
    logoAssetId: undefined,
    iconAssetId: undefined,
    brandBookAssetId: undefined,
    notes: "",
    assets: [],
  };
}

export function getTenantBranding(manifest: PlatformManifest): TenantBrandingDefinition {
  return manifest.branding ?? createDefaultBranding(manifest.tenant.name);
}

export function getThemeCssVariables(manifest: PlatformManifest): CSSProperties {
  const branding = getTenantBranding(manifest);
  const primary = normalizeHex(branding.primaryColor, "#005292");
  const secondary = normalizeHex(branding.secondaryColor, mix(primary, "#111827", 0.24));
  const accent = normalizeHex(branding.accentColor, "#f57d00");
  const surface = normalizeHex(branding.surfaceColor, "#ffffff");
  const text = normalizeHex(branding.textColor, "#111827");
  const pageBackground = normalizeHex(branding.pageBackground, mix(primary, "#ffffff", 0.82));

  return {
    "--blue": primary,
    "--blue-dark": secondary,
    "--blue-soft": mix(primary, "#ffffff", 0.9),
    "--orange": accent,
    "--orange-soft": mix(accent, "#ffffff", 0.88),
    "--surface": surface,
    "--surface-alt": mix(surface, "#0f172a", 0.03),
    "--text": text,
    "--text-soft": mix(text, "#ffffff", 0.3),
    "--text-muted": mix(text, "#ffffff", 0.55),
    "--page-background": pageBackground,
    "--font-brand": branding.fontFamily,
  } as CSSProperties;
}
