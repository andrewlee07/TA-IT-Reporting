import type {
  TenantBrandingDefinition,
  ThemeAccessibilityCheck,
  ThemeAccessibilityReport,
  ThemeTokenSuggestion,
} from "@/lib/platform/types";

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

function mix(leftHex: string, rightHex: string, weight: number): string {
  const left = hexToRgb(leftHex);
  const right = hexToRgb(rightHex);

  return rgbToHex(
    left.r + (right.r - left.r) * weight,
    left.g + (right.g - left.g) * weight,
    left.b + (right.b - left.b) * weight,
  );
}

function luminanceChannel(value: number): number {
  const normalized = value / 255;
  return normalized <= 0.03928 ? normalized / 12.92 : ((normalized + 0.055) / 1.055) ** 2.4;
}

function contrastRatio(foreground: string, background: string): number {
  const foregroundRgb = hexToRgb(foreground);
  const backgroundRgb = hexToRgb(background);
  const foregroundLum =
    0.2126 * luminanceChannel(foregroundRgb.r) +
    0.7152 * luminanceChannel(foregroundRgb.g) +
    0.0722 * luminanceChannel(foregroundRgb.b);
  const backgroundLum =
    0.2126 * luminanceChannel(backgroundRgb.r) +
    0.7152 * luminanceChannel(backgroundRgb.g) +
    0.0722 * luminanceChannel(backgroundRgb.b);

  const lighter = Math.max(foregroundLum, backgroundLum);
  const darker = Math.min(foregroundLum, backgroundLum);
  return Number((((lighter + 0.05) / (darker + 0.05)) * 100).toFixed(0)) / 100;
}

function makeCheck(key: string, label: string, foreground: string, background: string, requiredRatio: number): ThemeAccessibilityCheck {
  const ratio = contrastRatio(foreground, background);
  return {
    key,
    label,
    ratio,
    requiredRatio,
    passed: ratio >= requiredRatio,
  };
}

export function createThemeAccessibilityReport(branding: TenantBrandingDefinition): ThemeAccessibilityReport {
  const primary = normalizeHex(branding.primaryColor, "#005292");
  const secondary = normalizeHex(branding.secondaryColor, "#003d6e");
  const accent = normalizeHex(branding.accentColor, "#f57d00");
  const surface = normalizeHex(branding.surfaceColor, "#ffffff");
  const text = normalizeHex(branding.textColor, "#111827");
  const pageBackground = normalizeHex(branding.pageBackground, "#d8dde3");

  const checks = [
    makeCheck("text_on_surface", "Body text on surface", text, surface, 4.5),
    makeCheck("text_on_page", "Body text on page background", text, pageBackground, 4.5),
    makeCheck("primary_on_surface", "Primary surfaces against cards", primary, surface, 3),
    makeCheck("accent_on_surface", "Accent actions against cards", accent, surface, 3),
    makeCheck("secondary_on_page", "Sidebar chrome against page background", secondary, pageBackground, 3),
  ];

  const recommendations = checks.filter((check) => !check.passed).map((check) => {
    if (check.key === "accent_on_surface") {
      return "Darken the accent color or reserve it for filled buttons only.";
    }
    if (check.key === "primary_on_surface") {
      return "Increase contrast between primary and surface for navigation and key panels.";
    }
    return `Improve ${check.label.toLowerCase()} contrast to at least ${check.requiredRatio}:1.`;
  });

  const score = Math.round((checks.filter((check) => check.passed).length / checks.length) * 100);

  return {
    score,
    checks,
    recommendations,
  };
}

export function createThemeSuggestions(branding: TenantBrandingDefinition): ThemeTokenSuggestion[] {
  const primary = normalizeHex(branding.primaryColor, "#005292");
  const secondary = normalizeHex(branding.secondaryColor, "#003d6e");
  const accent = normalizeHex(branding.accentColor, "#f57d00");

  return [
    {
      key: "contrast_boost",
      label: "Contrast boost",
      description: "Sharper contrast for admin shells and dense operator views.",
      tokens: {
        primaryColor: mix(primary, "#0f172a", 0.18),
        secondaryColor: mix(secondary, "#020617", 0.28),
        accentColor: mix(accent, "#111827", 0.08),
        surfaceColor: "#ffffff",
        textColor: "#0f172a",
        pageBackground: mix(primary, "#ffffff", 0.9),
      },
    },
    {
      key: "calm_surface",
      label: "Calm surface",
      description: "Softer page backgrounds with steadier reading contrast for large form journeys.",
      tokens: {
        primaryColor: primary,
        secondaryColor: mix(primary, "#0f172a", 0.16),
        accentColor: accent,
        surfaceColor: "#fbfdff",
        textColor: "#111827",
        pageBackground: mix(primary, "#ffffff", 0.95),
      },
    },
    {
      key: "accent_forward",
      label: "Accent forward",
      description: "Pushes CTAs and workflow controls harder without breaking the established shell.",
      tokens: {
        primaryColor: mix(primary, accent, 0.08),
        secondaryColor: mix(secondary, accent, 0.06),
        accentColor: mix(accent, "#ffffff", 0.04),
        surfaceColor: "#ffffff",
        textColor: "#111827",
        pageBackground: mix(accent, "#ffffff", 0.93),
      },
    },
  ];
}
