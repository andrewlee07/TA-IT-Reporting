import { z } from "zod";

import {
  ANNOTATION_BORDER_COLORS,
  ANNOTATION_FONT_SIZES,
  ANNOTATION_SURFACE_COLORS,
  ANNOTATION_TEXT_COLORS,
  createEmptyReportAnnotationsState,
  type ReportAnnotation,
  type ReportAnnotationPoint,
  type ReportAnnotationStyle,
  type ReportAnnotationsState,
} from "@/lib/annotations/types";

const unitNumberSchema = z.number().finite().min(0).max(1);

export const reportAnnotationPointSchema = z.object({
  x: unitNumberSchema,
  y: unitNumberSchema,
});

export const reportAnnotationStyleSchema = z.object({
  fontSize: z
    .number()
    .int()
    .refine((value): value is (typeof ANNOTATION_FONT_SIZES)[number] => ANNOTATION_FONT_SIZES.includes(value as (typeof ANNOTATION_FONT_SIZES)[number])),
  bold: z.boolean(),
  italic: z.boolean(),
  textColor: z.enum(ANNOTATION_TEXT_COLORS),
  fillColor: z.enum(ANNOTATION_SURFACE_COLORS),
  borderColor: z.enum(ANNOTATION_BORDER_COLORS),
});

export const reportAnnotationSchema = z.object({
  id: z.string().min(1),
  slideId: z.string().min(1),
  type: z.enum(["bubble", "text"]),
  x: unitNumberSchema,
  y: unitNumberSchema,
  width: unitNumberSchema,
  height: unitNumberSchema,
  text: z.string().max(4000),
  zIndex: z.number().int().min(0).max(9999),
  style: reportAnnotationStyleSchema,
  tailAnchor: reportAnnotationPointSchema.nullable(),
});

export const reportAnnotationsStateSchema = z.object({
  reportId: z.string().min(1),
  reportingMonth: z.string().min(1),
  revisionId: z.string().nullable(),
  annotations: z.array(reportAnnotationSchema),
  updatedAt: z.string().nullable(),
});

export const reportAnnotationsSaveSchema = z.object({
  baseRevisionId: z.string().nullable(),
  annotations: z.array(reportAnnotationSchema),
});

function clampUnit(value: number): number {
  return Math.min(1, Math.max(0, value));
}

function sanitizeText(value: string): string {
  return value.replace(/\r\n/g, "\n").slice(0, 4000);
}

function normalizePoint(point: ReportAnnotationPoint | null): ReportAnnotationPoint | null {
  if (!point) {
    return null;
  }

  return {
    x: clampUnit(point.x),
    y: clampUnit(point.y),
  };
}

function normalizeStyle(style: ReportAnnotationStyle): ReportAnnotationStyle {
  const fontSize = ANNOTATION_FONT_SIZES.includes(style.fontSize) ? style.fontSize : ANNOTATION_FONT_SIZES[1];
  const textColor = ANNOTATION_TEXT_COLORS.includes(style.textColor) ? style.textColor : ANNOTATION_TEXT_COLORS[0];
  const fillColor = ANNOTATION_SURFACE_COLORS.includes(style.fillColor) ? style.fillColor : ANNOTATION_SURFACE_COLORS[0];
  const borderColor = ANNOTATION_BORDER_COLORS.includes(style.borderColor) ? style.borderColor : ANNOTATION_BORDER_COLORS[1];

  return {
    fontSize,
    bold: Boolean(style.bold),
    italic: Boolean(style.italic),
    textColor,
    fillColor,
    borderColor,
  };
}

export function normalizeReportAnnotation(annotation: ReportAnnotation): ReportAnnotation {
  return {
    ...annotation,
    x: clampUnit(annotation.x),
    y: clampUnit(annotation.y),
    width: clampUnit(annotation.width),
    height: clampUnit(annotation.height),
    text: sanitizeText(annotation.text),
    zIndex: Math.max(0, Math.min(9999, Math.round(annotation.zIndex))),
    style: normalizeStyle(annotation.style),
    tailAnchor: annotation.type === "bubble" ? normalizePoint(annotation.tailAnchor) : null,
  };
}

export function normalizeReportAnnotations(annotations: ReportAnnotation[]): ReportAnnotation[] {
  return annotations
    .map(normalizeReportAnnotation)
    .sort((left, right) => left.zIndex - right.zIndex || left.id.localeCompare(right.id));
}

export function parseReportAnnotationsState(
  reportId: string,
  reportingMonth: string,
  value: unknown,
): ReportAnnotationsState {
  if (!value) {
    return createEmptyReportAnnotationsState(reportId, reportingMonth);
  }

  const parsed = reportAnnotationsStateSchema.safeParse(value);

  if (!parsed.success) {
    return createEmptyReportAnnotationsState(reportId, reportingMonth);
  }

  return {
    reportId,
    reportingMonth,
    revisionId: parsed.data.revisionId,
    annotations: normalizeReportAnnotations(parsed.data.annotations),
    updatedAt: parsed.data.updatedAt,
  };
}
