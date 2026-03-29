export const ANNOTATION_FONT_SIZES = [14, 16, 20, 24, 28] as const;
export const ANNOTATION_TEXT_COLORS = ["#0f172a", "#005292", "#f57d00", "#219d98", "#b42318"] as const;
export const ANNOTATION_SURFACE_COLORS = ["#ffffff", "#fef3e7", "#e7f4f4", "#ebf2f9", "#fef2f2"] as const;
export const ANNOTATION_BORDER_COLORS = ["#d1d5db", "#005292", "#f57d00", "#219d98", "#b42318"] as const;

export type AnnotationFontSize = (typeof ANNOTATION_FONT_SIZES)[number];
export type AnnotationTextColor = (typeof ANNOTATION_TEXT_COLORS)[number];
export type AnnotationSurfaceColor = (typeof ANNOTATION_SURFACE_COLORS)[number];
export type AnnotationBorderColor = (typeof ANNOTATION_BORDER_COLORS)[number];

export type ReportAnnotationType = "bubble" | "text";
export type ReportAnnotationTool = "select" | ReportAnnotationType;

export interface ReportAnnotationStyle {
  fontSize: AnnotationFontSize;
  bold: boolean;
  italic: boolean;
  textColor: AnnotationTextColor;
  fillColor: AnnotationSurfaceColor;
  borderColor: AnnotationBorderColor;
}

export interface ReportAnnotationPoint {
  x: number;
  y: number;
}

export interface ReportAnnotation {
  id: string;
  slideId: string;
  type: ReportAnnotationType;
  x: number;
  y: number;
  width: number;
  height: number;
  text: string;
  zIndex: number;
  style: ReportAnnotationStyle;
  tailAnchor: ReportAnnotationPoint | null;
}

export interface ReportAnnotationsState {
  reportId: string;
  reportingMonth: string;
  revisionId: string | null;
  annotations: ReportAnnotation[];
  updatedAt: string | null;
}

export const DEFAULT_ANNOTATION_STYLE: ReportAnnotationStyle = {
  fontSize: 16,
  bold: false,
  italic: false,
  textColor: "#0f172a",
  fillColor: "#ffffff",
  borderColor: "#005292",
};

export function createEmptyReportAnnotationsState(reportId: string, reportingMonth: string): ReportAnnotationsState {
  return {
    reportId,
    reportingMonth,
    revisionId: null,
    annotations: [],
    updatedAt: null,
  };
}
