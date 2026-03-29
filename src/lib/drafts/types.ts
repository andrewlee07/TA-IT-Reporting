import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

export const EDITOR_SECTIONS = [
  "overview-setup",
  "availability-network",
  "support-operations",
  "security-assets",
  "change-delivery",
  "projects-roadmap",
  "finance-risks",
  "notes-narrative",
] as const;

export type SectionId = (typeof EDITOR_SECTIONS)[number];

export interface DraftUserRef {
  id: string;
  name: string;
  email: string | null;
}

export interface ArtifactSyncStatus {
  state: "pending" | "current" | "failed";
  revisionId: string | null;
  updatedAt: string | null;
  error: string | null;
  workbookObjectKey: string | null;
  jsonObjectKey: string | null;
}

export interface ReportRevisionMeta {
  revisionId: string;
  revisionNumber: number;
  updatedAt: string;
  updatedBy: DraftUserRef;
  changedSections: SectionId[];
}

export interface EditorPresence {
  reportId: string;
  reportingMonth: string;
  user: DraftUserRef;
  lastSeenAt: string;
}

export interface DraftManifest {
  reportId: string;
  title: string;
  originalFilename: string;
  reportSeriesKey: string;
  templateKey: string;
  templateVersion: number;
  currentMonth: string;
  availableMonths: string[];
  createdAt: string;
  updatedAt: string;
  workbookObjectKey: string;
  artifactSyncStatus: ArtifactSyncStatus;
  currentRevision: ReportRevisionMeta;
}

export interface EditableReportDraft {
  manifest: DraftManifest;
  snapshot: NormalizedReportSnapshot;
  activePresence: EditorPresence[];
}

export interface DraftRevisionRecord {
  meta: ReportRevisionMeta;
  snapshot: NormalizedReportSnapshot;
}
