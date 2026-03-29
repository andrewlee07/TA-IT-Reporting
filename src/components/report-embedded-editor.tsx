"use client";

import { useMemo } from "react";

import styles from "@/components/report-embedded-editor.module.css";
import type { EditableReportDraft, SectionId } from "@/lib/drafts/types";
import { buildFieldDefinitions, getEditorSectionConfig, type CollectionConfig, type FieldDefinition } from "@/lib/editor/config";

type SaveState = "idle" | "dirty" | "saving" | "saved" | "conflict" | "error";

interface ReportEmbeddedEditorProps {
  draft: EditableReportDraft;
  sectionId: SectionId;
  selectedMonth: string;
  saveState: SaveState;
  saveMessage: string;
  onDraftChange: (updater: (current: EditableReportDraft) => EditableReportDraft) => void;
  onMarkDirty: (sectionId: SectionId) => void;
  onSelectedMonthChange: (month: string) => void;
}

function formatDateLabel(value: string | null): string {
  if (!value) {
    return "Never";
  }

  return new Date(value).toLocaleString("en-GB", {
    dateStyle: "medium",
    timeStyle: "short",
  });
}

function formatPresenceLabel(lastSeenAt: string): string {
  const deltaSeconds = Math.max(0, Math.floor((Date.now() - Date.parse(lastSeenAt)) / 1000));
  if (deltaSeconds < 60) {
    return "active now";
  }

  const minutes = Math.floor(deltaSeconds / 60);
  return `${minutes}m ago`;
}

function inferRowKey(row: Record<string, unknown>, index: number): string {
  return (
    String(row.reportingMonth ?? "") +
    String(row.ticketId ?? row.projectName ?? row.workstreamName ?? row.officeName ?? row.assetType ?? row.budgetLine ?? row.riskIssue ?? row.headline ?? index)
  );
}

function renderValue(value: unknown): string {
  if (typeof value === "boolean") {
    return value ? "Yes" : "No";
  }

  if (value === null || value === undefined || value === "") {
    return "—";
  }

  return String(value);
}

function inferEntryTitle(collection: CollectionConfig, row: Record<string, unknown>, index: number, isPlaceholder: boolean): string {
  const preferredKeys = [
    "serviceName",
    "officeName",
    "assetType",
    "ticketId",
    "projectName",
    "workstreamName",
    "budgetLine",
    "riskIssue",
    "headline",
    "initiative",
    "entityName",
    "lane",
  ];

  for (const key of preferredKeys) {
    const value = row[key];
    if (typeof value === "string" && value.trim().length > 0) {
      return value;
    }
  }

  if (isPlaceholder) {
    return `New ${collection.label}`;
  }

  return `${collection.label} ${index + 1}`;
}

function RowFields({
  row,
  fields,
  disabled,
  onChange,
}: {
  row: Record<string, unknown>;
  fields: FieldDefinition[];
  disabled?: boolean;
  onChange: (nextRow: Record<string, unknown>) => void;
}) {
  return (
    <div className={styles.fieldGrid}>
      {fields.map((field) => {
        const value = row[field.key];
        const inputId = `embedded-field-${field.key}`;

        if (field.type === "checkbox") {
          return (
            <label className={styles.checkboxField} htmlFor={inputId} key={field.key}>
              <input
                checked={Boolean(value)}
                disabled={disabled}
                id={inputId}
                onChange={(event) => onChange({ ...row, [field.key]: event.target.checked })}
                type="checkbox"
              />
              <span>{field.label}</span>
            </label>
          );
        }

        if (field.type === "textarea") {
          return (
            <label className={styles.field} htmlFor={inputId} key={field.key}>
              <span>{field.label}</span>
              <textarea
                disabled={disabled}
                id={inputId}
                onChange={(event) => onChange({ ...row, [field.key]: event.target.value })}
                rows={4}
                value={String(value ?? "")}
              />
            </label>
          );
        }

        return (
          <label className={styles.field} htmlFor={inputId} key={field.key}>
            <span>{field.label}</span>
            <input
              disabled={disabled}
              id={inputId}
              onChange={(event) =>
                onChange({
                  ...row,
                  [field.key]:
                    field.type === "number"
                      ? event.target.value === ""
                        ? 0
                        : Number(event.target.value)
                      : event.target.value,
                })
              }
              step={field.step}
              type={field.type === "number" ? "number" : field.type === "date" ? "date" : "text"}
              value={value === null || value === undefined ? "" : String(value)}
            />
          </label>
        );
      })}
    </div>
  );
}

export function ReportEmbeddedEditor({
  draft,
  sectionId,
  selectedMonth,
  saveState,
  saveMessage,
  onDraftChange,
  onMarkDirty,
  onSelectedMonthChange,
}: ReportEmbeddedEditorProps) {
  const section = useMemo(() => getEditorSectionConfig(sectionId), [sectionId]);
  const selectedPresence = useMemo(
    () => draft.activePresence.filter((entry) => entry.reportingMonth === selectedMonth),
    [draft.activePresence, selectedMonth],
  );

  function updateCollectionRows(collection: CollectionConfig, rows: Array<Record<string, unknown>>) {
    if (!collection.setRows) {
      return;
    }

    onDraftChange((current) => ({
      ...current,
      snapshot: collection.setRows!(current.snapshot, selectedMonth, rows),
    }));
    onMarkDirty(sectionId);
  }

  return (
    <div className={styles.stack}>
      <div className={styles.statusBar}>
        <span className={styles.statusPill} data-tone={saveState}>
          {saveMessage}
        </span>
        <span className={styles.statusMeta}>Revision {draft.manifest.currentRevision.revisionNumber}</span>
        <span className={styles.statusMeta}>Last updated {formatDateLabel(draft.manifest.currentRevision.updatedAt)}</span>
        <span className={styles.statusMeta}>Workbook sync: {draft.manifest.artifactSyncStatus.state}</span>
        {selectedPresence.map((presence) => (
          <span className={styles.presencePill} key={`${presence.user.id}-${presence.reportingMonth}`}>
            {presence.user.name} · {formatPresenceLabel(presence.lastSeenAt)}
          </span>
        ))}
      </div>

      {sectionId === "overview-setup" ? (
        <div className={`block ${styles.editorBlock}`}>
          <div className="bh">
            <div>
              <div className="bh-title">Report Metadata</div>
              <div className="bh-sub">Keep the report identity, month scope and naming aligned with the workbook.</div>
            </div>
          </div>
          <div className="bb">
            <div className={styles.metaGrid}>
              <label className={styles.field}>
                <span>Title</span>
                <input
                  onChange={(event) => {
                    const value = event.target.value;
                    onDraftChange((current) => ({
                      ...current,
                      manifest: { ...current.manifest, title: value },
                    }));
                    onMarkDirty("overview-setup");
                  }}
                  type="text"
                  value={draft.manifest.title}
                />
              </label>
              <label className={styles.field}>
                <span>Report Series Key</span>
                <input
                  onChange={(event) => {
                    const value = event.target.value;
                    onDraftChange((current) => ({
                      ...current,
                      manifest: { ...current.manifest, reportSeriesKey: value },
                    }));
                    onMarkDirty("overview-setup");
                  }}
                  type="text"
                  value={draft.manifest.reportSeriesKey}
                />
              </label>
              <label className={styles.field}>
                <span>Current Month</span>
                <select
                  onChange={(event) => {
                    const value = event.target.value;
                    onDraftChange((current) => ({
                      ...current,
                      snapshot: {
                        ...current.snapshot,
                        currentMonth: value,
                      },
                    }));
                    onSelectedMonthChange(value);
                    onMarkDirty("overview-setup");
                  }}
                  value={draft.snapshot.currentMonth}
                >
                  {draft.snapshot.availableMonths.map((month) => (
                    <option key={month} value={month}>
                      {month}
                    </option>
                  ))}
                </select>
              </label>
            </div>
          </div>
        </div>
      ) : null}

      {section.collections.map((collection) => {
        const rows = collection.getRows(draft.snapshot, selectedMonth);
        const fallbackRow = (collection.createRow?.(selectedMonth) ?? {}) as Record<string, unknown>;
        const collectionRows =
          rows.length > 0
            ? rows
            : collection.layout === "table" && collection.createRow
              ? [fallbackRow]
              : collection.layout === "readonly" && Object.keys(fallbackRow).length > 0
                ? [fallbackRow]
                : rows;
        const rowFields = buildFieldDefinitions((collectionRows[0] ?? fallbackRow) as Record<string, unknown>);

        return (
          <div className="block" key={collection.key}>
            <div className="bh">
              <div className={styles.collectionToolbarTitle}>
                <div className="bh-title">{collection.label}</div>
                <div className="bh-sub">{collection.description}</div>
              </div>
              {collection.createRow && collection.layout !== "readonly" ? (
                <button
                  className={styles.actionButton}
                  onClick={() => {
                    updateCollectionRows(collection, [...rows, collection.createRow!(selectedMonth)]);
                  }}
                  type="button"
                >
                  Add entry
                </button>
              ) : null}
            </div>
            <div className="bb">
              {collection.layout === "single" ? (
                <RowFields
                  fields={rowFields}
                  onChange={(nextRow) => updateCollectionRows(collection, [nextRow])}
                  row={(rows[0] ?? fallbackRow) as Record<string, unknown>}
                />
              ) : collection.layout === "readonly" ? (
                collectionRows.length > 0 ? (
                  <div className={styles.readonlyValueGrid}>
                    {collectionRows.map((row, rowIndex) => (
                      <div className={styles.readonlyCard} key={inferEntryTitle(collection, row, rowIndex, false)}>
                        <div className={styles.entryCardHeader}>
                          <div>
                            <h4 className={styles.entryCardTitle}>{inferEntryTitle(collection, row, rowIndex, false)}</h4>
                            <p className={styles.entryCardSub}>Calculated from the structured inputs on this report.</p>
                          </div>
                        </div>
                        <div className={styles.readonlyMetrics}>
                          {rowFields.map((field) => (
                            <div className={styles.readonlyMetric} key={field.key}>
                              <span>{field.label}</span>
                              <strong>{renderValue((row as Record<string, unknown>)[field.key])}</strong>
                            </div>
                          ))}
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className={styles.readonlyPanel}>
                    <h4>{collection.label}</h4>
                    <p>This calculated section will appear once the supporting source data has been entered.</p>
                  </div>
                )
              ) : (
                <div className={styles.entryCardStack}>
                  {collectionRows.map((row, rowIndex) => {
                    const isPlaceholder = rows.length === 0;

                    return (
                      <div className={styles.entryCard} key={inferRowKey(row, rowIndex)}>
                        <div className={styles.entryCardHeader}>
                          <div>
                            <h4 className={styles.entryCardTitle}>{inferEntryTitle(collection, row, rowIndex, isPlaceholder)}</h4>
                            <p className={styles.entryCardSub}>
                              {isPlaceholder
                                ? "Start entering the first structured record for this report section."
                                : `Structured ${collection.label.toLowerCase()} input tied directly to the workbook and report.`}
                            </p>
                          </div>
                          {!isPlaceholder ? (
                            <button
                              className={styles.inlineDanger}
                              onClick={() => {
                                updateCollectionRows(
                                  collection,
                                  rows.filter((_, index) => index !== rowIndex),
                                );
                              }}
                              type="button"
                            >
                              Remove
                            </button>
                          ) : null}
                        </div>
                        <RowFields
                          fields={buildFieldDefinitions(row)}
                          onChange={(nextRow) => {
                            if (rows.length === 0) {
                              updateCollectionRows(collection, [nextRow]);
                              return;
                            }

                            updateCollectionRows(
                              collection,
                              rows.map((entry, index) => (index === rowIndex ? nextRow : entry)),
                            );
                          }}
                          row={row}
                        />
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
}
