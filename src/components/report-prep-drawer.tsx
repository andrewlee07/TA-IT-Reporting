"use client";

import type { ReportPrepCheck, ReportPrepReviewItem, ReportPrepSectionRollup, ReportPrepView } from "@/lib/reports/prep-center";

interface ReportPrepDrawerProps {
  activeTab: "readiness" | "rollover";
  isLoading: boolean;
  isOpen: boolean;
  isSaving: boolean;
  prep: ReportPrepView | null;
  onClose: () => void;
  onCopyPreviousSummary: () => void;
  onJump: (pageId: string, tabId: string | null) => void;
  onToggleAcknowledged: (checkId: string, nextAcknowledged: boolean) => void;
  onTabChange: (tab: "readiness" | "rollover") => void;
}

function ReadinessCheckCard({
  check,
  isSaving,
  onJump,
  onToggleAcknowledged,
}: {
  check: ReportPrepCheck;
  isSaving: boolean;
  onJump: (pageId: string, tabId: string | null) => void;
  onToggleAcknowledged: (checkId: string, nextAcknowledged: boolean) => void;
}) {
  return (
    <div className={`prep-check-card severity-${check.severity} ${check.acknowledged ? "is-reviewed" : ""}`}>
      <div className="prep-check-head">
        <div>
          <div className={`prep-severity-pill ${check.severity}`}>{check.severity}</div>
          <div className="prep-check-title">{check.title}</div>
        </div>
        <button className="prep-inline-link" onClick={() => onJump(check.pageId, check.tabId)} type="button">
          Open section
        </button>
      </div>
      <div className="prep-check-source">{check.source}</div>
      <div className="prep-check-reason">{check.reason}</div>
      {check.acknowledgeable ? (
        <div className="prep-check-actions">
          <button
            className="prep-ghost-btn"
            disabled={isSaving}
            onClick={() => onToggleAcknowledged(check.id, !check.acknowledged)}
            type="button"
          >
            {check.acknowledged ? "Remove review" : "Acknowledge"}
          </button>
        </div>
      ) : null}
    </div>
  );
}

function ReviewItem({
  item,
  onJump,
}: {
  item: ReportPrepReviewItem;
  onJump: (pageId: string, tabId: string | null) => void;
}) {
  return (
    <div className={`prep-queue-item priority-${item.priority}`}>
      <div>
        <div className="prep-queue-title">{item.label}</div>
        <div className="prep-queue-reason">{item.reason}</div>
      </div>
      <button className="prep-inline-link" onClick={() => onJump(item.pageId, item.tabId)} type="button">
        Open
      </button>
    </div>
  );
}

function SectionRollup({
  section,
  onJump,
}: {
  section: ReportPrepSectionRollup;
  onJump: (pageId: string, tabId: string | null) => void;
}) {
  return (
    <div className={`prep-section-card state-${section.state}`}>
      <div className="prep-section-head">
        <div>
          <div className="prep-section-title">{section.label}</div>
          <div className="prep-section-summary">{section.summary}</div>
        </div>
        <div className="prep-section-head-meta">
          <div className={`prep-state-pill ${section.state}`}>{section.state.replace("-", " ")}</div>
          <button className="prep-inline-link" onClick={() => onJump(section.pageId, section.tabId)} type="button">
            Open
          </button>
        </div>
      </div>
      <div className="prep-section-narrative">
        <span>Narrative</span>
        <strong>{section.narrativeLabel}</strong>
      </div>
      {section.metrics.length > 0 ? (
        <div className="prep-metric-grid">
          {section.metrics.map((metric) => (
            <div className={`prep-metric-card ${metric.changed ? "changed" : ""}`} key={`${section.id}-${metric.label}`}>
              <div className="prep-metric-label">{metric.label}</div>
              <div className="prep-metric-values">
                <span>{metric.currentValue}</span>
                <span>{metric.previousValue}</span>
              </div>
              <div className="prep-metric-delta">{metric.deltaLabel}</div>
            </div>
          ))}
        </div>
      ) : null}
    </div>
  );
}

export function ReportPrepDrawer({
  activeTab,
  isLoading,
  isOpen,
  isSaving,
  prep,
  onClose,
  onCopyPreviousSummary,
  onJump,
  onToggleAcknowledged,
  onTabChange,
}: ReportPrepDrawerProps) {
  if (!isOpen) {
    return null;
  }

  const summary = prep?.readiness.summary;
  const previousSummary = prep?.rollover.previousExecSummary;
  const readOnly = prep?.mode === "demo-readonly";

  return (
    <div className="prep-drawer-shell" role="presentation">
      <button aria-label="Close readiness center" className="prep-drawer-backdrop" onClick={onClose} type="button" />
      <aside aria-label="Readiness and rollover center" className="prep-drawer" role="dialog">
        <div className="prep-drawer-header">
          <div>
            <div className="prep-drawer-kicker">Author workspace</div>
            <h2 className="prep-drawer-title">Readiness &amp; Rollover Center</h2>
            <div className="prep-drawer-sub">
              {prep ? `${prep.reportingMonthLabel}${prep.previousMonthLabel ? ` · comparing to ${prep.previousMonthLabel}` : ""}` : "Loading month readiness"}
            </div>
          </div>
          <button aria-label="Close readiness center" className="prep-drawer-close" onClick={onClose} type="button">
            <svg fill="none" viewBox="0 0 16 16">
              <path d="M4 4l8 8M12 4 4 12" stroke="currentColor" strokeLinecap="round" strokeWidth="1.8" />
            </svg>
          </button>
        </div>

        <div className="prep-tab-row">
          <button
            className={`prep-tab-btn ${activeTab === "readiness" ? "active" : ""}`}
            onClick={() => onTabChange("readiness")}
            type="button"
          >
            Readiness
          </button>
          <button
            className={`prep-tab-btn ${activeTab === "rollover" ? "active" : ""}`}
            onClick={() => onTabChange("rollover")}
            type="button"
          >
            Rollover
          </button>
        </div>

        <div className="prep-drawer-body">
          {isLoading || !prep ? (
            <div className="prep-empty-state">
              <div className="prep-empty-title">Loading preparation data</div>
              <div className="prep-empty-copy">Review status, warnings, and month-to-month changes are being assembled for the active report.</div>
            </div>
          ) : activeTab === "readiness" ? (
            <>
              <div className="prep-summary-card">
                <div className="prep-summary-head">
                  <div className={`prep-readiness-pill ${summary?.status ?? "needs-attention"}`}>
                    {summary?.status === "ready" ? "Ready" : "Needs attention"}
                  </div>
                  {readOnly ? <div className="prep-readonly-pill">Demo · read only</div> : null}
                </div>
                <div className="prep-summary-stats">
                  <div className="prep-stat">
                    <strong>{summary?.blockingCount ?? 0}</strong>
                    <span>Blocking</span>
                  </div>
                  <div className="prep-stat">
                    <strong>{summary?.warningCount ?? 0}</strong>
                    <span>Warnings</span>
                  </div>
                  <div className="prep-stat">
                    <strong>{summary?.acknowledgedCount ?? 0}</strong>
                    <span>Reviewed</span>
                  </div>
                </div>
              </div>

              {prep.readiness.checks.length > 0 ? (
                <div className="prep-stack">
                  <div className="prep-section-label">Needs attention</div>
                  {prep.readiness.checks.map((check) => (
                    <ReadinessCheckCard
                      check={check}
                      isSaving={isSaving || readOnly}
                      key={check.id}
                      onJump={onJump}
                      onToggleAcknowledged={onToggleAcknowledged}
                    />
                  ))}
                </div>
              ) : (
                <div className="prep-empty-state compact">
                  <div className="prep-empty-title">No open blockers or warnings</div>
                  <div className="prep-empty-copy">This month’s pack is ready from a prep-check perspective.</div>
                </div>
              )}

              {prep.readiness.reviewedChecks.length > 0 ? (
                <div className="prep-stack">
                  <div className="prep-section-label">Reviewed warnings</div>
                  {prep.readiness.reviewedChecks.map((check) => (
                    <ReadinessCheckCard
                      check={check}
                      isSaving={isSaving || readOnly}
                      key={check.id}
                      onJump={onJump}
                      onToggleAcknowledged={onToggleAcknowledged}
                    />
                  ))}
                </div>
              ) : null}
            </>
          ) : !prep.rollover.previousMonth ? (
            <div className="prep-empty-state">
              <div className="prep-empty-title">No previous month in this workbook</div>
              <div className="prep-empty-copy">Add another reporting month to the workbook and rollover comparisons will appear here automatically.</div>
            </div>
          ) : (
            <>
              {previousSummary?.available ? (
                <div className="prep-helper-card">
                  <div className="prep-helper-kicker">Exec Summary helper</div>
                  <div className="prep-helper-title">Use the previous month summary as a draft</div>
                  <div className="prep-helper-copy">
                    {previousSummary.monthLabel} has saved summary content. You can copy it into the current month editor and then refine it.
                  </div>
                  <div className="prep-helper-excerpt">{previousSummary.excerpt || "A saved executive summary is available."}</div>
                  <button
                    className="prep-primary-btn"
                    disabled={readOnly}
                    onClick={onCopyPreviousSummary}
                    type="button"
                  >
                    {readOnly ? "Read-only demo" : "Copy previous month summary"}
                  </button>
                </div>
              ) : null}

              <div className="prep-stack">
                <div className="prep-section-label">Review queue</div>
                {prep.rollover.reviewQueue.length > 0 ? (
                  prep.rollover.reviewQueue.map((item) => <ReviewItem item={item} key={item.id} onJump={onJump} />)
                ) : (
                  <div className="prep-empty-state compact">
                    <div className="prep-empty-title">No standout rollover items</div>
                    <div className="prep-empty-copy">The current month is broadly aligned with the previous month and does not surface obvious review hotspots.</div>
                  </div>
                )}
              </div>

              <div className="prep-stack">
                <div className="prep-section-label">Section comparison</div>
                {prep.rollover.sections.map((section) => (
                  <SectionRollup key={section.id} onJump={onJump} section={section} />
                ))}
              </div>
            </>
          )}
        </div>
      </aside>
    </div>
  );
}
