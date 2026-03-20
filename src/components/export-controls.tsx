"use client";

import { useMemo, useState } from "react";

import type { ReportBlockDefinition } from "@/lib/report/blocks";

interface ExportControlsProps {
  reportId: string;
  reportTitle: string;
  selectedMonth: string;
  selectedPageId: string;
  blocks: ReportBlockDefinition[];
}

function parseFilename(disposition: string | null, fallback: string): string {
  if (!disposition) {
    return fallback;
  }

  const match = disposition.match(/filename=\"?([^\";]+)\"?/i);
  return match?.[1] ?? fallback;
}

export function ExportControls({ reportId, reportTitle, selectedMonth, selectedPageId, blocks }: ExportControlsProps) {
  const [busyKey, setBusyKey] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [blockId, setBlockId] = useState(blocks[0]?.id ?? "");

  const blockOptions = useMemo(() => blocks, [blocks]);

  async function downloadArtifact(payload: Record<string, string>) {
    setBusyKey(payload.exportType);
    setError(null);

    try {
      const response = await fetch(`/api/reports/${reportId}/exports`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      if (!response.ok) {
        const body = (await response.json()) as { error?: string };
        throw new Error(body.error ?? "Export failed.");
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = parseFilename(response.headers.get("content-disposition"), `${reportTitle}-${payload.exportType}`);
      anchor.click();
      URL.revokeObjectURL(url);
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Export failed.");
    } finally {
      setBusyKey(null);
    }
  }

  return (
    <div className="export-controls">
      <div className="export-grid">
        <button
          className="app-button app-button-secondary"
          disabled={busyKey !== null}
          onClick={() => void downloadArtifact({ exportType: "page-png", month: selectedMonth, pageId: selectedPageId })}
          type="button"
        >
          {busyKey === "page-png" ? "Rendering page..." : "Download page PNG"}
        </button>
        <button
          className="app-button app-button-secondary"
          disabled={busyKey !== null}
          onClick={() => void downloadArtifact({ exportType: "full-pdf", month: selectedMonth, pageId: selectedPageId })}
          type="button"
        >
          {busyKey === "full-pdf" ? "Rendering PDF..." : "Download full PDF"}
        </button>
      </div>

      <div className="export-block-row">
        <label className="export-select">
          <span>Block export</span>
          <select disabled={busyKey !== null || blockOptions.length === 0} onChange={(event) => setBlockId(event.target.value)} value={blockId}>
            {blockOptions.length === 0 ? <option value="">No export blocks on this page</option> : null}
            {blockOptions.map((block) => (
              <option key={block.id} value={block.id}>
                {block.label}
              </option>
            ))}
          </select>
        </label>
        <button
          className="app-button app-button-secondary"
          disabled={busyKey !== null || !blockId}
          onClick={() => void downloadArtifact({ exportType: "block-png", month: selectedMonth, pageId: selectedPageId, blockId })}
          type="button"
        >
          {busyKey === "block-png" ? "Rendering block..." : "Download block PNG"}
        </button>
      </div>

      {error ? <p className="form-error">{error}</p> : null}
    </div>
  );
}
