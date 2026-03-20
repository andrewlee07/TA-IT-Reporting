"use client";

import { useRouter } from "next/navigation";
import { useState } from "react";

export function UploadForm() {
  const router = useRouter();
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [issues, setIssues] = useState<string[]>([]);

  async function handleSubmit(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setIsSubmitting(true);
    setError(null);
    setIssues([]);

    const formData = new FormData(event.currentTarget);

    try {
      const response = await fetch("/api/reports", {
        method: "POST",
        body: formData,
      });

      const payload = (await response.json()) as {
        error?: string;
        issues?: string[];
        report?: { id: string };
      };

      if (!response.ok || !payload.report) {
        setError(payload.error ?? "Upload failed.");
        setIssues(payload.issues ?? []);
        return;
      }

      router.push(`/reports/${payload.report.id}`);
      router.refresh();
    } catch (caughtError) {
      setError(caughtError instanceof Error ? caughtError.message : "Upload failed.");
    } finally {
      setIsSubmitting(false);
    }
  }

  return (
    <form className="upload-form" onSubmit={handleSubmit}>
      <label className="upload-input">
        <span>Excel workbook</span>
        <input accept=".xlsx" name="workbook" required type="file" />
      </label>
      <button className="app-button app-button-primary" disabled={isSubmitting} type="submit">
        {isSubmitting ? "Validating and saving..." : "Upload workbook"}
      </button>
      {error ? <p className="form-error">{error}</p> : null}
      {issues.length > 0 ? (
        <ul className="form-issues">
          {issues.map((issue) => (
            <li key={issue}>{issue}</li>
          ))}
        </ul>
      ) : null}
    </form>
  );
}
