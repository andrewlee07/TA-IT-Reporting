"use client";

import { useMemo, useState, useTransition } from "react";

import { getThemeCssVariables } from "@/lib/platform/theme";
import type { FormDefinition, PlatformManifest } from "@/lib/platform/types";

import styles from "./platform-shell.module.css";

function fieldInputType(type: FormDefinition["fields"][number]["type"]): string {
  if (type === "number" || type === "currency") {
    return "number";
  }

  if (type === "date") {
    return "date";
  }

  if (type === "datetime") {
    return "datetime-local";
  }

  return "text";
}

export function PlatformPublicForm({
  manifest,
  form,
  tenantSlug,
  mode = "public",
}: {
  manifest: PlatformManifest;
  form: FormDefinition;
  tenantSlug: string;
  mode?: "public" | "draft-preview";
}) {
  const [currentStepIndex, setCurrentStepIndex] = useState(0);
  const [draft, setDraft] = useState<Record<string, unknown>>({});
  const [message, setMessage] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isPending, startTransition] = useTransition();
  const steps = form.steps.length > 0 ? form.steps : [{ id: "single", key: "single", title: form.title, fieldKeys: form.fields.map((field) => field.key) }];
  const currentStep = steps[currentStepIndex]!;
  const visibleFields = useMemo(
    () => form.fields.filter((field) => currentStep.fieldKeys.includes(field.key)),
    [currentStep.fieldKeys, form.fields],
  );

  async function submit(status: "draft" | "submitted"): Promise<void> {
    setError(null);
    setMessage(null);

    const response = await fetch(`/api/platform/tenants/${tenantSlug}/forms/${form.key}/submissions`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
      },
      body: JSON.stringify({
        data: draft,
        status,
      }),
    });
    const payload = (await response.json()) as { error?: string };
    if (!response.ok) {
      throw new Error(payload.error ?? "Form submission failed.");
    }

    setMessage(status === "draft" ? "Draft saved." : form.successMessage);
    if (status === "submitted") {
      setDraft({});
      setCurrentStepIndex(0);
    }
  }

  return (
    <div className={styles.previewShell} style={getThemeCssVariables(manifest)}>
      <div className={styles.previewBanner}>
        <div>
          <p className={styles.cardEyebrow}>{manifest.appShell.productName}</p>
          <h2>{form.title}</h2>
          <p className={styles.helperCopy}>{form.description ?? "Complete the form below to submit your request."}</p>
        </div>
        <div className={styles.inlineList}>
          <span className={styles.inlineTag}>{form.deliveryMode}</span>
          <span className={styles.inlineTag}>{steps.length} step{steps.length === 1 ? "" : "s"}</span>
          {mode === "draft-preview" ? <span className={styles.inlineTag}>draft preview</span> : null}
        </div>
      </div>

      <section className={styles.panelWide}>
        {message ? <div className={styles.successBanner}>{message}</div> : null}
        {error ? <div className={styles.errorBanner}>{error}</div> : null}
        {mode === "draft-preview" ? (
          <div className={styles.sidebarPanel}>
            Draft preview only. Publish this tenant version before sharing the live form or capturing real submissions.
          </div>
        ) : null}

        <div className={styles.sectionHeader}>
          <div>
            <p className={styles.cardEyebrow}>Step {currentStepIndex + 1}</p>
            <h3>{currentStep.title}</h3>
            {currentStep.description ? <p className={styles.helperCopy}>{currentStep.description}</p> : null}
          </div>
          <span className={styles.badge}>
            {currentStepIndex + 1} / {steps.length}
          </span>
        </div>

        <div className={styles.formGrid}>
          {visibleFields.map((field) => (
            <label className={styles.formFieldSpan} key={field.id}>
              <span>{field.label}</span>
              {field.type === "long_text" ? (
                <textarea
                  className={styles.textarea}
                  onChange={(event) => setDraft((current) => ({ ...current, [field.key]: event.target.value }))}
                  placeholder={field.placeholder}
                  value={String(draft[field.key] ?? field.defaultValue ?? "")}
                />
              ) : field.type === "boolean" ? (
                <label className={styles.checkboxField}>
                  <input
                    checked={Boolean(draft[field.key] ?? field.defaultValue ?? false)}
                    onChange={(event) => setDraft((current) => ({ ...current, [field.key]: event.target.checked }))}
                    type="checkbox"
                  />
                  <span>{field.helpText ?? field.tooltip ?? "Toggle this setting."}</span>
                </label>
              ) : field.type === "select" ? (
                <select
                  className={styles.select}
                  onChange={(event) => setDraft((current) => ({ ...current, [field.key]: event.target.value }))}
                  value={String(draft[field.key] ?? field.defaultValue ?? "")}
                >
                  <option value="">Choose an option</option>
                  {field.options?.map((option) => (
                    <option key={option} value={option}>
                      {option}
                    </option>
                  ))}
                </select>
              ) : (
                <input
                  className={styles.input}
                  onChange={(event) => setDraft((current) => ({ ...current, [field.key]: event.target.value }))}
                  placeholder={field.placeholder}
                  type={fieldInputType(field.type)}
                  value={String(draft[field.key] ?? field.defaultValue ?? "")}
                />
              )}
              {field.helpText || field.tooltip ? <small className={styles.metricMeta}>{field.helpText ?? field.tooltip}</small> : null}
            </label>
          ))}
        </div>

        <div className={styles.actionsRow}>
          <div className={styles.inlineList}>
            <button
              className={styles.secondaryButton}
              disabled={currentStepIndex === 0 || isPending}
              onClick={() => setCurrentStepIndex((current) => Math.max(0, current - 1))}
              type="button"
            >
              Previous
            </button>
            <button
              className={styles.secondaryButton}
              disabled={currentStepIndex >= steps.length - 1 || isPending}
              onClick={() => setCurrentStepIndex((current) => Math.min(steps.length - 1, current + 1))}
              type="button"
            >
              Next
            </button>
          </div>
          <div className={styles.inlineList}>
            {form.saveAndResume ? (
              <button
                className={styles.secondaryButton}
                disabled={isPending || mode === "draft-preview"}
                onClick={() =>
                  startTransition(() => {
                    void submit("draft").catch((caughtError) => {
                      setError(caughtError instanceof Error ? caughtError.message : "Failed to save draft.");
                    });
                  })
                }
                type="button"
              >
                Save draft
              </button>
            ) : null}
            <button
              className={styles.primaryButton}
              disabled={isPending || mode === "draft-preview"}
              onClick={() =>
                startTransition(() => {
                  void submit("submitted").catch((caughtError) => {
                    setError(caughtError instanceof Error ? caughtError.message : "Failed to submit form.");
                  });
                })
              }
              type="button"
            >
              {isPending ? "Working..." : form.submitLabel}
            </button>
          </div>
        </div>
      </section>
    </div>
  );
}
