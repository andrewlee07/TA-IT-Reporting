"use client";

import { useEffect, useMemo, useState, useSyncExternalStore, useTransition } from "react";

import { getThemeCssVariables } from "@/lib/platform/theme";
import { evaluateRuleAsBoolean } from "@/lib/platform/rule-engine";
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
  const [currentStepIndex, setCurrentStepIndex] = useState<number | null>(null);
  const [draft, setDraft] = useState<Record<string, unknown> | null>(null);
  const [message, setMessage] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [fieldErrors, setFieldErrors] = useState<Record<string, string>>({});
  const [isPending, startTransition] = useTransition();
  const storageKey = `ta-platform-form:${tenantSlug}:${form.key}`;
  const steps = useMemo(
    () =>
      form.steps.length > 0
        ? form.steps
        : [
            {
              id: "single",
              key: "single",
              title: form.title,
              fieldKeys: form.fields.map((field) => field.key),
            },
          ],
    [form.fields, form.steps, form.title],
  );
  const resumeSnapshot = useSyncExternalStore(
    () => () => undefined,
    () => {
      if (!form.saveAndResume) {
        return null;
      }

      try {
        return window.localStorage.getItem(storageKey);
      } catch {
        return null;
      }
    },
    () => null,
  );
  const parsedResumeSnapshot = useMemo(() => {
    if (!form.saveAndResume || !resumeSnapshot) {
      return null;
    }

    try {
      const payload = JSON.parse(resumeSnapshot) as { draft?: Record<string, unknown>; currentStepIndex?: number };
      return {
        draft:
          payload.draft &&
          typeof payload.draft === "object" &&
          !Array.isArray(payload.draft)
            ? payload.draft
            : {},
        currentStepIndex:
          typeof payload.currentStepIndex === "number" &&
          Number.isInteger(payload.currentStepIndex) &&
          payload.currentStepIndex >= 0
            ? payload.currentStepIndex
            : 0,
      };
    } catch {
      return null;
    }
  }, [form.saveAndResume, resumeSnapshot]);
  const effectiveDraft = useMemo(() => draft ?? parsedResumeSnapshot?.draft ?? {}, [draft, parsedResumeSnapshot]);
  const effectiveStepIndex = currentStepIndex ?? parsedResumeSnapshot?.currentStepIndex ?? 0;
  const visibleSteps = useMemo(
    () =>
      steps.filter((step) => {
        if (!step.visibilityRule) {
          return true;
        }
        return evaluateRuleAsBoolean(step.visibilityRule.rule ?? { mode: "text", expression: step.visibilityRule.expression }, effectiveDraft);
      }),
    [effectiveDraft, steps],
  );
  const resolvedSteps = visibleSteps.length > 0 ? visibleSteps : steps;
  const activeStepIndex = Math.min(effectiveStepIndex, Math.max(0, resolvedSteps.length - 1));
  const currentStep = resolvedSteps[activeStepIndex]!;
  const visibleFields = useMemo(
    () => form.fields.filter((field) => currentStep.fieldKeys.includes(field.key)),
    [currentStep.fieldKeys, form.fields],
  );

  useEffect(() => {
    if (!form.saveAndResume) {
      return;
    }

    try {
      window.localStorage.setItem(
        storageKey,
        JSON.stringify({
          draft: effectiveDraft,
          currentStepIndex: activeStepIndex,
        }),
      );
    } catch {
      // Resume storage is optional.
    }
  }, [activeStepIndex, effectiveDraft, form.saveAndResume, storageKey]);

  function validateCurrentStep(): boolean {
    const nextErrors: Record<string, string> = {};
    for (const field of visibleFields) {
      if (!field.required) {
        continue;
      }
      const value = effectiveDraft[field.key];
      if (value === undefined || value === null || value === "" || value === false) {
        nextErrors[field.key] = `${field.label} is required.`;
      }
    }
    setFieldErrors(nextErrors);
    return Object.keys(nextErrors).length === 0;
  }

  async function submit(status: "draft" | "submitted"): Promise<void> {
    setError(null);
    setMessage(null);
    setFieldErrors({});

    const response = await fetch(`/api/platform/tenants/${tenantSlug}/forms/${form.key}/submissions`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
      },
      body: JSON.stringify({
        data: effectiveDraft,
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
      try {
        window.localStorage.removeItem(storageKey);
      } catch {
        // Resume storage is optional.
      }
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
            <p className={styles.cardEyebrow}>Step {activeStepIndex + 1}</p>
            <h3>{currentStep.title}</h3>
            {currentStep.description ? <p className={styles.helperCopy}>{currentStep.description}</p> : null}
          </div>
          <span className={styles.badge}>
            {activeStepIndex + 1} / {resolvedSteps.length}
          </span>
        </div>

        <div className={styles.inlineList}>
          {resolvedSteps.map((step, index) => (
            <span className={styles.inlineTag} key={step.id}>
              {index + 1}. {step.title}
            </span>
          ))}
        </div>

        <div className={styles.formGrid}>
          {visibleFields.map((field) => (
            <label className={styles.formFieldSpan} key={field.id}>
              <span>{field.label}</span>
              {field.type === "long_text" ? (
                <textarea
                  className={styles.textarea}
                  onChange={(event) => setDraft((current) => ({ ...(current ?? effectiveDraft), [field.key]: event.target.value }))}
                  placeholder={field.placeholder}
                  value={String(effectiveDraft[field.key] ?? field.defaultValue ?? "")}
                />
              ) : field.type === "boolean" ? (
                <label className={styles.checkboxField}>
                  <input
                    checked={Boolean(effectiveDraft[field.key] ?? field.defaultValue ?? false)}
                    onChange={(event) => setDraft((current) => ({ ...(current ?? effectiveDraft), [field.key]: event.target.checked }))}
                    type="checkbox"
                  />
                  <span>{field.helpText ?? field.tooltip ?? "Toggle this setting."}</span>
                </label>
              ) : field.type === "select" ? (
                <select
                  className={styles.select}
                  onChange={(event) => setDraft((current) => ({ ...(current ?? effectiveDraft), [field.key]: event.target.value }))}
                  value={String(effectiveDraft[field.key] ?? field.defaultValue ?? "")}
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
                  onChange={(event) => setDraft((current) => ({ ...(current ?? effectiveDraft), [field.key]: event.target.value }))}
                  placeholder={field.placeholder}
                  type={fieldInputType(field.type)}
                  value={String(effectiveDraft[field.key] ?? field.defaultValue ?? "")}
                />
              )}
              {field.helpText || field.tooltip ? <small className={styles.metricMeta}>{field.helpText ?? field.tooltip}</small> : null}
              {fieldErrors[field.key] ? <small className={styles.errorText}>{fieldErrors[field.key]}</small> : null}
            </label>
          ))}
        </div>

        <div className={styles.actionsRow}>
          <div className={styles.inlineList}>
            <button
              className={styles.secondaryButton}
              disabled={activeStepIndex === 0 || isPending}
              onClick={() => setCurrentStepIndex(Math.max(0, activeStepIndex - 1))}
              type="button"
            >
              Previous
            </button>
            <button
              className={styles.secondaryButton}
              disabled={activeStepIndex >= resolvedSteps.length - 1 || isPending}
              onClick={() => {
                if (!validateCurrentStep()) {
                  return;
                }
                setCurrentStepIndex(Math.min(resolvedSteps.length - 1, activeStepIndex + 1));
              }}
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
                    if (!validateCurrentStep()) {
                      return;
                    }
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
