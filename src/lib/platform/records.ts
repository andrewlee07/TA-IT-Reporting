import { findObjectDefinition } from "@/lib/platform/manifest";
import { evaluateRuleAsBoolean, evaluateRuleExpression } from "@/lib/platform/rule-engine";
import type {
  FieldDefinition,
  MaskingMode,
  ObjectDefinition,
  PlatformManifest,
  PlatformRecord,
} from "@/lib/platform/types";

function escapeRegex(value: string): RegExp {
  return new RegExp(value);
}

function coerceValue(field: FieldDefinition, value: unknown): unknown {
  if (value === undefined) {
    return field.defaultValue ?? null;
  }

  if (value === null) {
    return null;
  }

  switch (field.type) {
    case "number":
    case "currency":
      if (typeof value === "number") {
        return value;
      }

      if (typeof value === "string" && value.trim() !== "") {
        return Number(value);
      }

      return null;
    case "boolean":
      if (typeof value === "boolean") {
        return value;
      }

      if (typeof value === "string") {
        return value === "true" || value === "1";
      }

      return Boolean(value);
    case "date":
    case "datetime":
      return typeof value === "string" ? value : String(value);
    default:
      return typeof value === "string" ? value : String(value);
  }
}

export function normalizeRecordInput(objectDefinition: ObjectDefinition, rawData: Record<string, unknown>): Record<string, unknown> {
  const normalizedEntries = objectDefinition.fields.map((field) => [field.key, coerceValue(field, rawData[field.key])]);
  return Object.fromEntries(normalizedEntries);
}

export function validateRecordInput(input: {
  objectDefinition: ObjectDefinition;
  manifest: PlatformManifest;
  rawData: Record<string, unknown>;
  existingRecords: PlatformRecord[];
  currentRecordId?: string;
}): { data: Record<string, unknown>; errors: string[] } {
  const normalizedData = normalizeRecordInput(input.objectDefinition, input.rawData);
  const errors: string[] = [];

  for (const field of input.objectDefinition.fields) {
    const value = normalizedData[field.key];
    const hasValue = value !== null && value !== undefined && value !== "";
    const ruleMakesMandatory = field.mandatoryRule ? evaluateRuleAsBoolean(field.mandatoryRule, normalizedData) : false;

    if ((field.required || ruleMakesMandatory) && !hasValue) {
      errors.push(field.validations.find((rule) => rule.type === "required")?.message ?? `${field.label} is required.`);
      continue;
    }

    for (const rule of field.validations) {
      if (!hasValue && rule.type !== "required") {
        continue;
      }

      if (rule.type === "min" && typeof value === "number" && typeof rule.value === "number" && value < rule.value) {
        errors.push(rule.message);
      }

      if (rule.type === "max" && typeof value === "number" && typeof rule.value === "number" && value > rule.value) {
        errors.push(rule.message);
      }

      if (rule.type === "regex" && typeof value === "string" && typeof rule.value === "string" && !escapeRegex(rule.value).test(value)) {
        errors.push(rule.message);
      }

      if (rule.rule && !evaluateRuleAsBoolean(rule.rule, { ...normalizedData, value })) {
        errors.push(rule.message);
      }
    }

    if (field.unique && hasValue) {
      const duplicate = input.existingRecords.find(
        (record) => record.id !== input.currentRecordId && record.data[field.key] === value,
      );
      if (duplicate) {
        errors.push(`${field.label} must be unique.`);
      }
    }

    if (field.type === "select" && field.options?.length && hasValue && typeof value === "string" && !field.options.includes(value)) {
      errors.push(`${field.label} must be one of: ${field.options.join(", ")}.`);
    }

    if (field.advancedValidation && hasValue && !evaluateRuleAsBoolean(field.advancedValidation, { ...normalizedData, value })) {
      errors.push(`${field.label} failed the advanced validation rule.`);
    }
  }

  for (const field of input.objectDefinition.fields) {
    if (!field.calculation) {
      continue;
    }

    try {
      if (!field.calculation.rule || evaluateRuleAsBoolean(field.calculation.rule, normalizedData)) {
        normalizedData[field.key] = evaluateRuleExpression(
          {
            mode: "text",
            expression: field.calculation.expression,
          },
          normalizedData,
        );
      }
    } catch (error) {
      errors.push(error instanceof Error ? error.message : `Failed to calculate ${field.label}.`);
    }
  }

  return {
    data: normalizedData,
    errors,
  };
}

function getMaskingMode(manifest: PlatformManifest): MaskingMode {
  const policy = manifest.maskingPolicies.find((candidate) => candidate.key === manifest.securityPolicy.defaultMaskingPolicyKey);
  return policy?.mode ?? "mask";
}

function maskString(value: string, mode: MaskingMode): string {
  if (mode === "block") {
    return "[blocked]";
  }

  if (mode === "anonymize") {
    return "[anonymized]";
  }

  return value.length <= 4 ? "****" : `${"*".repeat(Math.max(4, value.length - 4))}${value.slice(-4)}`;
}

export function maskRecordForAgent(input: {
  manifest: PlatformManifest;
  objectKey: string;
  record: PlatformRecord;
}): Record<string, unknown> {
  const objectDefinition = findObjectDefinition(input.manifest, input.objectKey);
  if (!objectDefinition) {
    return input.record.data;
  }

  const mode = getMaskingMode(input.manifest);
  return Object.fromEntries(
    objectDefinition.fields.map((field) => {
      const value = input.record.data[field.key];
      if (field.sensitivity === "public" || value === null || value === undefined) {
        return [field.key, value];
      }

      if (typeof value === "string") {
        return [field.key, maskString(value, mode)];
      }

      if (mode === "block") {
        return [field.key, "[blocked]"];
      }

      return [field.key, mode === "anonymize" ? "[anonymized]" : "[masked]"];
    }),
  );
}
