import jsonLogic from "json-logic-js";

import type { RuleExpressionDefinition } from "@/lib/platform/types";

const RESERVED_IDENTIFIERS = new Set(["true", "false", "null", "undefined"]);
const SAFE_EXPRESSION_PATTERN = /^[\w\s()+\-*/%.'&|<>=!?:"[\],]+$/;

function evaluateTextExpression(expression: string, scope: Record<string, unknown>): unknown {
  if (!SAFE_EXPRESSION_PATTERN.test(expression)) {
    throw new Error("Expression contains unsupported characters.");
  }

  const identifiers = Array.from(expression.matchAll(/\b[a-zA-Z_][a-zA-Z0-9_]*\b/g))
    .map((match) => match[0])
    .filter((identifier) => !RESERVED_IDENTIFIERS.has(identifier));

  const uniqueIdentifiers = Array.from(new Set(identifiers));
  for (const identifier of uniqueIdentifiers) {
    if (!(identifier in scope)) {
      throw new Error(`Unknown field reference "${identifier}" in expression.`);
    }
  }

  const evaluator = new Function(...uniqueIdentifiers, `return (${expression});`);
  return evaluator(...uniqueIdentifiers.map((identifier) => scope[identifier]));
}

export function normalizeRuleExpression(
  rule: RuleExpressionDefinition | null | undefined,
): RuleExpressionDefinition | undefined {
  if (!rule) {
    return undefined;
  }

  if (rule.mode === "json_logic") {
    if (!rule.jsonLogic || Object.keys(rule.jsonLogic).length === 0) {
      return undefined;
    }

    return {
      mode: "json_logic",
      summary: rule.summary?.trim() || undefined,
      jsonLogic: rule.jsonLogic,
    };
  }

  const expression = rule.expression?.trim();
  if (!expression) {
    return undefined;
  }

  return {
    mode: "text",
    summary: rule.summary?.trim() || undefined,
    expression,
  };
}

export function evaluateRuleExpression(
  rule: RuleExpressionDefinition | null | undefined,
  scope: Record<string, unknown>,
): unknown {
  const normalized = normalizeRuleExpression(rule);
  if (!normalized) {
    return undefined;
  }

  if (normalized.mode === "json_logic") {
    return jsonLogic.apply(normalized.jsonLogic ?? {}, scope);
  }

  return evaluateTextExpression(normalized.expression ?? "", scope);
}

export function evaluateRuleAsBoolean(
  rule: RuleExpressionDefinition | null | undefined,
  scope: Record<string, unknown>,
): boolean {
  const value = evaluateRuleExpression(rule, scope);
  return Boolean(value);
}
