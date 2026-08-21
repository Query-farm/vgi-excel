const REDACTED = "[redacted]";
const QUERY_REDACTED = "[query details redacted]";

const sensitiveKeys = /(authorization|auth|bearer|catalog|connection|credential|dsn|endpoint|formula|key|location|password|prompt|query|refresh.?token|request|response|result|rows?|secret|sheet|sql|table|token|url|values?|workbook)/i;
const sqlText = /\b(?:select\s+.+\s+from|with\s+.+\s+as\s*\(|insert\s+into|update\s+.+\s+set|delete\s+from|attach\s+.+\s+as|copy\s+.+\s+(?:to|from)|create\s+(?:or\s+replace\s+)?table|alter\s+table|drop\s+table|pragma\s+|call\s+[\w".]+\s*\()/is;

/**
 * Removes customer-controlled data from text before it reaches an error service.
 * This deliberately favors privacy over preserving a detailed remote error message.
 */
export function sanitizeTelemetryText(value: unknown): string {
  if (value === null || value === undefined) return "";
  let text = String(value).slice(0, 4_000);
  if (sqlText.test(text)) return QUERY_REDACTED;

  text = text
    .replace(/\b(?:Bearer|Basic)\s+[A-Za-z0-9._~+/=-]+/gi, "[credential redacted]")
    .replace(/\b(?:sk-ant-|sk-|pk_)[A-Za-z0-9_-]{12,}\b/g, "[credential redacted]")
    .replace(/\b(?:access[_ -]?token|refresh[_ -]?token|api[_ -]?key|password|secret)\s*[:=]\s*[^\s,;]+/gi, "$1=[credential redacted]")
    .replace(/https?:\/\/[^\s<>()\[\]{}]+/gi, "[endpoint redacted]")
    .replace(/\b[A-Za-z]:\\(?:[^\\\r\n]+\\)*[^\\\r\n]*/g, "[local path redacted]")
    .replace(/\/(?:Users|home)\/[^\s:]+/g, "[local path redacted]")
    .replace(/["'“”‘’][^"'“”‘’\r\n]{1,160}["'“”‘’]/g, "[identifier redacted]");

  return text || REDACTED;
}

/** Returns a bounded, telemetry-safe clone of arbitrary structured data. */
export function sanitizeTelemetryValue(value: unknown, depth = 0): unknown {
  if (depth > 5) return "[depth redacted]";
  if (value === null || typeof value === "boolean" || typeof value === "number") return value;
  if (typeof value === "string" || typeof value === "bigint" || typeof value === "symbol") {
    return sanitizeTelemetryText(value);
  }
  if (typeof value === "function" || typeof value === "undefined") return undefined;
  if (value instanceof Error) {
    return {
      name: sanitizeTelemetryText(value.name),
      message: sanitizeTelemetryText(value.message),
    };
  }
  if (Array.isArray(value)) return value.slice(0, 20).map((item) => sanitizeTelemetryValue(item, depth + 1));

  const source = value as Record<string, unknown>;
  const result: Record<string, unknown> = {};
  for (const [key, item] of Object.entries(source).slice(0, 40)) {
    result[key] = sensitiveKeys.test(key) ? REDACTED : sanitizeTelemetryValue(item, depth + 1);
  }
  return result;
}

export function isTelemetrySensitiveKey(key: string): boolean {
  return sensitiveKeys.test(key);
}

/** Maps arbitrary customer-controlled failure text to a small, fixed vocabulary. */
export function classifyTelemetryError(value: unknown): string {
  const text = value instanceof Error ? `${value.name} ${value.message}` : String(value ?? "");
  if (/abort|cancel|stopp?ed/i.test(text)) return "Operation canceled";
  if (/timed?\s*out|timeout|did not respond/i.test(text)) return "Operation timed out";
  if (/\b(?:401|403)\b|oauth|authenticat|unauthori|forbidden|token expired|invalid_grant/i.test(text)) return "Authentication failure";
  if (/\b(?:429|529)\b|rate.?limit|overload|anthropic/i.test(text)) return "AI provider failure";
  if (/parser error|syntax error/i.test(text)) return "SQL parser error";
  if (/binder error|binding error/i.test(text)) return "SQL binder error";
  if (/catalog|schema|function inventory/i.test(text)) return "Catalog operation failure";
  if (/worksheet|workbook|excel|range|cell|table name/i.test(text)) return "Excel operation failure";
  if (/webview|native bridge|xll|excel-dna/i.test(text)) return "Add-in bridge failure";
  if (/network|fetch|dns|socket|connect|http\s*[45]\d\d/i.test(text)) return "Network failure";
  return "Unexpected application error";
}
