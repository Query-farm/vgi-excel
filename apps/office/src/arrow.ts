import type { Table, Vector } from "apache-arrow";
import type { CellScalar, QueryResult } from "@query-farm/vgi-excel-core";

let resultTimeZone = "UTC";

export function setResultTimeZone(timeZone: string): void {
  if (timeZone.trim()) resultTimeZone = timeZone.trim();
}

export function arrowResult(table: Table, elapsedMs?: number): QueryResult {
  const columns = table.schema.fields.map((field) => ({ name: field.name, type: String(field.type) }));
  const decimalPolicies = table.schema.fields.map((field, column) => isDecimal(field.type)
    ? decimalColumnUsesNumbers(table.getChildAt(column), field.type)
    : undefined);
  const rows: CellScalar[][] = [];
  for (let row = 0; row < table.numRows; row++) {
    rows.push(columns.map((_column, column) => arrowCell(table.getChildAt(column), row, table.schema.fields[column]?.type, decimalPolicies[column])));
  }
  return { columns, rows, rowCount: table.numRows, elapsedMs };
}

export function arrowCell(vector: Vector | null, row: number, arrowType?: unknown, decimalAsNumber?: boolean): CellScalar {
  if (!vector || !vector.isValid(row)) return null;
  if (isTimestamp(arrowType)) {
    const raw = rawTimestamp(vector, row);
    if (raw != null) return formatRawArrowTimestamp(raw, arrowType, isZonedTimestamp(arrowType) ? resultTimeZone : undefined);
  }
  return excelScalar(vector.get(row), arrowType, decimalAsNumber);
}

export function excelScalar(value: unknown, arrowType?: unknown, decimalAsNumber?: boolean): CellScalar {
  if (value == null) return null;
  if (isTimestamp(arrowType)) {
    const timestamp = timestampMilliseconds(value);
    if (timestamp != null) return formatArrowTimestamp(timestamp, arrowType, isZonedTimestamp(arrowType) ? resultTimeZone : undefined);
  }
  if (isDate(arrowType) && typeof value === "number") return new Date(value).toISOString().slice(0, 10);
  if (isTime(arrowType)) { const time = formatArrowTime(value, arrowType); if (time) return time; }
  if (isDecimal(arrowType)) {
    const text = formatArrowDecimal(value, Number((arrowType as { scale?: unknown }).scale ?? 0));
    if (decimalAsNumber ?? isExcelSafeDecimal(text)) {
      const numeric = Number(text);
      if (Number.isFinite(numeric)) return Object.is(numeric, -0) ? 0 : numeric;
    }
    return text;
  }
  if (typeof value === "string" || typeof value === "boolean") return value;
  if (typeof value === "number") return Number.isFinite(value) ? value : String(value);
  if (typeof value === "bigint") {
    const n = Number(value);
    return Number.isSafeInteger(n) ? n : value.toString();
  }
  if (value instanceof Date) return isZonedTimestamp(arrowType) ? formatZonedTimestamp(value, resultTimeZone) : formatNaiveTimestamp(value);
  if (value instanceof Uint8Array) return toBase64(value);
  if (typeof value === "object" && "toString" in value) {
    const text = String(value);
    if (text !== "[object Object]") return text;
  }
  return JSON.stringify(value);
}

function isZonedTimestamp(arrowType: unknown): boolean {
  if (!arrowType) return false;
  const timezone = (arrowType as { timezone?: unknown }).timezone;
  if (typeof timezone === "string" && timezone.length > 0) return true;
  return /^Timestamp<[^>]+,\s*[^>]+>$/i.test(String(arrowType));
}

function isTimestamp(arrowType: unknown): boolean {
  if (!arrowType) return false;
  return /^Timestamp(?:<|$)/i.test(String(arrowType)) || (arrowType as { typeId?: unknown }).typeId === 10;
}

function isDate(arrowType: unknown): boolean {
  if (!arrowType) return false;
  return /^Date(?:32|64|<|$)/i.test(String(arrowType)) || (arrowType as { typeId?: unknown }).typeId === 8;
}

function isTime(arrowType: unknown): boolean {
  if (!arrowType) return false;
  return /^Time(?:32|64|<|$)/i.test(String(arrowType)) || (arrowType as { typeId?: unknown }).typeId === 9;
}

function isDecimal(arrowType: unknown): boolean {
  if (!arrowType) return false;
  return /^Decimal/i.test(String(arrowType)) || (arrowType as { typeId?: unknown }).typeId === 7;
}

function decimalColumnUsesNumbers(vector: Vector | null, arrowType: unknown): boolean {
  if (!vector) return false;
  const scale = Number((arrowType as { scale?: unknown }).scale ?? 0);
  for (let row = 0; row < vector.length; row++) {
    if (!vector.isValid(row)) continue;
    if (!isExcelSafeDecimal(formatArrowDecimal(vector.get(row), scale))) return false;
  }
  return true;
}

/** Excel retains at most 15 significant decimal digits. */
export function isExcelSafeDecimal(value: string): boolean {
  const match = /^-?(\d+)(?:\.(\d+))?$/.exec(value.trim());
  if (!match) return false;
  const digits = `${match[1]}${match[2] ?? ""}`.replace(/^0+/, "");
  return (digits.length || 1) <= 15 && Number.isFinite(Number(value));
}

function timestampMilliseconds(value: unknown): number | null {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value.getTime();
  if (typeof value === "number" || typeof value === "bigint") {
    const milliseconds = Number(value);
    return Number.isFinite(milliseconds) ? milliseconds : null;
  }
  return null;
}

function formatArrowTimestamp(milliseconds: number, arrowType: unknown, timeZone?: string): string {
  const precision = [0, 3, 6, 9][Number((arrowType as { unit?: unknown }).unit ?? 1)] ?? 3;
  let wholeSeconds = Math.floor(milliseconds / 1000);
  const scale = 10 ** precision;
  let fraction = precision ? Math.round((milliseconds - wholeSeconds * 1000) * 10 ** (precision - 3)) : 0;
  if (fraction >= scale) { wholeSeconds++; fraction = 0; }
  const date = new Date(wholeSeconds * 1000);
  const base = timeZone ? formatZonedTimestamp(date, timeZone) : formatNaiveTimestamp(date);
  if (!precision || fraction === 0) return base;
  return `${base}.${String(fraction).padStart(precision, "0")}`;
}

function rawTimestamp(vector: Vector, row: number): bigint | null {
  let local = row;
  for (const data of vector.data) {
    if (local >= data.length) { local -= data.length; continue; }
    const value = (data.values as BigInt64Array)[data.offset + local];
    return typeof value === "bigint" ? value : null;
  }
  return null;
}

function formatRawArrowTimestamp(raw: bigint, arrowType: unknown, timeZone?: string): string {
  const unit = Number((arrowType as { unit?: unknown }).unit ?? 1);
  const unitsPerSecond = [1n, 1_000n, 1_000_000n, 1_000_000_000n][unit] ?? 1_000n;
  let wholeSeconds = raw / unitsPerSecond;
  let remainder = raw % unitsPerSecond;
  if (remainder < 0) { wholeSeconds--; remainder += unitsPerSecond; }
  const date = new Date(Number(wholeSeconds) * 1000);
  const base = timeZone ? formatZonedTimestamp(date, timeZone) : formatNaiveTimestamp(date);
  if (remainder === 0n || unit === 0) return base;
  return `${base}.${String(remainder).padStart([0, 3, 6, 9][unit] ?? 3, "0")}`;
}

function formatArrowTime(value: unknown, arrowType: unknown): string | null {
  if (typeof value !== "number" && typeof value !== "bigint") return null;
  const unit = Number((arrowType as { unit?: unknown }).unit ?? 1);
  const factor = [1_000_000_000n, 1_000_000n, 1_000n, 1n][unit];
  if (factor == null || (typeof value === "number" && !Number.isInteger(value))) return null;
  const nanoseconds = BigInt(value) * factor;
  const hour = nanoseconds / 3_600_000_000_000n;
  const minute = nanoseconds / 60_000_000_000n % 60n;
  const second = nanoseconds / 1_000_000_000n % 60n;
  const digits = [0, 3, 6, 9][unit] ?? 0;
  const fraction = digits ? String(nanoseconds % 1_000_000_000n).padStart(9, "0").slice(0, digits).replace(/0+$/, "") : "";
  return `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}:${String(second).padStart(2, "0")}${fraction ? `.${fraction}` : ""}`;
}

function formatArrowDecimal(value: unknown, scale: number): string {
  const raw = String(value);
  if (!Number.isInteger(scale) || scale <= 0 || !/^-?\d+$/.test(raw)) return raw;
  const negative = raw.startsWith("-");
  const digits = negative ? raw.slice(1) : raw;
  const padded = digits.padStart(scale + 1, "0");
  return `${negative ? "-" : ""}${padded.slice(0, -scale)}.${padded.slice(-scale)}`;
}

function formatNaiveTimestamp(value: Date): string {
  return value.toISOString().replace("T", " ").replace(/\.000Z$/, "").replace(/Z$/, "");
}

function formatZonedTimestamp(value: Date, timeZone: string): string {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone, year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit", fractionalSecondDigits: 3, hourCycle: "h23",
  }).formatToParts(value);
  const part = (name: Intl.DateTimeFormatPartTypes) => parts.find((item) => item.type === name)?.value ?? "";
  const fraction = part("fractionalSecond");
  return `${part("year")}-${part("month")}-${part("day")} ${part("hour")}:${part("minute")}:${part("second")}${fraction === "000" ? "" : `.${fraction}`}`;
}

function toBase64(bytes: Uint8Array): string {
  let binary = "";
  for (const value of bytes) binary += String.fromCharCode(value);
  return btoa(binary);
}
