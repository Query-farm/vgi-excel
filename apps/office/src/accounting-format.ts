import type { QueryResult } from "@query-farm/vgi-excel-core";

export function excelColumnNumberFormat(type: string, rows: QueryResult["rows"], column: number): string {
  const values = rows.map((row) => row[column]).filter((value) => value != null);
  const decimal = /(?:DECIMAL\s*\(\s*\d+\s*,\s*(\d+)\s*\)|Decimal\[\d+e\+?(\d+)\])/i.exec(type);
  if (decimal) {
    if (values.some((value) => typeof value === "string")) return "@";
    const scale = Math.min(30, Number(decimal[1] ?? decimal[2] ?? 0));
    return scale > 0 ? `#,##0.${"0".repeat(scale)}` : "#,##0";
  }
  if (/^(?:U?(?:TINY|SMALL|BIG)?INT(?:EGER)?|HUGEINT)|Int(?:8|16|32|64)|Uint(?:8|16|32|64)/i.test(type))
    return values.some((value) => typeof value === "string") ? "@" : "#,##0";
  if (/DOUBLE|FLOAT|REAL|NUMBER|Float(?:16|32|64)/i.test(type)) return "#,##0.########";
  return "General";
}
