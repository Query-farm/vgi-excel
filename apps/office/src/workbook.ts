import type { CatalogFunction, QueryResult } from "@query-farm/vgi-excel-core";
import { wrapperFormula, wrapperName } from "@query-farm/vgi-excel-core";
import { sanitizeTableName } from "./browser-backend";
import { assertCompleteExcelResult, validateExcelRange } from "./excel-limits";
import { resolveBackend } from "./runtime";
import { excelColumnNumberFormat } from "./accounting-format";

interface ImportBackend {
  importTable(name: string, headers: string[], values: unknown[][]): Promise<void>;
}

export interface WorkbookWriteOutcome { sheet: string; table: string; address: string; rows: number }
export interface SnapshotSource { connection: string; sql: string }
export interface ManagedSnapshot extends SnapshotSource { table: string; updatedAt: string }
const SNAPSHOT_PREFIX = "cupola.snapshot.";
const WRITE_CHUNK_ROWS = 10_000;

export async function listWorkbookTables(): Promise<Array<{ name: string; worksheet: string; rows: number }>> {
  return Excel.run(async (context) => {
    const tables = context.workbook.tables;
    tables.load("items/name,items/worksheet/name,items/rows/count");
    await context.sync();
    return tables.items.map((table) => ({
      name: table.name,
      worksheet: table.worksheet.name,
      rows: table.rows.count,
    }));
  });
}

export async function importWorkbookTable(name: string): Promise<string> {
  const snapshot = await Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(name);
    const header = table.getHeaderRowRange();
    const body = table.getDataBodyRange();
    header.load("values");
    body.load("values");
    await context.sync();
    return { headers: header.values[0].map(String), values: body.values as unknown[][] };
  });
  await importSnapshot(name, snapshot.headers, snapshot.values);
  return `excel.${sanitizeTableName(name)}`;
}

export async function importSelection(): Promise<string> {
  const snapshot = await Excel.run(async (context) => {
    const selection = context.workbook.getSelectedRange();
    selection.load("values,rowCount,columnCount");
    await context.sync();
    if (selection.rowCount < 1 || selection.columnCount < 1) throw new Error("Select a non-empty range first.");
    const first = selection.values[0] as unknown[];
    const hasHeaders = first.every((value) => typeof value === "string" && value.trim().length > 0);
    const headers = hasHeaders ? first.map(String) : first.map((_value, index) => `column_${index + 1}`);
    return { headers, values: (hasHeaders ? selection.values.slice(1) : selection.values) as unknown[][] };
  });
  await importSnapshot("selection", snapshot.headers, snapshot.values);
  return "excel.selection";
}

export async function insertResult(result: QueryResult, tableName = "VGI_Result", source?: SnapshotSource): Promise<WorkbookWriteOutcome | null> {
  assertCompleteExcelResult(result);
  return Excel.run(async (context) => {
    const active = context.workbook.getActiveCell();
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRangeOrNullObject(true);
    sheet.load("name"); active.load("rowIndex,columnIndex"); usedRange.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
    await context.sync();
    validateExcelRange(active.rowIndex, active.columnIndex, result.rows.length + 1, result.columns.length);
    const target = sheet.getRangeByIndexes(active.rowIndex, active.columnIndex, result.rows.length + 1, result.columns.length);
    target.load("address"); await context.sync();
    const occupied = !usedRange.isNullObject && rangesOverlap(active.rowIndex, active.columnIndex, result.rows.length + 1, result.columns.length, usedRange.rowIndex, usedRange.columnIndex, usedRange.rowCount, usedRange.columnCount);
    if (occupied && !window.confirm(`The output range ${target.address} overlaps existing worksheet data. Continue?`)) return null;
    await writeResultRows(context, sheet, active.rowIndex, active.columnIndex, result);
    const existing = context.workbook.tables;
    existing.load("items/name");
    await context.sync();
    const usedNames = new Set(existing.items.map((table) => table.name.toUpperCase()));
    let name = sanitizeTableName(tableName);
    let suffix = 2;
    while (usedNames.has(name.toUpperCase())) name = `${sanitizeTableName(tableName)}_${suffix++}`;
    const table = context.workbook.tables.add(target, true);
    table.name = name;
    target.format.autofitColumns();
    if (source) context.workbook.settings.add(SNAPSHOT_PREFIX + name, JSON.stringify({ table: name, ...source, updatedAt: new Date().toISOString() }));
    await context.sync();
    return { sheet: sheet.name, table: name, address: target.address, rows: result.rows.length };
  });
}

export async function goToTable(name: string): Promise<void> {
  await Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(name);
    table.getRange().select();
    await context.sync();
  });
}

export async function listManagedSnapshots(): Promise<ManagedSnapshot[]> {
  return Excel.run(async (context) => {
    const settings = context.workbook.settings;
    settings.load("items/key,value");
    await context.sync();
    return settings.items.filter((item) => item.key.startsWith(SNAPSHOT_PREFIX)).flatMap((item) => {
      try { return [JSON.parse(String(item.value)) as ManagedSnapshot]; } catch { return []; }
    });
  });
}

export async function refreshSnapshot(snapshot: ManagedSnapshot): Promise<WorkbookWriteOutcome> {
  const result = await (await resolveBackend(snapshot.connection)).query(snapshot.sql);
  assertCompleteExcelResult(result);
  return Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(snapshot.table);
    const oldRange = table.getRange();
    const sheet = table.worksheet;
    oldRange.load("address,rowIndex,columnIndex"); sheet.load("name");
    await context.sync();
    validateExcelRange(oldRange.rowIndex, oldRange.columnIndex, result.rows.length + 1, result.columns.length);
    const target = sheet.getRangeByIndexes(oldRange.rowIndex, oldRange.columnIndex, result.rows.length + 1, result.columns.length);
    oldRange.clear("Contents"); await context.sync();
    await writeResultRows(context, sheet, oldRange.rowIndex, oldRange.columnIndex, result);
    table.resize(target); target.format.autofitColumns();
    context.workbook.settings.add(SNAPSHOT_PREFIX + snapshot.table, JSON.stringify({ ...snapshot, updatedAt: new Date().toISOString() }));
    target.load("address");
    await context.sync();
    return { sheet: sheet.name, table: snapshot.table, address: target.address, rows: result.rows.length };
  });
}

async function writeResultRows(context: Excel.RequestContext, sheet: Excel.Worksheet, startRow: number, startColumn: number, result: QueryResult): Promise<void> {
  sheet.getRangeByIndexes(startRow, startColumn, 1, result.columns.length).values = [result.columns.map((column) => column.name)];
  await context.sync();
  const numberFormats = result.columns.map((column, columnIndex) => excelColumnNumberFormat(column.type, result.rows, columnIndex));
  for (let offset = 0; offset < result.rows.length; offset += WRITE_CHUNK_ROWS) {
    const values = result.rows.slice(offset, offset + WRITE_CHUNK_ROWS).map((row) => row.map((cell) => cell ?? ""));
    const range = sheet.getRangeByIndexes(startRow + 1 + offset, startColumn, values.length, result.columns.length);
    range.values = values;
    range.numberFormat = values.map(() => numberFormats);
    await context.sync();
  }
}

function rangesOverlap(rowA: number, columnA: number, rowsA: number, columnsA: number, rowB: number, columnB: number, rowsB: number, columnsB: number): boolean {
  return rowA < rowB + rowsB && rowB < rowA + rowsA && columnA < columnB + columnsB && columnB < columnA + columnsA;
}

export async function forgetSnapshot(name: string): Promise<void> {
  await Excel.run(async (context) => { context.workbook.settings.getItem(SNAPSHOT_PREFIX + name).delete(); await context.sync(); });
}

export async function createFunctionWrappers(functions: CatalogFunction[]): Promise<string[]> {
  return Excel.run(async (context) => {
    const names = context.workbook.names;
    names.load("items/name");
    await context.sync();
    const occupied = new Set(names.items.map((item) => item.name.toUpperCase()));
    const created: string[] = [];
    for (const fn of functions.filter((item) => item.kind === "scalar")) {
      const name = wrapperName(fn, occupied);
      names.add(name, wrapperFormula(fn));
      occupied.add(name);
      created.push(name);
    }
    await context.sync();
    return created;
  });
}

async function importSnapshot(name: string, headers: string[], values: unknown[][]): Promise<void> {
  const backend = await resolveBackend();
  if (!("importTable" in backend) || typeof (backend as unknown as ImportBackend).importTable !== "function") {
    throw new Error("The active query backend cannot import workbook tables.");
  }
  await (backend as unknown as ImportBackend).importTable(name, headers, values);
}
