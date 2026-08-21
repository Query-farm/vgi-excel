import type { QueryResult } from "@query-farm/vgi-excel-core";

export const EXCEL_WORKSHEET_ROWS = 1_048_576;
export const EXCEL_WORKSHEET_COLUMNS = 16_384;
export const EXCEL_MAX_DATA_ROWS = EXCEL_WORKSHEET_ROWS - 1;

export function assertCompleteExcelResult(result: QueryResult): void {
  if (!result.columns.length) throw new Error("The query returned no columns.");
  if (result.truncated || result.rows.length < result.rowCount) throw new Error("Only a query preview is loaded. Load the complete result before writing it to Excel.");
  if (result.rows.length > EXCEL_MAX_DATA_ROWS) throw new Error(`The query returned ${result.rows.length.toLocaleString()} rows. Excel tables can contain at most ${EXCEL_MAX_DATA_ROWS.toLocaleString()} data rows on a worksheet.`);
}

export function validateExcelRange(startRow: number, startColumn: number, rows: number, columns: number): void {
  if (startRow < 0 || startColumn < 0 || rows < 1 || columns < 1) throw new Error("The Excel output range is invalid.");
  if (startRow + rows > EXCEL_WORKSHEET_ROWS) throw new Error(`This result does not fit below the selected cell. Excel supports ${EXCEL_WORKSHEET_ROWS.toLocaleString()} worksheet rows, including the table header.`);
  if (startColumn + columns > EXCEL_WORKSHEET_COLUMNS) throw new Error(`This result does not fit to the right of the selected cell. Excel supports ${EXCEL_WORKSHEET_COLUMNS.toLocaleString()} worksheet columns.`);
}
