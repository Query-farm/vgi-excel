import { afterEach, expect, it, vi } from "vitest";
vi.mock("./confirmation", () => ({ confirmAction: vi.fn() }));
import { confirmAction } from "./confirmation";
vi.mock("./runtime", () => ({ resolveBackend: vi.fn() }));
import { insertResult } from "./workbook";

afterEach(() => vi.unstubAllGlobals());
it("inserts a static table without adding or removing workbook refresh metadata", async () => {
  const addSetting = vi.fn(), removeSetting = vi.fn();
  const range = { load: vi.fn(), address: "Sheet1!A1:A2", format: { autofitColumns: vi.fn() }, values: [] as unknown[][], numberFormat: [] as string[][] };
  const table = { name: "" };
  const addTable = vi.fn(() => table);
  const sheet = { name: "Sheet1", load: vi.fn(), getUsedRangeOrNullObject: () => ({ isNullObject: true, load: vi.fn() }), getRangeByIndexes: () => range };
  const context = { sync: vi.fn(async () => {}), workbook: {
    getActiveCell: () => ({ rowIndex: 0, columnIndex: 0, load: vi.fn() }),
    worksheets: { getActiveWorksheet: () => sheet },
    tables: { load: vi.fn(), items: [], add: addTable },
    settings: { add: addSetting, getItem: () => ({ delete: removeSetting }) },
  } };
  vi.stubGlobal("Excel", { run: async (action: (value: unknown) => unknown) => action(context) });
  const outcome = await insertResult({ columns: [{ name: "value", type: "INTEGER" }], rows: [[42]], rowCount: 1 });
  expect(outcome?.table).toBe("VGI_Result");
  expect(range.values).toEqual([[42]]);
  expect(addTable).toHaveBeenCalledOnce();
  expect(addSetting).not.toHaveBeenCalled();
  expect(removeSetting).not.toHaveBeenCalled();
});


it.each([false, true])("waits for overwrite approval before writing occupied cells (%s)", async accepted => {
  let decide!: (accepted: boolean) => void;
  vi.mocked(confirmAction).mockImplementationOnce(() => new Promise(resolve => { decide = resolve; }));
  const range = { load: vi.fn(), address: "Sheet1!A1:A2", format: { autofitColumns: vi.fn() }, values: [] as unknown[][], numberFormat: [] as string[][] };
  const addTable = vi.fn(() => ({ name: "" }));
  const sheet = { name: "Sheet1", load: vi.fn(), getUsedRangeOrNullObject: () => ({ isNullObject: false, rowIndex: 0, columnIndex: 0, rowCount: 2, columnCount: 1, load: vi.fn() }), getRangeByIndexes: () => range };
  const context = { sync: vi.fn(async () => {}), workbook: {
    getActiveCell: () => ({ rowIndex: 0, columnIndex: 0, load: vi.fn() }),
    worksheets: { getActiveWorksheet: () => sheet },
    tables: { load: vi.fn(), items: [], add: addTable },
  } };
  vi.stubGlobal("Excel", { run: async (action: (value: unknown) => unknown) => action(context) });
  const pending = insertResult({ columns: [{ name: "value", type: "INTEGER" }], rows: [[42]], rowCount: 1 });
  await vi.waitFor(() => expect(decide).toBeTypeOf("function"));
  expect(addTable).not.toHaveBeenCalled();
  expect(range.values).toEqual([]);
  decide(accepted);
  const outcome = await pending;
  expect(addTable).toHaveBeenCalledTimes(accepted ? 1 : 0);
  expect(range.values).toEqual(accepted ? [[42]] : []);
  if (!accepted) expect(outcome).toBeNull();
});
