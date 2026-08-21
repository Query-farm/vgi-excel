import type { QueryResult } from "@query-farm/vgi-excel-core";
import { captureError } from "./telemetry";

export interface DesktopConnection {
  name: string;
  catalog: string;
  location: string;
  authentication: "anonymous" | "oauth";
  attachOptions?: Record<string, string | number | boolean | null>;
  isDefault?: boolean;
  isSignedIn?: boolean;
}

export interface WorkbookOverview { workbook: string; activeSheet: string; selection: string; worksheets: Array<{ name: string; usedRange: string; formulaCount: number; tables: Array<{ name: string; address: string; rows: number; columns: number }> }> }
export interface WorkbookRange { sheet: string; address: string; values: unknown[][]; formulas: unknown[][] }
export interface WorkbookFormula { sheet: string; address: string; formula: string; value: unknown }
export interface WorkbookWriteRequest { mode: "new_sheet" | "replace_table"; result: QueryResult; sheetName?: string; tableName: string }
export interface ProductInformation { name: string; version: string; build: string }
export interface WorkbookWriteOutcome { sheet: string; table: string; address: string }
export interface ManagedSnapshot { table: string; connection: string; sql: string; updatedAt: string }
export interface PowerQueryOutcome { query: string; loaded: boolean; sheet?: string; table?: string; message: string }

interface HostResponse { id: number; result?: unknown; error?: string }

let nextId = 1;
const pending = new Map<number, { method: string; resolve(value: unknown): void; reject(error: Error): void }>();

function receiveHostResponse(value: unknown): void {
  let message = value as HostResponse;
  if (typeof value === "string") {
    try { message = JSON.parse(value) as HostResponse; }
    catch { return; }
  }
  if (!message || typeof message.id !== "number") return;
  const request = pending.get(message.id);
  if (!request) return;
  pending.delete(message.id);
  if (message.error) {
    const error = new Error(message.error);
    captureError(error, `bridge.${request.method}`);
    request.reject(error);
  }
  else request.resolve(message.result);
}

if (typeof window !== "undefined") {
  window.vgiReceiveHostResponse = receiveHostResponse;
  window.chrome?.webview?.addEventListener("message", (event) => receiveHostResponse(event.data));
}

export function invoke<T>(method: string, params: Record<string, unknown> = {}, timeoutMs = 120_000): Promise<T> {
  const webview = typeof window === "undefined" ? undefined : window.chrome?.webview;
  if (!webview) return Promise.reject(new Error("The VGI native bridge is unavailable."));
  const id = nextId++;
  return new Promise<T>((resolve, reject) => {
    const timeout = globalThis.setTimeout(() => {
      if (!pending.delete(id)) return;
      const error = new Error(`The Excel add-in did not respond to ${method}. Close and reopen Cupola for Excel, then try again.`);
      captureError(error, `bridge.${method}.timeout`);
      reject(error);
    }, timeoutMs);
    pending.set(id, {
      method,
      resolve: (value) => { globalThis.clearTimeout(timeout); resolve(value as T); },
      reject: (error) => { globalThis.clearTimeout(timeout); reject(error); },
    });
    webview.postMessage({ id, method, params });
  });
}

export const host = {
  appInfo: () => invoke<ProductInformation>("app.info", {}, 15_000),
  diagnostics: () => invoke<string>("app.diagnostics", {}, 15_000),
  ready: () => invoke<boolean>("ui.ready", {}, 15_000),
  loadAgentKey: () => invoke<string | null>("agent.key.load", {}, 15_000),
  saveAgentKey: (key: string) => invoke<boolean>("agent.key.save", { key }, 15_000),
  deleteAgentKey: () => invoke<boolean>("agent.key.delete", {}, 15_000),
  traceAgent: (event: Record<string, unknown>) => invoke<boolean>("agent.trace", { event }, 15_000),
  copyText: (value: string) => invoke<boolean>("clipboard.write", { value }, 15_000),
  connections: () => invoke<DesktopConnection[]>("connections.list", {}, 15_000),
  saveConnection: (connection: DesktopConnection, makeDefault = false) => invoke<DesktopConnection[]>("connections.save", { connection, makeDefault }, 15_000),
  useConnection: (name: string) => invoke<DesktopConnection[]>("connections.use", { name }, 15_000),
  removeConnection: (name: string) => invoke<DesktopConnection[]>("connections.remove", { name }, 15_000),
  testConnection: (connection: DesktopConnection) => invoke<QueryResult>("connections.test", { connection }),
  signIn: (connection: DesktopConnection) => invoke<DesktopConnection[]>("connections.signIn", { connection }),
  signOut: (connection: DesktopConnection) => invoke<DesktopConnection[]>("connections.signOut", { connection }),
  query: (sql: string, connection: string, agent = false, maxRows = 10_000) =>
    invoke<QueryResult>("query.run", { sql, connection, agent, maxRows }),
  insert: (result: QueryResult, tableName: string, source?: { connection: string; sql: string }) => invoke<WorkbookWriteOutcome>("excel.insert", { result, tableName, ...source }),
  insertQuery: (sql: string, connection: string, tableName: string) => invoke<WorkbookWriteOutcome>("excel.insertQuery", { sql, connection, tableName }, 300_000),
  createPowerQuery: (sql: string, connection: string, name?: string, loadToWorksheet = true) => invoke<PowerQueryOutcome>("excel.createPowerQuery", { sql, connection, name, loadToWorksheet }, 300_000),
  activateTable: (tableName: string) => invoke<boolean>("excel.activateTable", { tableName }, 15_000),
  snapshots: () => invoke<ManagedSnapshot[]>("excel.snapshots", {}, 15_000),
  refreshSnapshot: (tableName: string) => invoke<WorkbookWriteOutcome>("excel.refreshSnapshot", { tableName }),
  forgetSnapshot: (tableName: string) => invoke<boolean>("excel.forgetSnapshot", { tableName }, 15_000),
  workbookOverview: () => invoke<WorkbookOverview>("excel.workbookOverview", {}, 30_000),
  readRange: (sheet: string, address: string) => invoke<WorkbookRange>("excel.readRange", { sheet, address }, 30_000),
  listFormulas: (sheet?: string, limit = 200) => invoke<WorkbookFormula[]>("excel.listFormulas", { sheet, limit }, 30_000),
  writeResult: (request: WorkbookWriteRequest) => invoke<WorkbookWriteOutcome>("excel.writeResult", { ...request }, 30_000),
};
