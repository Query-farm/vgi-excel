import * as duckdb from "@haybarn/haybarn-wasm";
import { installVgiOAuthBridge } from "@haybarn/haybarn-wasm/vgi";
import { tableFromArrays, type Table } from "apache-arrow";
import {
  assertHttpsConnection,
  qualifiedFunctionName,
  quoteIdentifier,
  quoteLiteral,
  type CellMatrix,
  type ConnectionDefinition,
  type QueryBackend,
  type QueryOptions,
  type QueryResult,
} from "@query-farm/vgi-excel-core";
import { arrowCell, arrowResult, setResultTimeZone } from "./arrow";
import { getServiceToken } from "./config";
import { sessionTokenKey } from "./config";
import { isRecoverableAuthError } from "./auth-errors";
import { signIn } from "./oauth";

type AsyncConnection = Awaited<ReturnType<duckdb.AsyncDuckDB["connect"]>>;

const DEFAULT_ARTIFACT_BASE = typeof document === "undefined"
  ? "/haybarn/"
  : new URL("./haybarn/", document.baseURI).toString();

export interface BrowserRuntimeDiagnostics {
  assetBase: string;
  crossOriginIsolated: boolean;
  sharedArrayBuffer: boolean;
  selectedBundle?: "mvp" | "eh" | "coi";
}

const runtimeDiagnostics: BrowserRuntimeDiagnostics = {
  assetBase: DEFAULT_ARTIFACT_BASE,
  crossOriginIsolated: globalThis.crossOriginIsolated === true,
  sharedArrayBuffer: typeof SharedArrayBuffer !== "undefined",
};

export function browserRuntimeDiagnostics(): BrowserRuntimeDiagnostics { return { ...runtimeDiagnostics }; }

export class BrowserBackend implements QueryBackend {
  private static boot: Promise<{ db: duckdb.AsyncDuckDB }> | null = null;
  private runtime: Promise<AsyncConnection> | null = null;
  private attached = false;

  constructor(private readonly definition: ConnectionDefinition) {}

  async query(sql: string, options: QueryOptions = {}): Promise<QueryResult> {
    const connection = await this.connection();
    const started = performance.now();
    const cancel = () => void connection.cancelSent();
    options.signal?.addEventListener("abort", cancel, { once: true });
    try {
      const table = await connection.query(sql);
      const result = arrowResult(table, performance.now() - started);
      if (options.maxRows && result.rows.length > options.maxRows) {
        result.rows = result.rows.slice(0, options.maxRows);
        result.truncated = true;
      }
      return result;
    } finally {
      options.signal?.removeEventListener("abort", cancel);
    }
  }

  async call(functionName: string, args: CellMatrix[], options: QueryOptions = {}): Promise<QueryResult> {
    const connection = await this.connection();
    const rows = args[0]?.length ?? 1;
    const columns = args[0]?.[0]?.length ?? 1;
    const data: Record<string, unknown[]> = { _row: [], _column: [] };
    args.forEach((_arg, index) => (data[`arg_${index}`] = []));
    for (let row = 0; row < rows; row++) {
      for (let column = 0; column < columns; column++) {
        data._row.push(row);
        data._column.push(column);
        args.forEach((arg, index) => data[`arg_${index}`].push(arg[row]?.[column] ?? null));
      }
    }
    const tableName = `_vgi_excel_call_${crypto.randomUUID().replaceAll("-", "")}`;
    await connection.insertArrowTable(tableFromArrays(data), { name: tableName, create: true });
    try {
      const functionSql = qualifiedFunctionName(functionName);
      const params = args.map((_arg, index) => quoteIdentifier(`arg_${index}`)).join(", ");
      const table = await connection.query(
        `SELECT _row, _column, ${functionSql}(${params}) AS value FROM ${quoteIdentifier(tableName)} ORDER BY _row, _column`,
      );
      const values = Array.from({ length: rows }, () => Array(columns).fill(null)) as CellMatrix;
      for (let index = 0; index < table.numRows; index++) {
        const row = Number(table.getChildAt(0)?.get(index));
        const column = Number(table.getChildAt(1)?.get(index));
        values[row][column] = arrowResultCell(table, 2, index);
      }
      return {
        columns: Array.from({ length: columns }, (_value, index) => ({ name: `value_${index + 1}`, type: "ANY" })),
        rows: values,
        rowCount: rows,
      };
    } finally {
      await connection.query(`DROP TABLE IF EXISTS ${quoteIdentifier(tableName)}`);
      if (options.signal?.aborted) await connection.cancelSent();
    }
  }

  async importTable(name: string, headers: string[], values: unknown[][]): Promise<void> {
    const connection = await this.connection();
    await connection.query("CREATE SCHEMA IF NOT EXISTS excel");
    const safeName = sanitizeTableName(name);
    await connection.query(`DROP TABLE IF EXISTS excel.${quoteIdentifier(safeName)}`);
    const columns: Record<string, unknown[]> = {};
    headers.forEach((header, index) => {
      const safeHeader = uniqueHeader(header, index, columns);
      columns[safeHeader] = values.map((row) => row[index] ?? null);
    });
    await connection.insertArrowTable(tableFromArrays(columns), { schema: "excel", name: safeName, create: true });
  }

  async catalogRows(sql: string): Promise<QueryResult> {
    return this.query(sql);
  }

  private async connection(): Promise<AsyncConnection> {
    assertHttpsConnection(this.definition);
    const connection = await (this.runtime ??= BrowserBackend.ensureBooted().then(({ db }) => db.connect()));
    if (!this.attached) {
      await configureTimeZone(connection);
      await connection.query("INSTALL vgi FROM community").catch(() => undefined);
      await connection.query("LOAD vgi");
      try {
        await this.attach(connection);
      } catch (error) {
        if (!isRecoverableAuthError(error)) throw error;
        sessionStorage.removeItem(sessionTokenKey(this.definition.location));
        await signIn(this.definition.location);
        await this.attach(connection);
      }
      this.attached = true;
    }
    return connection;
  }

  private async attach(connection: AsyncConnection): Promise<void> {
    const alias = this.definition.catalog ?? this.definition.name;
    const options: string[] = ["TYPE vgi", `LOCATION ${quoteLiteral(this.definition.location)}`];
    const token = getServiceToken(this.definition.location);
    if (token?.refresh_token) options.push(`oauth_refresh_token ${quoteLiteral(token.refresh_token)}`);
    else if (token?.access_token) options.push(`bearer_token ${quoteLiteral(token.access_token)}`);
    for (const [key, value] of Object.entries(this.definition.attachOptions ?? {})) {
      if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) throw new Error(`Invalid ATTACH option: ${key}`);
      options.push(`${key} ${sqlValue(value)}`);
    }
    await connection.query(`ATTACH OR REPLACE ${quoteLiteral(alias)} AS ${quoteIdentifier(alias)} (${options.join(", ")})`);
  }

  private static ensureBooted(): Promise<{ db: duckdb.AsyncDuckDB }> {
    if (!this.boot) this.boot = bootHaybarn();
    return this.boot;
  }
}

async function bootHaybarn(): Promise<{ db: duckdb.AsyncDuckDB }> {
  const base = (import.meta.env.VITE_HAYBARN_ASSET_BASE as string | undefined) ?? DEFAULT_ARTIFACT_BASE;
  runtimeDiagnostics.assetBase = base;
  const bundles: duckdb.DuckDBBundles = {
    mvp: { mainModule: `${base}duckdb-mvp.wasm`, mainWorker: `${base}duckdb-browser-mvp.worker.js` },
    eh: { mainModule: `${base}duckdb-eh.wasm`, mainWorker: `${base}duckdb-browser-eh.worker.js` },
    coi: {
      mainModule: `${base}duckdb-coi.wasm`,
      mainWorker: `${base}duckdb-browser-coi.worker.js`,
      pthreadWorker: `${base}duckdb-browser-coi.pthread.worker.js`,
    },
  };
  const selected = await duckdb.selectBundle(bundles);
  runtimeDiagnostics.selectedBundle = selected === bundles.coi ? "coi" : selected === bundles.eh ? "eh" : "mvp";
  const worker = new Worker(selected.mainWorker!);
  if (runtimeDiagnostics.crossOriginIsolated && runtimeDiagnostics.sharedArrayBuffer) installVgiOAuthBridge(worker);
  const db = new duckdb.AsyncDuckDB(new duckdb.ConsoleLogger(), worker);
  await db.instantiate(selected.mainModule, selected.pthreadWorker);
  await db.open({ arrowLosslessConversion: true });
  return { db };
}

function arrowResultCell(table: Table, column: number, row: number): string | number | boolean | null {
  return arrowCell(table.getChildAt(column), row, table.schema.fields[column]?.type);
}

interface TimeZoneConnection { query(sql: string): Promise<unknown> }

export function browserTimeZone(): string | undefined {
  try { return Intl.DateTimeFormat().resolvedOptions().timeZone || undefined; }
  catch { return undefined; }
}

export async function configureTimeZone(connection: TimeZoneConnection, timeZone = browserTimeZone()): Promise<string | undefined> {
  if (!timeZone) return undefined;
  await connection.query("INSTALL icu").catch(() => undefined);
  await connection.query("LOAD icu").catch(() => undefined);
  await connection.query(`SET TimeZone=${quoteLiteral(timeZone)}`);
  setResultTimeZone(timeZone);
  return timeZone;
}

function sqlValue(value: unknown): string {
  if (value == null) return "NULL";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") return String(value);
  return quoteLiteral(String(value));
}

export function sanitizeTableName(name: string): string {
  const cleaned = name.trim().replace(/[^A-Za-z0-9_]/g, "_").replace(/_+/g, "_");
  return cleaned || "selection";
}

function uniqueHeader(header: string, index: number, columns: Record<string, unknown[]>): string {
  const base = sanitizeTableName(header || `column_${index + 1}`);
  let name = base;
  let suffix = 2;
  while (name in columns) name = `${base}_${suffix++}`;
  return name;
}
