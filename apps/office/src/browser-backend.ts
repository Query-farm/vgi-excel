import * as duckdb from "@haybarn/haybarn-wasm";
import { installVgiOAuthBridge } from "@haybarn/haybarn-wasm/vgi";
import { tableFromArrays, tableFromIPC, type Table } from "apache-arrow";
import {
  assertHttpsConnection,
  connectionSignIn,
  ConnectionSignInError,
  catalogNames,
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
import { runConnectionTest } from "./connection-test";

type AsyncConnection = Awaited<ReturnType<duckdb.AsyncDuckDB["connect"]>>;
class ConnectionSignInRequired extends Error {}

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

/** Load the exact file deployed with this app, independently of the extension cache. */
export function bundledVgiLoadSql(bundle: "mvp" | "eh" | "coi" = runtimeDiagnostics.selectedBundle ?? "mvp", base = typeof document === "undefined" ? "https://cupola.invalid/" : new URL(import.meta.env.BASE_URL, document.baseURI).href): string {
  return `LOAD ${quoteLiteral(new URL(__VGI_EXTENSION_PATHS__[bundle], base).href)}`;
}

export class BrowserBackend implements QueryBackend {
  private static boot: Promise<{ db: duckdb.AsyncDuckDB }> | null = null;
  private runtime: Promise<AsyncConnection> | null = null;
  private attached = false;

  constructor(
    private readonly definition: ConnectionDefinition,
    private readonly runtimeFactory = () => BrowserBackend.ensureBooted(),
    private readonly interactiveAuth = true,
  ) {}

  static async discoverCatalogs(location: string, progress?: (status: string) => void): Promise<string[]> {
    assertHttpsConnection({ name: "discovery", location });
    const probe = () => runConnectionTest(async signal => {
      const { db } = await bootHaybarn(signal);
      const connection = await db.connect();
      signal?.throwIfAborted();
      await connection.query("LOAD json");
      await connection.query(bundledVgiLoadSql());
      const token = getServiceToken(location);
      await connection.query(`SET vgi_oauth_enabled=${token?.refresh_token ? "true" : "false"}`);
      const options = ["oauth_cache := 'none'"];
      if (token?.refresh_token) options.push(`oauth_refresh_token := ${quoteLiteral(token.refresh_token)}`);
      else if (token?.access_token) options.push(`bearer_token := ${quoteLiteral(token.access_token)}`);
      const result = arrowResult(await connection.query(`SELECT catalog FROM vgi_catalogs(${quoteLiteral(location)}, ${options.join(", ")})`));
      return catalogNames(result.rows);
    });
    try { return await probe(); }
    catch (error) {
      if (!isRecoverableAuthError(error)) throw error;
      sessionStorage.removeItem(sessionTokenKey(location));
      progress?.("Waiting for sign-in…");
      try { await signIn(location); }
      catch { throw new Error("Sign-in wasn’t completed. Try Find catalogs again, or enter the catalog name manually."); }
      progress?.("Finding catalogs…");
      return probe();
    }
  }

  static async testConnection(definition: ConnectionDefinition): Promise<QueryResult> {
    assertHttpsConnection(definition);
    const probe = () => runConnectionTest(signal => new BrowserBackend(definition, () => bootHaybarn(signal), false)
      .query("SELECT current_catalog(), current_schema();"));
    try { return await probe(); }
    catch (error) {
      if (!(error instanceof ConnectionSignInRequired)) throw error;
      // Human sign-in is separate from the bounded network probe. The failed
      // probe's worker has already been disposed before opening the dialog.
      sessionStorage.removeItem(sessionTokenKey(definition.location));
      await signIn(definition.location);
      return probe();
    }
  }

  private queryTail: Promise<void> = Promise.resolve();

  query(sql: string, options: QueryOptions = {}): Promise<QueryResult> {
    // One pending query per connection: a queued request must never interrupt its predecessor.
    const pending = this.queryTail.then(async () => {
      options.signal?.throwIfAborted();
      if (connectionSignIn.snapshot().includes(this.definition.name)) throw new ConnectionSignInError(this.definition.name);
      const connection = await this.connection(options.signal);
      const started = performance.now();
      const table = await this.queryConnection(connection, sql, options.signal);
      const result = arrowResult(table, performance.now() - started);
      if (options.maxRows && result.rows.length > options.maxRows) {
        result.rows = result.rows.slice(0, options.maxRows); result.truncated = true;
      }
      return result;
    });
    const reported = pending.catch(error => { throw error instanceof ConnectionSignInRequired ? error : connectionSignIn.error(this.definition.name, error); });
    this.queryTail = pending.then(() => {}, () => {});
    return reported;
  }

  private async queryConnection(connection: AsyncConnection, sql: string, signal?: AbortSignal): Promise<Table> {
    signal?.throwIfAborted();
    let cancellation: Promise<boolean> | undefined;
    const cancel = () => { cancellation = connection.cancelSent().catch(() => false); };
    signal?.addEventListener("abort", cancel, { once: true });
    try {
      return signal ? await tableFromIPC(await connection.send(sql, false)) : await connection.query(sql);
    } catch (error) {
      if (signal?.aborted) throw new DOMException("Query cancelled.", "AbortError");
      throw error;
    } finally { signal?.removeEventListener("abort", cancel); await cancellation; }
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

  private async connection(signal?: AbortSignal): Promise<AsyncConnection> {
    assertHttpsConnection(this.definition);
    const connection = await waitForConnection(this.runtime ??= this.runtimeFactory().then(({ db }) => db.connect()), signal);
    if (!this.attached) {
      await configureTimeZone(connection);
      // Load before metadata queries; threaded WASM can stall when JSON autoloads mid-query.
      signal?.throwIfAborted();
      await this.queryConnection(connection, "LOAD json", signal);
      await this.queryConnection(connection, bundledVgiLoadSql(), signal);
      try {
        await this.attach(connection, signal);
      } catch (error) {
        if (!isRecoverableAuthError(error)) throw error;
        if (!this.interactiveAuth) throw new ConnectionSignInRequired("Sign-in is required.");
        sessionStorage.removeItem(sessionTokenKey(this.definition.location));
        await signIn(this.definition.location, signal);
        await this.attach(connection, signal);
      }
      this.attached = true;
    }
    return connection;
  }

  private async attach(connection: AsyncConnection, signal?: AbortSignal): Promise<void> {
    const alias = this.definition.catalog ?? this.definition.name;
    const options: string[] = ["TYPE vgi", `LOCATION ${quoteLiteral(this.definition.location)}`];
    const token = getServiceToken(this.definition.location);
    if (!this.interactiveAuth || !runtimeDiagnostics.crossOriginIsolated || !runtimeDiagnostics.sharedArrayBuffer) await this.queryConnection(connection, `SET vgi_oauth_enabled=${token?.refresh_token ? "true" : "false"}`);
    if (token?.refresh_token) options.push(`oauth_refresh_token ${quoteLiteral(token.refresh_token)}`);
    else if (token?.access_token) options.push(`bearer_token ${quoteLiteral(token.access_token)}`);
    for (const [key, value] of Object.entries(this.definition.attachOptions ?? {})) {
      if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) throw new Error(`Invalid ATTACH option: ${key}`);
      options.push(`${key} ${sqlValue(value)}`);
    }
    await this.queryConnection(connection, `ATTACH OR REPLACE ${quoteLiteral(alias)} AS ${quoteIdentifier(alias)} (${options.join(", ")})`, signal);
  }

  private static ensureBooted(): Promise<{ db: duckdb.AsyncDuckDB }> {
    if (!this.boot) this.boot = bootHaybarn();
    return this.boot;
  }
}

async function bootHaybarn(signal?: AbortSignal): Promise<{ db: duckdb.AsyncDuckDB }> {
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
  signal?.throwIfAborted();
  runtimeDiagnostics.selectedBundle = selected.mainModule === bundles.coi?.mainModule ? "coi" : selected.mainModule === bundles.eh?.mainModule ? "eh" : "mvp";
  const worker = new Worker(selected.mainWorker!);
  signal?.addEventListener("abort", () => worker.terminate(), { once: true });
  if (runtimeDiagnostics.crossOriginIsolated && runtimeDiagnostics.sharedArrayBuffer) installVgiOAuthBridge(worker);
  const db = new duckdb.AsyncDuckDB(new duckdb.ConsoleLogger(), worker);
  await db.instantiate(selected.mainModule, selected.pthreadWorker);
  // Keep pending-query cancellation reusable on the current threaded WASM build.
  // Multiple execution threads can leave the engine blocked after interruption.
  await db.open({ arrowLosslessConversion: true, maximumThreads: 1 });
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

function waitForConnection<T>(pending: Promise<T>, signal?: AbortSignal): Promise<T> {
  if (!signal) return pending;
  signal.throwIfAborted();
  return new Promise<T>((resolve, reject) => {
    const abort = () => reject(new DOMException("Stopped", "AbortError"));
    signal.addEventListener("abort", abort, { once: true });
    pending.then(value => { signal.removeEventListener("abort", abort); resolve(value); }, error => { signal.removeEventListener("abort", abort); reject(error); });
  });
}
