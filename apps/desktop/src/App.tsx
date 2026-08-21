import { useEffect, useMemo, useRef, useState } from "react";
import { activateQueryDocument, addQueryDocument, formatAttachOptionsJson, loadQueryDocumentState, parseAttachOptionsJson, removeQueryDocument, renameQueryDocument, saveQueryDocumentState, updateQueryDocumentSql, type QueryDocumentState, type QueryResult } from "@query-farm/vgi-excel-core";
import { AgentSession, DEFAULT_MODEL } from "./agent";
import { activateAgentConversation, addAgentConversation, loadAgentConversationState, removeAgentConversation, renameAgentConversation, saveAgentConversationState, titleFromPrompt, type AgentConversationDocument, type AgentConversationState, type ChatMessage, type StagedWorkbookAction, type ToolEvent } from "./agent-conversations";
import { host, type DesktopConnection, type ManagedSnapshot, type PowerQueryOutcome } from "./bridge";
import { ChatMarkdown } from "./ChatMarkdown";
import { AboutDialog, ProductVersion } from "./ProductVersion";
import { QueryTabs } from "./QueryTabs";
import { captureError } from "./telemetry";
import { Notice, Onboarding, TabPanel, WorkspaceTabs, formatSql, resultTsv, type NoticeValue } from "./Ux";
import { Braces, ChevronDown, ChevronLeft, ChevronRight, Copy, Database, Eye, EyeOff, FileCode2, Folder, FolderOpen, History, Play, Plug, Sparkles, Table2, TableProperties, WandSparkles } from "lucide-react";

type Workspace = "query" | "agent" | "catalog" | "connections";
type View = Workspace | "workbook";
type PendingQuery = { id: string; sql: string; name?: string };
export type CatalogDetails = { title: string; metadata?: QueryResult; fields: QueryResult };

export function App(): React.JSX.Element {
  const requested = new URLSearchParams(location.search).get("tab");
  const initial: View = requested === "sql" ? "query" : requested && ["query", "catalog", "agent", "connections", "workbook"].includes(requested) ? requested as View : "query";
  const [view, setView] = useState<View>(initial);
  const [workspace, setWorkspace] = useState<Workspace>(initial === "workbook" ? "query" : initial);
  const [connections, setConnections] = useState<DesktopConnection[]>([]);
  const [notice, setNotice] = useState<NoticeValue>(null);
  const [busy, setBusy] = useState(false);
  const [fatal, setFatal] = useState<string | null>(null);
  const [about, setAbout] = useState(false);
  const [pendingQueries, setPendingQueries] = useState<PendingQuery[]>([]);
  const [diagnostics, setDiagnostics] = useState(`Cupola for Excel ${__APP_VERSION__}\nBuild ${__BUILD_ID__}`);
  const retry = useRef<null | (() => void)>(null);
  const active = connections.find((item) => item.isDefault) ?? connections[0];

  async function refreshConnections(): Promise<void> {
    setNotice(null);
    try { const values = await host.connections(); setConnections(values); setNotice(null); void host.ready().catch(() => undefined); }
    catch (error) { const text = message(error); setNotice({ kind: "error", message: text }); if (/native bridge|did not respond/i.test(text)) setFatal(text); }
  }
  async function perform<T>(label: string, fn: () => Promise<T | string>, success?: string | null): Promise<T | undefined> {
    setBusy(true); setNotice(null); retry.current = () => { void perform(label, fn, success); };
    try { const value = await fn(); return value as T; }
    catch (error) { setNotice({ kind: "error", message: message(error) }); return undefined; }
    finally { setBusy(false); }
  }
  async function openAbout(): Promise<void> {
    setAbout(true);
    try { setDiagnostics(await host.diagnostics()); } catch { }
  }
  async function copyDiagnostics(): Promise<void> { try { await host.copyText(diagnostics); } catch (error) { setNotice({ kind: "error", message: message(error) }); } }
  function openWorkspace(next: Workspace): void { setWorkspace(next); setView(next); }
  function queueQuery(sql: string, name?: string, navigate = false): void { setPendingQueries((values) => [...values, { id: crypto.randomUUID(), sql, name }]); if (navigate) openWorkspace("query"); }
  useEffect(() => { void refreshConnections(); }, []);
  useEffect(() => { window.vgiSelectTab = (value) => { if (["sql", "query", "catalog", "agent", "connections"].includes(value)) openWorkspace(value === "sql" ? "query" : value as Workspace); }; return () => { delete window.vgiSelectTab; }; }, []);

  if (fatal) return <main className="recovery"><section><img src="./cupola-mark.svg" alt=""/><p className="eyebrow">Cupola for Excel</p><h1>Cupola could not connect to Excel</h1><p>{fatal}</p><ol><li>Close this window.</li><li>Close every Excel window.</li><li>Open Excel again and choose Cupola from the ribbon.</li></ol><div className="actions"><button className="primary" onClick={() => { setFatal(null); void refreshConnections(); }}>Try again</button><button onClick={() => void copyDiagnostics()}>Copy diagnostics</button></div></section><ProductVersion onAbout={() => void openAbout()}/>{about && <AboutDialog diagnostics={diagnostics} onClose={() => setAbout(false)} onCopy={() => void copyDiagnostics()}/>}</main>;

  return <main className="app-shell">
    <header><div className="brand"><img className="mark" src="./cupola-mark.svg" alt=""/><h1>Cupola <span>for Excel</span></h1></div><div className="header-actions"><button className="icon-button" aria-label="Workbook data" title="Workbook data" onClick={() => setView("workbook")}><TableProperties aria-hidden="true"/></button></div></header>
    <WorkspaceTabs value={view === "workbook" ? workspace : view} tabs={[{ id: "query", label: "Query Editor", icon: <FileCode2 aria-hidden="true"/> }, { id: "agent", label: "Ask AI", icon: <Sparkles aria-hidden="true"/> }, { id: "catalog", label: "Catalog", icon: <Database aria-hidden="true"/> }, { id: "connections", label: "Connections", icon: <Plug aria-hidden="true"/> }]} onChange={openWorkspace}/>
    <Notice value={notice} onDismiss={() => setNotice(null)} onRetry={notice?.kind === "error" && retry.current ? retry.current : undefined} onDiagnostics={() => void copyDiagnostics()}/>
    {!active && view !== "connections" ? <Onboarding onConnect={() => openWorkspace("connections")}/> : <>
      <TabPanel id="query" active={view === "query"} busy={busy}><SqlPanel key={active.name} connection={active} busy={busy} perform={perform} pendingQueries={pendingQueries} onPendingConsumed={(ids) => setPendingQueries((values) => values.filter((value) => !ids.includes(value.id)))}/></TabPanel>
      <TabPanel id="agent" active={view === "agent"} busy={busy}><AgentPanel key={active.name} connection={active} setBusy={setBusy} setNotice={setNotice} onCreateQuery={(value) => queueQuery(value.sql, value.name)}/></TabPanel>
      <TabPanel id="catalog" active={view === "catalog"} busy={busy}><CatalogPanel active={view === "catalog"} connection={active} busy={busy} perform={perform} openSql={(sql) => queueQuery(sql, undefined, true)}/></TabPanel>
    </>}
    <TabPanel id="connections" active={view === "connections"} busy={busy}><section className="settings-view" aria-labelledby="connections-title"><div className="section-heading"><div><p className="eyebrow">Data sources</p><h2 id="connections-title">Connections</h2></div></div><ConnectionsPanel values={connections} setValues={setConnections} busy={busy} perform={perform}/></section></TabPanel>
    {view === "workbook" && <section className="settings-view" aria-labelledby="workbook-title"><div className="section-heading"><div><p className="eyebrow">Workbook</p><h2 id="workbook-title">Cupola tables</h2></div><button onClick={() => openWorkspace(workspace)}>Back</button></div><DesktopWorkbookPanel busy={busy} perform={perform}/></section>}
    <ProductVersion onAbout={() => void openAbout()}/>
    {about && <AboutDialog diagnostics={diagnostics} onClose={() => setAbout(false)} onCopy={() => void copyDiagnostics()}/>}
  </main>;
}

type Perform = <T>(label: string, fn: () => Promise<T | string>, success?: string | null) => Promise<T | undefined>;
const QUERY_PREVIEW_OPTIONS = [200, 500, 1_000, 2_000];

function connectionLabel(name: string, catalog?: string): string {
  return !catalog || name.localeCompare(catalog, undefined, { sensitivity: "base" }) === 0 ? name : `${name} · ${catalog}`;
}

function QuerySplitter({ value, onChange }: { value: number; onChange(value: number): void }): React.JSX.Element {
  function start(event: React.PointerEvent<HTMLDivElement>): void {
    const container = event.currentTarget.closest<HTMLElement>(".query-editor");
    if (!container) return;
    const bounds = container.getBoundingClientRect();
    const move = (pointer: PointerEvent) => onChange(Math.max(22, Math.min(72, ((pointer.clientY - bounds.top) / bounds.height) * 100)));
    const stop = () => { window.removeEventListener("pointermove", move); window.removeEventListener("pointerup", stop); };
    window.addEventListener("pointermove", move); window.addEventListener("pointerup", stop, { once: true }); event.preventDefault();
  }
  return <div className="query-splitter" role="separator" aria-label="Resize query editor and results" aria-orientation="horizontal" aria-valuemin={22} aria-valuemax={72} aria-valuenow={Math.round(value)} tabIndex={0} title="Drag to resize the query editor and results" onPointerDown={start} onKeyDown={(event) => { if (event.key === "ArrowUp" || event.key === "ArrowDown") { event.preventDefault(); onChange(Math.max(22, Math.min(72, value + (event.key === "ArrowDown" ? 3 : -3)))); } }}><span aria-hidden="true"/></div>;
}

type DesktopQueryRuntime = { result: QueryResult | null; previewLimit: number; previewOffset: number; inserted: { sheet: string; table: string; address: string } | null; powerQuery: PowerQueryOutcome | null; executedSql?: string };
const EMPTY_DESKTOP_QUERY_RUNTIME: DesktopQueryRuntime = { result: null, previewLimit: 200, previewOffset: 0, inserted: null, powerQuery: null };

function SqlPanel({ connection, busy, perform, pendingQueries, onPendingConsumed }: { connection?: DesktopConnection; busy: boolean; perform: Perform; pendingQueries: PendingQuery[]; onPendingConsumed(ids: string[]): void }): React.JSX.Element {
  const scope = connection?.name ?? "default";
  const [documents, setDocuments] = useState<QueryDocumentState>(() => {
    let state = loadQueryDocumentState(scope);
    const legacyDraft = sessionStorage.getItem("cupola.query.draft");
    if (legacyDraft) { state = addQueryDocument(state, legacyDraft); sessionStorage.removeItem("cupola.query.draft"); saveQueryDocumentState(scope, state); }
    return state;
  });
  const [runtimes, setRuntimes] = useState<Record<string, DesktopQueryRuntime>>({});
  const [runningId, setRunningId] = useState<string | null>(null);
  const [editorPercent, setEditorPercent] = useState(() => { const value = Number(localStorage.getItem("cupola.query.editorPercent")); return Number.isFinite(value) && value >= 22 && value <= 72 ? value : 42; });
  const [history, setHistory] = useState<string[]>(() => { try { return JSON.parse(localStorage.getItem("cupola.query.history") ?? "[]") as string[]; } catch { return []; } });
  const activeDocument = documents.documents.find((value) => value.id === documents.activeId) ?? documents.documents[0];
  const runtime = runtimes[activeDocument.id] ?? EMPTY_DESKTOP_QUERY_RUNTIME;
  const { result, previewLimit, previewOffset, inserted, powerQuery } = runtime;
  const sql = activeDocument.sql;
  function updateDocuments(change: (value: QueryDocumentState) => QueryDocumentState): void { setDocuments((previous) => { const next = change(previous); saveQueryDocumentState(scope, next); return next; }); }
  function setSql(value: string): void { updateDocuments((state) => updateQueryDocumentSql(state, state.activeId, value)); }
  function patchRuntime(id: string, value: Partial<DesktopQueryRuntime>): void { setRuntimes((previous) => ({ ...previous, [id]: { ...EMPTY_DESKTOP_QUERY_RUNTIME, ...previous[id], ...value } })); }
  function closeDocument(id: string): void { updateDocuments((state) => removeQueryDocument(state, id)); setRuntimes((previous) => { const { [id]: _removed, ...remaining } = previous; return remaining; }); }
  useEffect(() => { saveQueryDocumentState(scope, documents); }, []);
  useEffect(() => { if (!pendingQueries.length) return; updateDocuments((state) => pendingQueries.reduce((value, pending) => addQueryDocument(value, pending.sql, pending.name), state)); onPendingConsumed(pendingQueries.map((value) => value.id)); }, [pendingQueries]);
  async function run(): Promise<void> {
    if (!connection) return;
    const documentId = activeDocument.id, statement = sql;
    setRunningId(documentId);
    try {
      const value = await perform<QueryResult>("Running query…", () => host.query(statement, connection.name), null);
      if (!value || typeof value === "string") return;
      patchRuntime(documentId, { result: value, inserted: null, powerQuery: null, previewLimit: 200, previewOffset: 0, executedSql: statement });
      setHistory((previous) => { const next = [statement.trim(), ...previous.filter((item) => item !== statement.trim())].slice(0, 12); localStorage.setItem("cupola.query.history", JSON.stringify(next)); return next; });
    } finally { setRunningId((value) => value === documentId ? null : value); }
  }
  async function insert(): Promise<void> {
    if (!result) return;
    const documentId = activeDocument.id, statement = runtime.executedSql ?? sql;
    const needsFullResult = result.truncated || result.rows.length < result.rowCount;
    const value = await perform<{ sheet: string; table: string; address: string }>(needsFullResult ? "Loading full result and inserting snapshot…" : "Inserting snapshot…", () => needsFullResult ? host.insertQuery(statement, connection!.name, "VGI_Result") : host.insert(result, "VGI_Result", { connection: connection!.name, sql: statement }), "Snapshot inserted.");
    if (value && typeof value !== "string") patchRuntime(documentId, { inserted: value });
  }
  async function createPowerQuery(): Promise<void> {
    const documentId = activeDocument.id, statement = runtime.executedSql ?? sql;
    const value = await perform<PowerQueryOutcome>("Creating Power Query…", () => host.createPowerQuery(statement, connection!.name, activeDocument.name, true), null);
    if (value && typeof value !== "string") patchRuntime(documentId, { powerQuery: value });
  }
  useEffect(() => { localStorage.setItem("cupola.query.editorPercent", String(editorPercent)); }, [editorPercent]);
  const previewStart = result?.rows.length ? previewOffset + 1 : 0;
  const previewEnd = result ? Math.min(previewOffset + previewLimit, result.rows.length) : 0;
  return <section className="query-editor">
    <QueryTabs documents={documents.documents} activeId={documents.activeId} runningId={runningId} onSelect={(id) => updateDocuments((state) => activateQueryDocument(state, id))} onAdd={() => updateDocuments((state) => addQueryDocument(state))} onClose={closeDocument} onRename={(id, name) => updateDocuments((state) => renameQueryDocument(state, id, name))}/>
    <div className="query-toolbar" role="toolbar" aria-label="Query editor actions"><button className="primary run-query" disabled={busy || !connection || !sql.trim()} onClick={() => void run()} title="Run query (Ctrl+Enter)"><Play aria-hidden="true"/><span>{busy ? "Running…" : "Run"}</span></button><span className="toolbar-divider" aria-hidden="true"/><button className="toolbar-button" onClick={() => setSql(formatSql(sql))} title="Format SQL"><WandSparkles aria-hidden="true"/><span>Format</span></button><button className="toolbar-button" onClick={() => void host.copyText(sql)} title="Copy SQL"><Copy aria-hidden="true"/><span>Copy SQL</span></button><details className="query-history-menu"><summary title="Query history"><History aria-hidden="true"/><span>History</span>{history.length > 0 && <small>{history.length}</small>}</summary><div className="query-history-popover">{history.length ? history.map((item, index) => <button key={index} onClick={() => setSql(item)} title={item}><code>{item.replace(/\s+/g, " ").slice(0, 90)}</code></button>) : <p>No queries run yet.</p>}</div></details></div>
    <div className="editor-surface" style={{ flexBasis: `${editorPercent}%` }}><label className="sr-only" htmlFor="sql-editor">SQL query</label><textarea id="sql-editor" value={sql} onChange={(event) => setSql(event.target.value)} onKeyDown={(event) => { if ((event.ctrlKey || event.metaKey) && event.key === "Enter") { event.preventDefault(); void run(); } }} spellCheck={false}/></div>
    <div className="editor-statusbar"><span>{connection ? connectionLabel(connection.name, connection.catalog) : "No connection"}</span><span>Ctrl+Enter to run</span></div>
    <QuerySplitter value={editorPercent} onChange={setEditorPercent}/>
    <div className="query-results-pane"><div className="results-toolbar"><div><Table2 aria-hidden="true"/><strong>Results</strong>{result && <small>{result.rowCount.toLocaleString()} rows · showing {previewStart.toLocaleString()}–{previewEnd.toLocaleString()}{result.truncated ? ` · ${result.rows.length.toLocaleString()} loaded for preview` : ""}</small>}</div>{result && <div className="compact-actions">{result.rows.length > previewLimit && <span className="preview-pager"><button aria-label="Previous preview page" title="Previous preview page" disabled={previewOffset === 0} onClick={() => patchRuntime(activeDocument.id, { previewOffset: Math.max(0, previewOffset - previewLimit) })}><ChevronLeft aria-hidden="true"/></button><button aria-label="Next preview page" title="Next preview page" disabled={previewOffset + previewLimit >= result.rows.length} onClick={() => patchRuntime(activeDocument.id, { previewOffset: Math.min(Math.max(0, result.rows.length - 1), previewOffset + previewLimit) })}><ChevronRight aria-hidden="true"/></button></span>}{result.rows.length > QUERY_PREVIEW_OPTIONS[0] && <label className="preview-size">Rows<select aria-label="Rows shown per result page" value={previewLimit} onChange={(event) => patchRuntime(activeDocument.id, { previewLimit: Number(event.target.value), previewOffset: 0 })}>{QUERY_PREVIEW_OPTIONS.map((value) => <option key={value} value={value}>{value.toLocaleString()}</option>)}</select></label>}<button aria-label={result.truncated ? "Copy loaded preview" : "Copy results"} title={result.truncated ? "Copy the rows loaded for preview" : "Copy results"} onClick={() => void host.copyText(resultTsv(result))}><Copy aria-hidden="true"/><span>{result.truncated ? "Copy preview" : "Copy results"}</span></button><button aria-label="Load to Power Query" title="Create a refreshable Excel Power Query using the Cupola ODBC driver" disabled={busy} onClick={() => void createPowerQuery()}><Database aria-hidden="true"/><span>Power Query</span></button><button aria-label="Insert complete snapshot" title="Insert the complete query result as an Excel table" disabled={busy} onClick={() => void insert()}><TableProperties aria-hidden="true"/><span>Insert snapshot</span></button></div>}</div>{powerQuery && <p className={powerQuery.loaded ? "workbook-outcome" : "notice warning"}><strong>{powerQuery.query}</strong> — {powerQuery.message}</p>}{inserted && <WorkbookOutcome value={inserted}/>} {result ? <ResultGrid result={result} label="Query results" limit={previewLimit} offset={previewOffset}/> : <div className="empty-results"><Table2 aria-hidden="true"/><p>Run a query to see results.</p></div>}<p className="microcopy result-disclaimer">Query tabs are saved locally. Power Query results participate in Excel Refresh All; snapshots are point-in-time Excel tables.</p></div>
  </section>;
}

function CatalogPanel({ active, connection, busy, perform, openSql }: { active: boolean; connection?: DesktopConnection; busy: boolean; perform: Perform; openSql(sql: string): void }): React.JSX.Element {
  const [result, setResult] = useState<QueryResult | null>(null);
  const [filter, setFilter] = useState("");
  const [details, setDetails] = useState<CatalogDetails | null>(null);
  const visible = useMemo(() => filterResult(result, filter), [result, filter]);
  async function explore(): Promise<void> {
    if (!connection) return;
    const catalog = sqlLiteral(connection.catalog);
    await perform("Exploring catalog…", async () => {
      setDetails(null);
      const value = await host.query(`SELECT table_catalog AS catalog, table_schema AS schema, table_name AS name, CASE WHEN table_type='VIEW' THEN 'view' ELSE 'table' END AS object_type, table_type AS kind, '' AS summary FROM information_schema.tables WHERE table_catalog=${catalog} AND table_schema NOT IN ('information_schema','pg_catalog') UNION ALL SELECT database_name, schema_name, function_name, CASE WHEN function_type IN ('macro','table_macro') THEN 'macro' ELSE 'function' END, function_type, COALESCE(description, comment, '') FROM duckdb_functions() WHERE database_name=${catalog} ORDER BY 1,2,4,3`, connection.name, false, 20_000);
      setResult(value);
      return `${value.rowCount.toLocaleString()} catalog objects loaded.`;
    });
  }
  async function inspect(row: QueryResult["rows"][number]): Promise<void> {
    if (!connection || !result) return;
    const value = record(result, row), catalog = String(value.catalog ?? ""), schema = String(value.schema ?? ""), name = String(value.name ?? ""), objectType = String(value.object_type ?? "");
    if (!catalog || !schema || !name) return;
    await perform(`Inspecting ${name}…`, async () => {
      const whereFlat = `database_name=${sqlLiteral(catalog)} AND schema_name=${sqlLiteral(schema)} AND function_name=${sqlLiteral(name)}`;
      if (objectType === "function" || objectType === "macro") {
        const [metadata, fields] = await Promise.all([
          host.query(`SELECT function_type, description, comment, return_type, CAST(to_json(parameters) AS VARCHAR) AS parameters, CAST(to_json(parameter_types) AS VARCHAR) AS parameter_types, CAST(to_json(examples) AS VARCHAR) AS examples, CAST(to_json(tags) AS VARCHAR) AS tags, macro_definition FROM duckdb_functions() WHERE ${whereFlat}`, connection.name, false, 1_000),
          host.query(`SELECT arg_position, arg_name, arg_type, CASE WHEN is_named THEN 'named' WHEN is_positional THEN 'positional' WHEN is_varargs THEN 'varargs' ELSE 'other' END AS kind, arg_default, arg_choices, arg_range, arg_pattern, arg_description FROM vgi_function_arguments() WHERE catalog_name=${sqlLiteral(catalog)} AND schema_name=${sqlLiteral(schema)} AND function_name=${sqlLiteral(name)} ORDER BY field_index`, connection.name, false, 1_000),
        ]);
        setDetails({ title: `${catalog}.${schema}.${name}`, metadata, fields });
      } else {
        const fields = await host.query(`SELECT column_name, data_type, is_nullable, column_default, comment FROM duckdb_columns() WHERE database_name=${sqlLiteral(catalog)} AND schema_name=${sqlLiteral(schema)} AND table_name=${sqlLiteral(name)} ORDER BY column_index`, connection.name, false, 10_000);
        setDetails({ title: `${catalog}.${schema}.${name}`, fields });
      }
      return `Showing metadata for ${catalog}.${schema}.${name}.`;
    });
  }
  useEffect(() => { setResult(null); setDetails(null); if (active && connection) void explore(); }, [active, connection?.name]);
  return <section className="catalog-panel"><div className="section-heading"><div><h2>Catalog</h2><p>Browse schemas, tables, functions, signatures, constraints, and examples.</p></div><button disabled={busy || !connection} onClick={() => void explore()}>Refresh</button></div>{result && <label className="catalog-search" htmlFor="catalog-filter">Search catalog<input id="catalog-filter" value={filter} onChange={(event) => setFilter(event.target.value)} placeholder="Function, table, description…"/></label>}<div className="catalog-browser"><CatalogTree result={visible} catalog={connection?.catalog ?? "Catalog"} forceExpanded={!!filter.trim()} onSelect={(row) => void inspect(row)}/><div className="catalog-inspector">{details ? <CatalogDetail value={details} onInsert={(sql) => openSql(sql)} onCopy={(value) => void host.copyText(value)}/> : <div className="empty-detail"><h3>Select a catalog object</h3><p>Choose a table or function to inspect its columns, signature, documentation, and examples.</p></div>}</div></div></section>;
}

type CatalogTreeFolder = { id: string; label: string; order: number; rows: QueryResult["rows"] };
type CatalogTreeSchema = { id: string; name: string; folders: CatalogTreeFolder[] };

export function catalogTreeModel(result: QueryResult | null, includeHidden = false): CatalogTreeSchema[] {
  if (!result) return [];
  const schemas = new Map<string, Map<string, CatalogTreeFolder>>();
  const folder = (kind: string): Omit<CatalogTreeFolder, "rows"> => {
    if (kind === "table") return { id: "tables", label: "Tables", order: 0 };
    if (kind === "view") return { id: "views", label: "Views", order: 1 };
    if (kind === "function") return { id: "functions", label: "Functions", order: 2 };
    if (kind === "macro") return { id: "macros", label: "Macros", order: 3 };
    return { id: "objects", label: "Other objects", order: 4 };
  };
  for (const row of result.rows) {
    const value = record(result, row), name = String(value.name ?? ""), schema = String(value.schema ?? "main"), kind = String(value.object_type ?? value.kind ?? "object").toLowerCase();
    if (!includeHidden && name.includes("$")) continue;
    const definition = folder(kind), folders = schemas.get(schema) ?? new Map<string, CatalogTreeFolder>();
    const current = folders.get(definition.id) ?? { ...definition, rows: [] };
    current.rows.push(row); folders.set(definition.id, current); schemas.set(schema, folders);
  }
  return [...schemas].sort(([left], [right]) => left.localeCompare(right, undefined, { sensitivity: "base" })).map(([name, folders]) => {
    const values = [...folders.values()].sort((left, right) => left.order - right.order || left.label.localeCompare(right.label));
    for (const value of values) value.rows.sort((left, right) => String(record(result, left).name ?? "").localeCompare(String(record(result, right).name ?? ""), undefined, { sensitivity: "base" }));
    return { id: `schema:${name}`, name, folders: values };
  });
}

function TreeBranch({ level, label, expanded, onToggle, children }: { level: number; label: string; expanded: boolean; onToggle(): void; children: React.ReactNode }): React.JSX.Element {
  return <li role="none" className="tree-node"><button type="button" role="treeitem" aria-level={level} aria-expanded={expanded} className="tree-branch" onClick={onToggle} onKeyDown={(event) => { if (event.key === "ArrowRight" && !expanded) { event.preventDefault(); onToggle(); } else if (event.key === "ArrowLeft" && expanded) { event.preventDefault(); onToggle(); } }}><span className="tree-chevron" aria-hidden="true">{expanded ? <ChevronDown/> : <ChevronRight/>}</span><span className="tree-folder" aria-hidden="true">{expanded ? <FolderOpen/> : <Folder/>}</span><span className="tree-label">{label}</span></button>{expanded && <ul role="group" aria-label={label}>{children}</ul>}</li>;
}

export function CatalogTree({ result, catalog, forceExpanded = false, onSelect }: { result: QueryResult | null; catalog: string; forceExpanded?: boolean; onSelect(row: QueryResult["rows"][number]): void }): React.JSX.Element {
  const [showHidden, setShowHidden] = useState(false);
  const model = useMemo(() => catalogTreeModel(result, showHidden), [result, showHidden]);
  const completeModel = useMemo(() => catalogTreeModel(result, true), [result]);
  const hiddenCount = useMemo(() => result?.rows.reduce((count, row) => count + (String(record(result, row).name ?? "").includes("$") ? 1 : 0), 0) ?? 0, [result]);
  const rootId = `catalog:${catalog}`;
  const allBranchIds = useMemo(() => [rootId, ...model.flatMap((schema) => [schema.id, ...schema.folders.map((folder) => `${schema.id}:${folder.id}`)])], [rootId, model]);
  const completeBranchIds = useMemo(() => [rootId, ...completeModel.flatMap((schema) => [schema.id, ...schema.folders.map((folder) => `${schema.id}:${folder.id}`)])], [rootId, completeModel]);
  const [expanded, setExpanded] = useState<Set<string>>(new Set());
  const [selected, setSelected] = useState<string | null>(null);
  const loading = result === null;
  useEffect(() => {
    const initial = new Set<string>([rootId]);
    if (model[0]) { initial.add(model[0].id); for (const folder of model[0].folders) initial.add(`${model[0].id}:${folder.id}`); }
    setExpanded(initial); setSelected(null); setShowHidden(false);
  }, [rootId, loading]);
  if (!result) return <div className="catalog-tree skeleton" aria-label="Catalog loading">Loading catalog…</div>;
  const shown = forceExpanded ? new Set(allBranchIds) : expanded;
  const toggle = (id: string) => setExpanded((current) => { const next = new Set(current); if (next.has(id)) next.delete(id); else next.add(id); return next; });
  const keyboard = (event: React.KeyboardEvent<HTMLDivElement>) => {
    if (!["ArrowDown", "ArrowUp", "Home", "End"].includes(event.key)) return;
    const items = [...event.currentTarget.querySelectorAll<HTMLElement>("[role='treeitem']")];
    const current = items.indexOf(document.activeElement as HTMLElement);
    const index = event.key === "Home" ? 0 : event.key === "End" ? items.length - 1 : Math.max(0, Math.min(items.length - 1, current + (event.key === "ArrowDown" ? 1 : -1)));
    if (items[index]) { event.preventDefault(); items[index].focus(); }
  };
  const rootExpanded = shown.has(rootId);
  return <aside className="catalog-tree" aria-label="Catalog object sidebar"><div className="catalog-tree-toolbar"><div className="tree-toolbar-primary"><strong>Object tree</strong>{hiddenCount > 0 && <button type="button" className="tree-hidden-toggle" aria-pressed={showHidden} title={showHidden ? "Hide catalog items whose names contain $" : `Show ${hiddenCount} catalog items whose names contain $`} onClick={() => { if (showHidden) setShowHidden(false); else { setShowHidden(true); setExpanded(new Set(completeBranchIds)); } }}>{showHidden ? <EyeOff aria-hidden="true"/> : <Eye aria-hidden="true"/>}<span>{showHidden ? "Hide hidden" : `Show hidden (${hiddenCount})`}</span></button>}</div></div><div className="catalog-tree-scroll" role="tree" aria-label={`${catalog} catalog objects`} onKeyDown={keyboard}><ul className="tree-root" role="none"><TreeBranch level={1} label={catalog} expanded={rootExpanded} onToggle={() => toggle(rootId)}>{model.map((schema) => <TreeBranch key={schema.id} level={2} label={schema.name} expanded={shown.has(schema.id)} onToggle={() => toggle(schema.id)}>{schema.folders.map((folder) => { const folderId = `${schema.id}:${folder.id}`; return <TreeBranch key={folderId} level={3} label={folder.label} expanded={shown.has(folderId)} onToggle={() => toggle(folderId)}>{folder.rows.map((row, index) => { const value = record(result, row), kind = String(value.object_type ?? value.kind ?? "object").toLowerCase(), name = String(value.name ?? ""), key = `${schema.name}:${kind}:${name}:${index}`; return <li role="none" key={key}><button type="button" role="treeitem" aria-level={4} aria-selected={selected === key} aria-label={`${name}, ${kind}`} className="tree-leaf" title={`${catalog}.${schema.name}.${name}`} onClick={() => { setSelected(key); onSelect(row); }}><span className="tree-spacer"/><span className={`object-icon ${kind}`} aria-hidden="true">{kind === "table" || kind === "view" ? <Table2/> : <Braces/>}</span><span className="tree-label">{name}</span></button></li>; })}</TreeBranch>; })}</TreeBranch>)}</TreeBranch></ul>{model.length === 0 && <p className="tree-empty">No matching catalog objects.</p>}</div></aside>;
}

export function CatalogDetail({ value, onInsert, onCopy }: { value: CatalogDetails; onInsert?(sql: string): void; onCopy?(value: string): void }): React.JSX.Element {
  if (!value.metadata) { const sql = `SELECT *\nFROM ${quoteQualified(value.title)}\nLIMIT 100;`; return <div className="catalog-detail"><div className="detail-heading"><h3>{value.title}</h3><div className="compact-actions">{onCopy && <button onClick={() => onCopy(value.title)}>Copy name</button>}{onInsert && <button className="primary" onClick={() => onInsert(sql)}>Insert into query</button>}</div></div><h4>Columns</h4><ResultGrid result={value.fields} label={`Columns for ${value.title}`}/></div>; }
  const metadata = value.metadata.rows[0] ? record(value.metadata, value.metadata.rows[0]) : {};
  const tags = jsonObject(metadata.tags), args = value.fields.rows.map((row) => record(value.fields, row));
  const signatureArgs = args.map((arg) => `${String(arg.arg_name ?? "arg")}${arg.kind === "named" ? " := " : " "}${String(arg.arg_type ?? "ANY")}`);
  const fallbackNames = jsonArray(metadata.parameters).map(String), fallbackTypes = jsonArray(metadata.parameter_types).map(String);
  const signature = `${value.title}(${signatureArgs.length ? signatureArgs.join(", ") : fallbackTypes.map((type, index) => `${fallbackNames[index] ?? `arg_${index + 1}`} ${type}`).join(", ")})`;
  const doc = String(tags["vgi.doc_md"] ?? metadata.description ?? metadata.comment ?? "");
  const category = tags["vgi.category"] ? String(tags["vgi.category"]) : null;
  const taggedExamples = jsonArray(parseJsonValue(tags["vgi.example_queries"]));
  const examples = [...jsonArray(metadata.examples).map((sql) => ({ sql: String(sql), description: "" })), ...taggedExamples.map((item) => typeof item === "object" && item ? { sql: String((item as Record<string, unknown>).sql ?? ""), description: String((item as Record<string, unknown>).description ?? "") } : { sql: String(item), description: "" })].filter((item, index, all) => item.sql && all.findIndex((candidate) => candidate.sql === item.sql) === index);
  const resultColumns = jsonArray(parseJsonValue(tags["vgi.result_columns_schema"])).filter((item): item is Record<string, unknown> => !!item && typeof item === "object");
  return <div className="catalog-detail"><div className="detail-heading"><h3>{value.title}</h3><div className="compact-actions">{onCopy && <button onClick={() => onCopy(value.title)}>Copy name</button>}{onInsert && examples[0] && <button className="primary" onClick={() => onInsert(examples[0].sql)}>Run example</button>}</div></div><div className="catalog-badges"><span>{String(metadata.function_type ?? "function")}</span>{category && <span>{category}</span>}{metadata.return_type != null && metadata.return_type !== "" && <span>returns {String(metadata.return_type)}</span>}</div><h4>Signature</h4><pre className="catalog-signature"><code>{signature}</code></pre>{doc && <ChatMarkdown content={doc}/>}<h4>Arguments and constraints</h4><ResultGrid result={value.fields} label={`Arguments for ${value.title}`}/>{resultColumns.length > 0 && <><h4>Result columns</h4><ResultGrid result={objectResult(resultColumns, ["name", "type", "description"])} label={`Result columns for ${value.title}`}/></>}{examples.length > 0 && <div className="catalog-examples"><h4>Examples</h4>{examples.map((example, index) => <div key={index}>{example.description && <p>{example.description}</p>}<pre className="markdown-code"><button className="copy-code" onClick={() => onCopy?.(example.sql)}>Copy SQL</button><code>{example.sql}</code></pre>{onInsert && <button onClick={() => onInsert(example.sql)}>Open in Query</button>}</div>)}</div>}</div>;
}

function AgentPanel({ connection, setBusy, setNotice, onCreateQuery }: { connection?: DesktopConnection; setBusy(value: boolean): void; setNotice(value: NoticeValue): void; onCreateQuery(value: { name: string; sql: string }): void }): React.JSX.Element {
  const scope = connection?.name ?? "default";
  const [key, setKey] = useState("");
  const [keyStored, setKeyStored] = useState(false);
  const [keyBusy, setKeyBusy] = useState(false);
  const [conversations, setConversations] = useState<AgentConversationState>(() => loadAgentConversationState(scope, DEFAULT_MODEL));
  const [controller, setController] = useState<AbortController | null>(null);
  const [runningId, setRunningId] = useState<string | null>(null);
  const [settings, setSettings] = useState(false);
  const sessions = useRef(new Map<string, AgentSession>());
  const scroll = useRef<HTMLDivElement>(null);
  const activeConversation = conversations.documents.find((value) => value.id === conversations.activeId) ?? conversations.documents[0];
  const messages = activeConversation.displayMessages;
  const prompt = activeConversation.draft;
  const model = activeConversation.model;
  const staged = activeConversation.staged ?? null;
  const setError = (error: unknown) => setNotice({ kind: "error", message: message(error) });

  function sessionFor(value: AgentConversationDocument): AgentSession {
    let session = sessions.current.get(value.id);
    if (!session) { session = new AgentSession(); session.restore(value.agentMessages); sessions.current.set(value.id, session); }
    return session;
  }
  function updateConversation(id: string, update: (value: AgentConversationDocument) => AgentConversationDocument): void {
    setConversations((state) => ({ ...state, documents: state.documents.map((value) => value.id === id ? { ...update(value), updatedAt: Date.now() } : value) }));
  }
  function addConversation(): void { if (!controller) setConversations((state) => addAgentConversation(state, model || DEFAULT_MODEL)); }
  function closeConversation(id: string): void {
    if (id === runningId) return;
    sessions.current.delete(id);
    setConversations((state) => removeAgentConversation(state, id, DEFAULT_MODEL));
  }

  useEffect(() => {
    let mounted = true;
    void host.loadAgentKey().then((value) => {
      if (!mounted || !value) return;
      setKey(value);
      setKeyStored(true);
    }).catch((error) => { if (mounted) setNotice({ kind: "error", message: message(error) }); });
    return () => { mounted = false; };
  }, [setNotice]);
  useEffect(() => { saveAgentConversationState(scope, conversations); }, [scope, conversations]);
  useEffect(() => { scroll.current?.scrollTo({ top: scroll.current.scrollHeight, behavior: "smooth" }); }, [messages]);

  async function ask(): Promise<void> {
    if (!connection || !key.trim() || !prompt.trim()) return;
    const conversationId = activeConversation.id;
    const question = prompt.trim();
    const session = sessionFor(activeConversation);
    updateConversation(conversationId, (value) => ({
      ...value,
      name: value.displayMessages.length === 0 && /^Conversation \d+$/.test(value.name) ? titleFromPrompt(question) : value.name,
      draft: "",
      displayMessages: [...value.displayMessages, { role: "user", text: question }, { role: "assistant", text: "", tools: [], workbookActions: [], streaming: true, activity: "Thinking…" }],
    }));
    const abort = new AbortController(); setController(abort); setRunningId(conversationId); setBusy(true); setNotice(null);
    const updateLast = (fn: (value: ChatMessage) => ChatMessage) => updateConversation(conversationId, (value) => ({ ...value, displayMessages: value.displayMessages.map((message, index) => index === value.displayMessages.length - 1 ? fn(message) : message) }));
    try {
      await session.run(key, model, question, connection, {
        onText: (chunk) => updateLast((value) => ({ ...value, text: value.text + chunk, activity: undefined })),
        onTool: (name, state, detail, callId) => updateLast((value) => {
          const tools = [...(value.tools ?? [])]; const id = callId ?? crypto.randomUUID(); const prior = tools.findIndex((item) => item.id === id);
          if (prior >= 0) { const priorValue = tools[prior]; tools[prior] = { ...priorValue, state, detail, elapsedMs: state === "done" || state === "error" ? Math.round(performance.now() - (priorValue.startedAt ?? performance.now())) : priorValue.elapsedMs }; } else tools.push({ id, name, state, detail, startedAt: performance.now() });
          return { ...value, tools, activity: state === "done" ? "Reviewing results…" : undefined };
        }),
        onRetry: (value) => updateLast((message) => ({ ...message, activity: value ?? "Thinking…" })),
        onResult: (result) => updateConversation(conversationId, (value) => ({ ...value, staged: result })),
        onWorkbookAction: (action) => updateLast((value) => ({ ...value, workbookActions: [...(value.workbookActions ?? []), { ...action, status: "pending" }] })),
        onQueryDocument: onCreateQuery,
      }, abort.signal);
    } catch (error) {
      if ((error as Error).name === "AbortError") updateLast((value) => ({ ...value, stopped: true, activity: undefined }));
      else { captureError(error, "agent.run"); setError(error); updateLast((value) => ({ ...value, text: value.text || `Error: ${message(error)}` })); }
    } finally {
      updateConversation(conversationId, (value) => ({ ...value, agentMessages: session.snapshot(), displayMessages: value.displayMessages.map((message, index) => index === value.displayMessages.length - 1 ? { ...message, streaming: false, activity: undefined } : message) }));
      setController(null); setRunningId(null); setBusy(false);
    }
  }
  function reset(): void {
    sessionFor(activeConversation).reset();
    updateConversation(activeConversation.id, (value) => ({ ...value, displayMessages: [], agentMessages: [], staged: null }));
    setNotice(null);
  }
  async function saveKey(): Promise<void> {
    setKeyBusy(true); setNotice(null);
    try { await host.saveAgentKey(key); setKeyStored(true); setSettings(false); }
    catch (error) { setError(error); }
    finally { setKeyBusy(false); }
  }
  async function forgetKey(): Promise<void> {
    setKeyBusy(true); setNotice(null);
    try { await host.deleteAgentKey(); setKey(""); setKeyStored(false); }
    catch (error) { setError(error); }
    finally { setKeyBusy(false); }
  }

  function updateWorkbookAction(messageIndex: number, id: string, patch: Partial<StagedWorkbookAction>): void {
    updateConversation(activeConversation.id, (conversation) => ({ ...conversation, displayMessages: conversation.displayMessages.map((value, index) => index !== messageIndex ? value : {
      ...value, workbookActions: value.workbookActions?.map((action) => action.id === id ? { ...action, ...patch } : action),
    }) }));
  }

  const workbookActions = messages.flatMap((message, messageIndex) => (message.workbookActions ?? []).map((action) => ({ action, messageIndex })));
  const unresolvedActions = workbookActions.filter(({ action }) => action.status !== "done");
  const visibleActions = unresolvedActions.length ? unresolvedActions : workbookActions.slice(-1);

  return <section className="agent"><QueryTabs documents={conversations.documents} activeId={activeConversation.id} runningId={runningId} label="AI conversations" itemName="conversation" onSelect={(id) => setConversations((state) => activateAgentConversation(state, id))} onAdd={addConversation} onClose={closeConversation} onRename={(id, name) => setConversations((state) => renameAgentConversation(state, id, name))}/><div className="agent-toolbar"><div><span className={`health-dot ${keyStored ? "configured" : ""}`} aria-hidden="true"/><span>{keyStored ? "Conversations saved locally" : "AI setup required"}</span></div><div className="compact-actions"><button disabled={!messages.length || !!controller} onClick={reset}>Clear current</button><button aria-expanded={settings} onClick={() => setSettings(!settings)}>AI settings</button></div></div>{settings && <div className="agent-settings"><div className="credential"><label htmlFor="anthropic-key">Anthropic API key</label><div className="key-row"><input id="anthropic-key" type="password" value={key} onChange={(event) => { setKey(event.target.value); setKeyStored(false); }} autoComplete="off" placeholder="Enter an Anthropic API key"/><button disabled={keyBusy || !!controller || !key.trim() || keyStored} onClick={() => void saveKey()}>Save securely</button>{keyStored && <button disabled={keyBusy || !!controller} onClick={() => void forgetKey()}>Forget</button>}</div><small>{keyStored ? "Protected by Windows Credential Manager; conversation history is stored separately" : "The key is not stored"}</small></div><label>Model<input value={model} onChange={(event) => updateConversation(activeConversation.id, (value) => ({ ...value, model: event.target.value }))}/></label></div>}<div className="chat" ref={scroll}>{!messages.length && <div className="empty"><img className="empty-mark" src="./cupola-mark.svg" alt=""/><h2>Ask AI about your data</h2><p>This conversation is saved locally for this VGI connection. The agent can explore schemas, test read-only SQL, and stage reviewed workbook changes.</p>{!keyStored && <button onClick={() => setSettings(true)}>Configure AI</button>}</div>}{messages.map((item, index) => <article key={index} className={`message ${item.role}`}><strong>{item.role === "user" ? "You" : "Cupola"}</strong>{item.streaming && !item.text && !item.tools?.some((tool) => tool.state === "writing" || tool.state === "running") && <AgentActivity value={item.activity ?? "Thinking…"}/>} {item.tools?.map((tool) => <ToolCall key={tool.id} value={tool}/>)}{item.text && (item.role === "assistant" ? <ChatMarkdown content={item.text} streaming={item.streaming}/> : <div className="message-text">{item.text}</div>)}{item.stopped && <div className="agent-stopped" role="status">Stopped</div>}</article>)}</div>{visibleActions.length > 0 && <div className="workbook-action-tray" aria-label="Workbook actions"><div className="workbook-action-tray-heading"><strong>Ready for Excel</strong><small>Review before changing your workbook</small></div>{visibleActions.map(({ action, messageIndex }) => <WorkbookActionCard key={action.id} value={action} onChange={(patch) => updateWorkbookAction(messageIndex, action.id, patch)}/>)}</div>}<div className="composer"><label className="sr-only" htmlFor="agent-prompt">Ask AI</label><textarea id="agent-prompt" value={prompt} onChange={(event) => updateConversation(activeConversation.id, (value) => ({ ...value, draft: event.target.value }))} onKeyDown={(event) => { if (event.key === "Enter" && !event.shiftKey) { event.preventDefault(); void ask(); } }} rows={3} placeholder="Ask a question about the selected VGI catalog…"/><div className="actions"><button className="primary" disabled={!!controller || !connection || !key.trim() || !prompt.trim()} onClick={() => void ask()}>{controller ? "Working…" : "Send"}</button>{controller && runningId === activeConversation.id && <button className="danger" onClick={() => controller.abort()}>Stop</button>}{controller && runningId !== activeConversation.id && runningId && <button onClick={() => setConversations((state) => activateAgentConversation(state, runningId))}>View running</button>}<InsertButton result={staged} disabled={!!controller}/></div></div></section>;
}

function DesktopWorkbookPanel({ busy, perform }: { busy: boolean; perform: Perform }): React.JSX.Element {
  const [snapshots, setSnapshots] = useState<ManagedSnapshot[]>([]);
  async function load(): Promise<void> { setSnapshots(await host.snapshots()); }
  useEffect(() => { void load(); }, []);
  return <div className="workbook-panel"><p>Cupola-managed tables remember their HTTPS VGI connection and source SQL. Refresh them here or from <strong>Cupola → Refresh Cupola tables</strong> on the ribbon.</p><ul className="card-list">{snapshots.map((snapshot) => <li key={snapshot.table}><div><strong>{snapshot.table}</strong><small>{snapshot.connection} · updated {new Date(snapshot.updatedAt).toLocaleString()}</small><code title={snapshot.sql}>{snapshot.sql.slice(0, 110)}{snapshot.sql.length > 110 ? "…" : ""}</code></div><div className="row-actions"><button disabled={busy} onClick={() => void perform("Refreshing snapshot…", async () => { await host.refreshSnapshot(snapshot.table); await load(); return `${snapshot.table} refreshed.`; })}>Refresh</button><button onClick={() => void host.activateTable(snapshot.table)}>Go to table</button><button onClick={() => void perform("Removing refresh metadata…", async () => { await host.forgetSnapshot(snapshot.table); await load(); return "Refresh metadata removed; the Excel table was kept."; })}>Stop managing</button></div></li>)}{!snapshots.length && <li><div><strong>No managed snapshots yet</strong><small>Run a query and choose “Insert snapshot as Excel table.”</small></div></li>}</ul><p className="hint">Cupola refresh is intentionally separate from Power Query and Excel Refresh All. The workbook stores source SQL and a connection name, never OAuth tokens or API keys.</p></div>;
}

function ConnectionsPanel({ values, setValues, busy, perform }: { values: DesktopConnection[]; setValues(v: DesktopConnection[]): void; busy: boolean; perform: Perform }): React.JSX.Element {
  const blank = (): DesktopConnection => ({ name: "", catalog: "", location: "", authentication: "anonymous", attachOptions: {} });
  const active = values.find((value) => value.isDefault) ?? values[0];
  const [form, setForm] = useState<DesktopConnection>(() => active ? { ...active } : blank());
  const [originalName, setOriginalName] = useState(active?.name ?? "");
  const [attachOptionsText, setAttachOptionsText] = useState(() => formatAttachOptionsJson(active?.attachOptions));
  const [status, setStatus] = useState("");
  useEffect(() => { if (!originalName && active) { setForm({ ...active }); setOriginalName(active.name); setAttachOptionsText(formatAttachOptionsJson(active.attachOptions)); } }, [active?.name]);
  function edit(value: DesktopConnection): void { setForm({ ...value }); setOriginalName(value.name); setAttachOptionsText(formatAttachOptionsJson(value.attachOptions)); setStatus(""); }
  function create(): void { setForm(blank()); setOriginalName(""); setAttachOptionsText(""); setStatus(""); }
  function definition(): DesktopConnection { return { ...form, name: form.name.trim(), catalog: form.catalog.trim(), location: form.location.trim(), attachOptions: parseAttachOptionsJson(attachOptionsText) }; }
  async function run(label: string, success: string, fn: () => Promise<DesktopConnection[]>): Promise<void> {
    setStatus(label);
    const next = await perform<DesktopConnection[]>(label, fn, null);
    if (!next || typeof next === "string") { setStatus(""); return; }
    setValues(next);
    const updated = next.find((value) => value.name === form.name);
    if (updated) { setForm({ ...updated }); setAttachOptionsText(formatAttachOptionsJson(updated.attachOptions)); }
    setStatus(success);
  }
  async function test(): Promise<void> {
    setStatus("Testing connection…");
    const result = await perform<QueryResult>("Testing connection…", () => host.testConnection(definition()), null);
    if (!result || typeof result === "string") { setStatus(""); return; }
    const next = await host.connections();
    setValues(next);
    const updated = next.find((value) => value.name === form.name);
    if (updated) { setForm({ ...updated }); setAttachOptionsText(formatAttachOptionsJson(updated.attachOptions)); }
    setStatus(updated?.isSignedIn ? "Connected · signed in securely." : "Connected · no sign-in required.");
  }
  let attachOptionsError = "";
  try { parseAttachOptionsJson(attachOptionsText); } catch (error) { attachOptionsError = message(error); }
  const valid = !!form.name.trim() && !!form.catalog.trim() && /^https:\/\//i.test(form.location.trim()) && !attachOptionsError;
  return <div className="connection-layout">
    <aside className="connection-sidebar"><button className="primary new-connection" onClick={create}>New connection</button><div className="connection-list">{values.map((value) => <button key={value.name} className={`${originalName === value.name ? "connection selected-connection" : "connection"} ${value.isDefault ? "active-connection" : ""}`} onClick={() => edit(value)} title={value.location}><span className={`health-dot ${value.authentication === "anonymous" || value.isSignedIn ? "configured" : ""}`}/><span><strong>{value.name}{value.isDefault ? " · active" : ""}</strong><small>{value.catalog} · {value.isSignedIn ? "Signed in" : value.authentication === "oauth" ? "Sign-in opens when needed" : "Ready"}</small></span></button>)}{!values.length && <p>No connections configured yet.</p>}</div></aside>
    <div className="connection-form">
      <h3>{originalName ? `Edit ${originalName}` : "New HTTPS VGI connection"}</h3>
      <label>Connection name<input value={form.name} onChange={(event) => { setForm({ ...form, name: event.target.value }); setStatus(""); }} aria-describedby="connection-name-help"/></label>
      <small id="connection-name-help">A local identifier used by formulas, saved snapshots, and active-connection selection.</small>
      <label>VGI catalog<input value={form.catalog} onChange={(event) => { setForm({ ...form, catalog: event.target.value }); setStatus(""); }} placeholder="open_meteo"/></label>
      <small>The catalog is the alias used by DuckDB when VGI is attached.</small>
      <label>HTTPS endpoint<input value={form.location} onChange={(event) => { setForm({ ...form, location: event.target.value, authentication: "anonymous", isSignedIn: false }); setStatus(""); }} placeholder="https://vgi.example.com/"/></label>
      <small className="connection-auth-hint">If this service requires authentication, Cupola opens your browser when it connects.</small>
      <details className="advanced connection-advanced"><summary>Advanced ATTACH options</summary><label htmlFor="desktop-attach-options">Options as JSON<textarea id="desktop-attach-options" value={attachOptionsText} onChange={(event) => { setAttachOptionsText(event.target.value); setStatus(""); }} rows={4} spellCheck={false} placeholder={'{"region":"us-east"}'}/></label><small>Non-secret string, number, boolean, or null values. Cupola manages TYPE, LOCATION, and OAuth credentials.</small>{attachOptionsError && <p className="field-error" role="alert">{attachOptionsError}</p>}</details>
      <div className="actions connection-actions"><button disabled={busy || !valid} onClick={() => void test()}>{busy && status.startsWith("Testing") ? "Testing…" : "Test connection"}</button><button className="primary" disabled={busy || !valid} onClick={() => void run("Saving connection…", "Connection saved.", async () => { const prepared = definition(); let next = await host.saveConnection(prepared, values.length === 0); if (originalName && originalName !== prepared.name) next = await host.removeConnection(originalName); setOriginalName(prepared.name); return next; })}>Save changes</button>{originalName && !form.isDefault && <button disabled={busy} onClick={() => void run("Switching connection…", `${form.name} is now active.`, () => host.useConnection(form.name))}>Use as active</button>}</div>
      {status && <p className="connection-status" role="status">{status}</p>}
      {form.isSignedIn && <div className="oauth-card"><div><strong>Signed in securely</strong><small>The refresh session is encrypted for your Windows account. Cupola will reuse it automatically.</small></div><button disabled={busy} onClick={() => void run("Signing out…", "Signed out. Cupola will prompt again if this service requires authentication.", () => host.signOut(form))}>Sign out</button></div>}
      {originalName && <div className="danger-zone"><button className="danger" disabled={busy} onClick={() => { if (!window.confirm(`Remove “${originalName}” and its saved OAuth session?`)) return; void run("Removing connection…", "Connection and saved OAuth session removed.", async () => { const next = await host.removeConnection(originalName); setForm(blank()); setOriginalName(""); setAttachOptionsText(""); return next; }); }}>Remove connection</button></div>}
      <p className="hint">Connection definitions contain no credentials. Cupola connects only to HTTPS VGI services; local commands and subprocess connectors are not supported.</p>
    </div>
  </div>;
}

function AgentActivity({ value }: { value: string }): React.JSX.Element { return <div className="agent-activity" role="status" aria-live="polite"><span className="agent-activity-dot" aria-hidden="true"/><span>{value}</span></div>; }
function ToolCall({ value }: { value: ToolEvent }): React.JSX.Element {
  const sql = value.name === "run_sql" && value.detail;
  return <details className={`tool ${value.state}`} open={value.state === "error"}><summary><span className="tool-dot"/> {toolLabel(value.name, value.state)}{value.elapsedMs != null && <small> · {(value.elapsedMs / 1000).toFixed(1)}s</small>}</summary>{value.detail && <div className="tool-detail">{sql && <div className="tool-detail-heading"><span>Executed SQL</span><button type="button" onClick={() => void host.copyText(value.detail!)}>Copy SQL</button></div>}<pre><code>{value.detail}</code></pre></div>}</details>;
}
function WorkbookActionCard({ value, onChange }: { value: StagedWorkbookAction; onChange(patch: Partial<StagedWorkbookAction>): void }): React.JSX.Element {
  const destinationKind = value.mode === "new_sheet" ? "New worksheet" : "Existing table";
  const destinationName = value.mode === "new_sheet" ? value.sheetName ?? "VGI Result" : value.tableName;
  async function apply(): Promise<void> {
    if (value.mode === "replace_table" && !window.confirm(`Replace all data in Excel table “${value.tableName}” with ${value.result.rowCount.toLocaleString()} rows?`)) return;
    onChange({ status: "writing", detail: undefined });
    try {
      const result = await host.writeResult({ mode: value.mode, result: value.result, sheetName: value.sheetName, tableName: value.tableName });
      onChange({ status: "done", detail: `${result.sheet}!${result.address} · ${result.table}` });
    } catch (error) { onChange({ status: "error", detail: message(error) }); }
  }
  return <div className={`workbook-action ${value.status}`}><div className="workbook-action-copy"><strong>{value.mode === "new_sheet" ? "Create Excel table snapshot" : "Update Excel table snapshot"}</strong><div className="workbook-action-summary"><span>{value.result.rowCount.toLocaleString()} rows</span><span aria-hidden="true">→</span><span>{destinationKind}</span><strong>“{destinationName}”</strong></div><p>Manual snapshot · not included in Excel Refresh All</p>{value.detail && <small>{value.detail}</small>}</div><div className="workbook-action-buttons">{value.status === "done" ? <button onClick={() => void host.activateTable(value.tableName)}>Go to table</button> : <button className="primary" disabled={value.status === "writing"} onClick={() => void apply()}>{value.status === "pending" ? "Confirm" : value.status === "writing" ? "Writing…" : "Try again"}</button>}</div></div>;
}
function InsertButton({ result, disabled }: { result: QueryResult | null; disabled: boolean }): React.JSX.Element { const [inserting, setInserting] = useState(false); const [outcome, setOutcome] = useState<{ sheet: string; table: string; address: string } | null>(null); return <>{<button disabled={disabled || !result || inserting} onClick={() => { setInserting(true); void host.insert(result!, "VGI_Result").then(setOutcome).finally(() => setInserting(false)); }}>{inserting ? "Inserting…" : "Insert snapshot"}</button>}{outcome && <button onClick={() => void host.activateTable(outcome.table)}>Go to {outcome.table}</button>}</>; }
function WorkbookOutcome({ value }: { value: { sheet: string; table: string; address: string } }): React.JSX.Element { return <div className="workbook-outcome" role="status"><div><strong>Snapshot inserted</strong><span>{value.table} · {value.sheet}!{value.address}</span></div><button onClick={() => void host.activateTable(value.table)}>Go to table</button></div>; }
function ResultGrid({ result, onRow, hint, label = "Results", onCopy, limit = 200, offset = 0 }: { result: QueryResult | null; onRow?(row: QueryResult["rows"][number]): void; hint?: string; label?: string; onCopy?(): void | Promise<unknown>; limit?: number; offset?: number }): React.JSX.Element | null { if (!result) return null; const rows = result.rows.slice(offset, offset + limit); return <div className="result-wrap"><div className="result-meta"><span>{result.rowCount.toLocaleString()} rows{result.elapsedMs != null ? ` · ${Math.round(result.elapsedMs)} ms` : ""}{result.rowCount > rows.length ? ` · showing ${rows.length}` : ""}{hint ? ` · ${hint}` : ""}</span>{onCopy && <button onClick={() => void onCopy()}>Copy results</button>}</div><table aria-label={label}><caption className="sr-only">{label}</caption><thead><tr>{result.columns.map((c) => <th key={c.name} scope="col"><span>{c.name}</span><small>{c.type}</small></th>)}</tr></thead><tbody>{rows.map((row, ri) => <tr key={offset + ri} className={onRow ? "selectable-row" : undefined} tabIndex={onRow ? 0 : undefined} onClick={onRow ? () => onRow(row) : undefined} onKeyDown={onRow ? (event) => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); onRow(row); } } : undefined}>{row.map((cell, ci) => <td key={ci}>{cell == null ? <em>NULL</em> : String(cell)}</td>)}</tr>)}</tbody></table></div>; }

function message(error: unknown): string { return error instanceof Error ? error.message : String(error); }
function sqlLiteral(value: string): string { return `'${value.replaceAll("'", "''")}'`; }
function record(result: QueryResult, row: QueryResult["rows"][number]): Record<string, unknown> { return Object.fromEntries(result.columns.map((column, index) => [column.name, row[index]])); }
function filterResult(result: QueryResult | null, filter: string): QueryResult | null { if (!result || !filter.trim()) return result; const needle = filter.trim().toLocaleLowerCase(); const rows = result.rows.filter((row) => row.some((cell) => String(cell ?? "").toLocaleLowerCase().includes(needle))); return { ...result, rows, rowCount: rows.length, truncated: false }; }
function parseJsonValue(value: unknown): unknown { if (typeof value !== "string") return value; try { return JSON.parse(value); } catch { return value; } }
function jsonArray(value: unknown): unknown[] { const parsed = parseJsonValue(value); return Array.isArray(parsed) ? parsed : []; }
function jsonObject(value: unknown): Record<string, unknown> { const parsed = parseJsonValue(value); return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed as Record<string, unknown> : {}; }
function objectResult(values: Record<string, unknown>[], columns: string[]): QueryResult { return { columns: columns.map((name) => ({ name, type: "VARCHAR" })), rows: values.map((value) => columns.map((name) => value[name] == null ? null : String(value[name]))), rowCount: values.length, truncated: false }; }
function quoteQualified(value: string): string { return value.split(".").map((part) => `"${part.replaceAll('"', '""')}"`).join("."); }
function toolLabel(name: string, state: ToolEvent["state"]): string { const labels: Record<string, string> = { run_sql: "SQL query", read_query_results: "Reading query results", list_tables: "Catalog inventory", list_functions: "Function inventory", describe_table: "Table description", create_query_tab: "Creating query tab", workbook_overview: "Workbook overview", read_range: "Reading worksheet range", list_formulas: "Formula inventory", stage_result_to_new_sheet: "Staging new worksheet", stage_result_to_table: "Staging table update" }; const suffix = state === "writing" ? " · preparing" : state === "running" ? " · running" : state === "error" ? " · error" : " · complete"; return (labels[name] ?? name) + suffix; }
