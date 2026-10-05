import { reauthenticateConnection } from "./reauthenticate";
import { AuthRecovery } from "./AuthRecovery";
import { connectionSignIn, ConnectionSignInError } from "@query-farm/vgi-excel-core";
import { confirmAction } from "./confirmation";
import { appendTranscript } from "@query-farm/vgi-excel-core";
import { AIRequestDiagnostics, AgentTranscript, AgentResult, AgentProgress, Clarification, useClarification } from "./AgentExtras";
import { toolStage } from "@query-farm/vgi-excel-core";
import { ModelCatalog, useModelCatalog } from "./ModelCatalog";
import { modelEfforts, clampMaxTokens } from "@query-farm/vgi-excel-core";
import { createPortal } from "react-dom";
import { SettingsPage, type SettingsSection } from "./SettingsPage";
import { openResultsWindow } from "./results-window";
import { catalogDraft, connectionNameError, catalogDiscoveryError } from "@query-farm/vgi-excel-core";
import { AI_MODELS, AI_EFFORT_LEVELS, DEFAULT_AI_MODEL, normalizeEffort, supportsEffort } from "@query-farm/vgi-excel-core";
import { useEffect, useMemo, useRef, useState, useSyncExternalStore } from "react";
import { activateQueryDocument, addQueryDocument, assertHttpsConnection, formatAttachOptionsJson, loadQueryDocumentState, parseAttachOptionsJson, removeQueryDocument, renameQueryDocument, saveQueryDocumentState, updateQueryDocumentSql, type CatalogFunction, type ConnectionDefinition, type QueryDocumentState, type QueryResult } from "@query-farm/vgi-excel-core";
import { OfficeAgentSession } from "./anthropic";
import { activateOfficeAgentConversation, addOfficeAgentConversation, loadOfficeAgentConversationState, officeConversationTitle, removeOfficeAgentConversation, renameOfficeAgentConversation, saveOfficeAgentConversationState, type OfficeAgentConversation, type OfficeAgentConversationState, type OfficeChatMessage, type OfficeToolEvent as ToolEvent } from "./agent-conversations";
import { agentContext, discoverFunctions, rows } from "./catalog";
import { ChatMarkdown } from "./ChatMarkdown";
import { AboutContent, AboutDialog } from "./ProductVersion";
import { ResultsMore, Notice, Onboarding, TabPanel, WorkspaceTabs, formatSql, resultTsv, type NoticeValue } from "./Ux";
import { getDefaultConnectionName, getServiceToken, loadConnections, saveConnections, sessionTokenKey, setDefaultConnectionName } from "./config";
import { resetRuntime, resolveBackend } from "./runtime";
import { BrowserBackend, browserRuntimeDiagnostics } from "./browser-backend";
import { officeTreeModel, type CatalogObject } from "./catalog-tree";
import { createFunctionWrappers, forgetSnapshot, goToTable, importSelection, importWorkbookTable, insertResult, listManagedSnapshots, listWorkbookTables, refreshSnapshot, type ManagedSnapshot, type WorkbookWriteOutcome } from "./workbook";
import { EXCEL_MAX_DATA_ROWS } from "./excel-limits";
import { QueryTabs } from "./QueryTabs";
import { captureError } from "./telemetry";
import { LoaderCircle, RefreshCw, Braces, ChevronUp, ChevronDown, ChevronLeft, ChevronRight, Copy, Database, Eye, EyeOff, FileCode2, Folder, FolderOpen, Play, Settings, Sparkles, Square, Maximize2, Table2, TableProperties, WandSparkles } from "lucide-react";

type Workspace = "query" | "agent" | "catalog" | "connections";
type View = Workspace | "workbook";
type OfficeConnection = ConnectionDefinition & { authentication?: "anonymous" | "oauth" };
type PendingQuery = { id: string; sql: string; name?: string };
type Action = <T>(label: string, fn: () => Promise<T>, success?: string | null) => Promise<T | undefined>;
const QUERY_PREVIEW_OPTIONS = [200, 500, 1_000, 2_000];

function connectionLabel(name: string, catalog?: string): string {
  return !catalog || name.localeCompare(catalog, undefined, { sensitivity: "base" }) === 0 ? name : `${name} · ${catalog}`;
}

function QuerySplitter({ value, onChange }: { value: number; onChange(value: number): void }): React.JSX.Element {
  function start(event: React.PointerEvent<HTMLDivElement>): void {
    const container = event.currentTarget.closest<HTMLElement>(".query-editor"); if (!container) return;
    const bounds = container.getBoundingClientRect();
    const move = (pointer: PointerEvent) => onChange(Math.max(22, Math.min(72, ((pointer.clientY - bounds.top) / bounds.height) * 100)));
    const stop = () => { window.removeEventListener("pointermove", move); window.removeEventListener("pointerup", stop); };
    window.addEventListener("pointermove", move); window.addEventListener("pointerup", stop, { once: true }); event.preventDefault();
  }
  return <div className="query-splitter" role="separator" aria-label="Resize query editor and results" aria-orientation="horizontal" aria-valuemin={22} aria-valuemax={72} aria-valuenow={Math.round(value)} tabIndex={0} title="Drag to resize the query editor and results" onPointerDown={start} onKeyDown={(event) => { if (event.key === "ArrowUp" || event.key === "ArrowDown") { event.preventDefault(); onChange(Math.max(22, Math.min(72, value + (event.key === "ArrowDown" ? 3 : -3)))); } }}><span aria-hidden="true"/></div>;
}

export function App(): React.JSX.Element {
  const [view, setView] = useState<View>("query");
  const [workspace, setWorkspace] = useState<Workspace>("query");
  const [settingsSection, setSettingsSection] = useState<SettingsSection>("connections");
  const [aiSettingsTarget, setAISettingsTarget] = useState<HTMLDivElement | null>(null);
  function openAISettings(): void { setSettingsSection("ai"); setView("connections"); }
  const [connections, setConnections] = useState<OfficeConnection[]>(() => loadConnections() as OfficeConnection[]);
  const [defaultName, setDefaultName] = useState(getDefaultConnectionName() ?? "");
  const [notice, setNotice] = useState<NoticeValue>(null);
  const [busy, setBusy] = useState(false);
  const [about, setAbout] = useState(false);
  const [pendingQueries, setPendingQueries] = useState<PendingQuery[]>([]);
  const retry = useRef<null | (() => void)>(null);
  const active = useMemo(() => connections.find((item) => item.name === defaultName) ?? connections[0], [connections, defaultName]);

  const action: Action = async (label, fn, success = "Done.") => {
    setBusy(true); setNotice(null); retry.current = () => { void action(label, fn, success); };
    try { return await fn(); }
    catch (error) { captureError(error, "ui.action"); if (!(error instanceof ConnectionSignInError)) setNotice({ kind: "error", message: message(error) }); return undefined; }
    finally { setBusy(false); }
  };
  async function reauthenticate(name: string): Promise<void> {
    const connection = connections.find(value => value.name === name);
    if (!connection) throw new Error("Connection is no longer available.");
    setBusy(true);
    try {
      await reauthenticateConnection(connection);
      setNotice({ kind: "info", message: "Signed in. You can run your query again." });
    } finally { setBusy(false); }
  }
  function updateConnections(value: OfficeConnection[]): void { saveConnections(value); setConnections(value); resetRuntime(); }
  function chooseDefault(name: string): void { setDefaultName(name); setDefaultConnectionName(name); resetRuntime(); setNotice(null); }
  function openWorkspace(next: Workspace): void { if (next === "connections") setSettingsSection("connections"); if (next !== "connections") setWorkspace(next); setView(next); }
  function queueQuery(sql: string, name?: string, navigate = false): void { setPendingQueries((values) => [...values, { id: crypto.randomUUID(), sql, name }]); if (navigate) openWorkspace("query"); }
  function openQuery(sql: string): void { queueQuery(sql, undefined, true); }
  const runtime = browserRuntimeDiagnostics();
  const diagnostics = `Cupola for Excel ${__APP_VERSION__}\nBuild ${__BUILD_ID__}\nTransport: HTTPS VGI only\nEngine: Haybarn WebAssembly (${runtime.selectedBundle ?? "not started"})\nCross-origin isolated: ${runtime.crossOriginIsolated}\nSharedArrayBuffer: ${runtime.sharedArrayBuffer}\nEngine assets: ${runtime.assetBase}\nConnection: ${active?.name ?? "none"}\nCatalog: ${active?.catalog ?? "none"}`;
  async function copy(value: string, _success?: string): Promise<void> { try { await navigator.clipboard.writeText(value); } catch (error) { setNotice({ kind: "error", message: message(error) }); } }

  return <main className="app-shell">
    <header><div className="brand"><img className="mark" src="./cupola-mark.svg" alt=""/><h1>Cupola <span>for Excel</span></h1></div>
    <WorkspaceTabs value={view} tabs={[{ id: "query", label: "Query Editor", icon: <FileCode2 aria-hidden="true"/> }, { id: "agent", label: "Ask AI", icon: <Sparkles aria-hidden="true"/> }, { id: "catalog", label: "Catalog View", icon: <Database aria-hidden="true"/> }]} onChange={openWorkspace}/><div className="header-actions"><button className="icon-button" title="Workbook tools" aria-label="Workbook tools" onClick={() => setView("workbook")}><TableProperties aria-hidden="true"/></button><button className="icon-button" aria-label="Settings" title="Settings" aria-pressed={view === "connections"} aria-controls="panel-connections" onClick={() => setView("connections")}><Settings aria-hidden="true"/></button></div></header>
    <AuthRecovery connections={connections} busy={busy} onSignIn={reauthenticate} onOpenConnections={() => { openWorkspace("connections"); requestAnimationFrame(() => document.querySelector<HTMLElement>("#panel-connections button")?.focus()); }}/>
    <Notice value={notice} onDismiss={() => setNotice(null)} onRetry={notice?.kind === "error" && retry.current ? retry.current : undefined} onDiagnostics={() => void copy(diagnostics, "Diagnostics copied.")}/>
    {!active ? (view !== "connections" && view !== "agent" ? <Onboarding onConnect={() => openWorkspace("connections")}/> : null) : <>
      <TabPanel id="query" active={view === "query"} busy={busy}><QueryPanel key={active.name} connection={active} pendingQueries={pendingQueries} onPendingConsumed={(ids) => setPendingQueries((values) => values.filter((value) => !ids.includes(value.id)))} busy={busy} action={action} copy={copy}/></TabPanel>

      <TabPanel id="catalog" active={view === "catalog"} busy={busy}><CatalogPanel active={view === "catalog"} connection={active} busy={busy} action={action} openQuery={openQuery} copy={copy}/></TabPanel>
    </>}
      <TabPanel id="agent" active={view === "agent"} busy={busy}><AgentPanel active={view === "agent"} settingsActive={view === "connections" && settingsSection === "ai"} settingsTarget={aiSettingsTarget} onConfigure={openAISettings} key={active?.name ?? "default"} connection={active} busy={busy} setBusy={setBusy} setNotice={setNotice} copy={copy} onCreateQuery={(value, navigate) => queueQuery(value.sql, value.name, navigate)}/></TabPanel>
    <div id="panel-connections" hidden={view !== "connections"} className="workspace-panel"><SettingsPage section={settingsSection} onSection={setSettingsSection} onBack={() => openWorkspace(workspace)} connections={<ConnectionsPanel onSignIn={reauthenticate} connections={connections} defaultName={defaultName} disabled={busy} onChange={updateConnections} onDefault={chooseDefault} action={action}/>} aiTarget={setAISettingsTarget} about={<><AboutContent diagnostics={diagnostics} onCopy={() => void copy(diagnostics, "Diagnostics copied.")}/><AIRequestDiagnostics/></>}/></div>
    {view === "workbook" && <SettingsFrame title="Workbook data" back={() => openWorkspace(workspace)}><WorkbookPanel connection={active} disabled={busy} action={action}/></SettingsFrame>}

    {about && <AboutDialog onClose={() => setAbout(false)} onCopy={() => void copy(diagnostics, "Diagnostics copied.")}/>}
  </main>;
}

function SettingsFrame({ title, back, children }: { title: string; back(): void; children: React.ReactNode }): React.JSX.Element {
  return <section className="settings-view" aria-labelledby="settings-title"><div className="section-heading"><div><p className="eyebrow">Settings</p><h2 id="settings-title">{title}</h2></div><button onClick={back}>Back</button></div>{children}</section>;
}

type OfficeQueryRuntime = { result: QueryResult | null; previewLimit: number; previewOffset: number; outcome: WorkbookWriteOutcome | null; executedSql?: string };
const EMPTY_OFFICE_QUERY_RUNTIME: OfficeQueryRuntime = { result: null, previewLimit: 200, previewOffset: 0, outcome: null };

function QueryPanel({ connection, pendingQueries, onPendingConsumed, busy, action, copy }: { connection?: OfficeConnection; pendingQueries: PendingQuery[]; onPendingConsumed(ids: string[]): void; busy: boolean; action: Action; copy(value: string, success?: string): Promise<void> }): React.JSX.Element {
  const scope = connection?.name ?? "default";
  const [documents, setDocuments] = useState<QueryDocumentState>(() => loadQueryDocumentState(scope));
  const [runtimes, setRuntimes] = useState<Record<string, OfficeQueryRuntime>>({});
  const [runningId, setRunningId] = useState<string | null>(null);
  const [queryStatus, setQueryStatus] = useState("");
  const [resultsWindowError, setResultsWindowError] = useState("");
  const [cancelling, setCancelling] = useState(false);
  const activeRun = useRef<AbortController | null>(null);
  const [resultsCollapsed, setResultsCollapsed] = useState(() => localStorage.getItem("cupola.query.resultsCollapsed") === "true");
  useEffect(() => { localStorage.setItem("cupola.query.resultsCollapsed", String(resultsCollapsed)); }, [resultsCollapsed]);
  const [editorPercent, setEditorPercent] = useState(() => { const value = Number(localStorage.getItem("cupola.query.editorPercent")); return Number.isFinite(value) && value >= 22 && value <= 72 ? value : 32; });
  const activeDocument = documents.documents.find((value) => value.id === documents.activeId) ?? documents.documents[0];
  const runtime = runtimes[activeDocument.id] ?? EMPTY_OFFICE_QUERY_RUNTIME;
  const { result, previewLimit, previewOffset, outcome } = runtime;
  const sql = activeDocument.sql;
  function updateDocuments(change: (value: QueryDocumentState) => QueryDocumentState): void { setDocuments((previous) => { const next = change(previous); saveQueryDocumentState(scope, next); return next; }); }
  function setSql(value: string): void { updateDocuments((state) => updateQueryDocumentSql(state, state.activeId, value)); }
  function patchRuntime(id: string, value: Partial<OfficeQueryRuntime>): void { setRuntimes((previous) => ({ ...previous, [id]: { ...EMPTY_OFFICE_QUERY_RUNTIME, ...previous[id], ...value } })); }
  function closeDocument(id: string): void { updateDocuments((state) => removeQueryDocument(state, id)); setRuntimes((previous) => { const { [id]: _removed, ...remaining } = previous; return remaining; }); }
  useEffect(() => { saveQueryDocumentState(scope, documents); }, []);
  useEffect(() => { if (!pendingQueries.length) return; updateDocuments((state) => pendingQueries.reduce((value, pending) => addQueryDocument(value, pending.sql, pending.name), state)); onPendingConsumed(pendingQueries.map((value) => value.id)); }, [pendingQueries]);
  async function run(): Promise<void> {
    if (!connection || busy || activeRun.current) return;
    const documentId = activeDocument.id, statement = sql;
    const request = new AbortController();
    activeRun.current = request;
    setQueryStatus("Running query…"); setCancelling(false);
    setRunningId(documentId);
    try {
      const value = await action("Running query…", async () => {
        try { return await (await resolveBackend(connection.name, true)).query(statement, { maxRows: 10_000, signal: request.signal }); }
        catch (error) {
          if (error instanceof DOMException && error.name === "AbortError") { setQueryStatus("Query cancelled."); return null; }
          setQueryStatus(""); throw error;
        }
      }, null); if (!value) return;
      setQueryStatus(value.rows.length ? "Query completed." : "Query completed. No rows returned.");
      patchRuntime(documentId, { result: value, outcome: null, previewLimit: 200, previewOffset: 0, executedSql: statement });
    } finally { activeRun.current = null; setCancelling(false); setRunningId((value) => value === documentId ? null : value); }
  }
  async function cancelQuery(): Promise<void> {
    const request = activeRun.current;
    if (!request || cancelling) return;
    setCancelling(true); setQueryStatus("Cancelling…");
    request.abort();
  }
  useEffect(() => () => {
    activeRun.current?.abort();
  }, []);
  async function openResults(): Promise<void> {
    if (!result) return;
    setResultsWindowError("");
    try { await openResultsWindow({ title: activeDocument.name, result }); }
    catch (error) { setResultsWindowError(error instanceof Error ? error.message : "Could not open the results window. Try again."); }
  }
  async function insert(): Promise<void> { if (!result || !connection) return; const documentId = activeDocument.id, statement = runtime.executedSql ?? sql; const needsFullResult = result.truncated || result.rows.length < result.rowCount; const value = await action(needsFullResult ? "Loading full result and inserting snapshot…" : "Inserting snapshot…", async () => { const complete = needsFullResult ? await (await resolveBackend(connection.name)).query(statement, { maxRows: EXCEL_MAX_DATA_ROWS + 1 }) : result; return insertResult(complete, "VGI_Result"); }, "Snapshot inserted."); if (value) patchRuntime(documentId, { outcome: value }); }
  useEffect(() => { localStorage.setItem("cupola.query.editorPercent", String(editorPercent)); }, [editorPercent]);
  const hasResults = !!result?.rows.length;
  const editorExpanded = resultsCollapsed || !hasResults;
  const previewStart = result?.rows.length ? previewOffset + 1 : 0;
  const previewEnd = result ? Math.min(previewOffset + previewLimit, result.rows.length) : 0;
  return <section className={`query-editor${editorExpanded ? " results-collapsed" : ""}`}>
    <QueryTabs documents={documents.documents} activeId={documents.activeId} runningId={runningId} onSelect={(id) => updateDocuments((state) => activateQueryDocument(state, id))} onAdd={() => updateDocuments((state) => addQueryDocument(state))} onClose={closeDocument} onRename={(id, name) => updateDocuments((state) => renameQueryDocument(state, id, name))}/>
    <div className="query-toolbar" role="toolbar" aria-label="Query editor actions">{runningId ? <span className="query-run-progress"><button className="primary run-query" aria-label="Cancel query" title="Cancel query" disabled={cancelling} onClick={() => void cancelQuery()}><LoaderCircle className="busy-spinner" aria-hidden="true"/><span>Cancel</span></button><span className="query-running-status" aria-hidden="true">{cancelling ? "Cancelling…" : "Running query…"}</span></span> : <button className="primary run-query" disabled={busy || !connection || !sql.trim()} onClick={() => void run()} title="Run query (Ctrl+Enter)"><Play aria-hidden="true"/><span>{busy ? "Running…" : "Run"}</span></button>}<span className="toolbar-divider" aria-hidden="true"/><button className="toolbar-button" onClick={() => setSql(formatSql(sql))} title="Format SQL"><WandSparkles aria-hidden="true"/><span>Format</span></button><button className="toolbar-button" onClick={() => void copy(sql)} title="Copy SQL"><Copy aria-hidden="true"/><span>Copy</span></button></div>
    <div className="editor-surface" style={{ flexBasis: editorExpanded ? "0px" : `${editorPercent}%` }}><label className="sr-only" htmlFor="sql-editor">SQL query</label><textarea id="sql-editor" value={sql} onChange={(event) => setSql(event.target.value)} onKeyDown={(event) => { if ((event.ctrlKey || event.metaKey) && event.key === "Enter") { event.preventDefault(); void run(); } }} spellCheck={false}/></div>
    <div className="editor-statusbar"><span>{connection ? connectionLabel(connection.name, connection.catalog) : "No connection"}</span><span role="status" aria-live="polite">{queryStatus || "Ctrl+Enter to run"}</span></div>
    {hasResults && !resultsCollapsed && <QuerySplitter value={editorPercent} onChange={setEditorPercent}/>}
    {hasResults && result && <div className="query-results-pane">
      <div className="results-toolbar"><div><Table2 aria-hidden="true"/><strong>Results</strong><small>{result.rowCount.toLocaleString()} rows{!resultsCollapsed && result.truncated && ` · ${result.rows.length.toLocaleString()} loaded`}</small></div><button className="preview-toggle" aria-label={resultsCollapsed ? "Show results" : "Hide results"} title={resultsCollapsed ? "Show results" : "Hide results"} aria-expanded={!resultsCollapsed} aria-controls="query-results-preview" onClick={() => setResultsCollapsed(value => !value)}>{resultsCollapsed ? <ChevronUp aria-hidden="true"/> : <ChevronDown aria-hidden="true"/>}{resultsCollapsed && <span>Show results</span>}</button></div>
      {!resultsCollapsed && result.rows.length > QUERY_PREVIEW_OPTIONS[0] && <div className="results-paging"><small>{previewStart.toLocaleString()}–{previewEnd.toLocaleString()} of {result.rows.length.toLocaleString()} loaded</small>{result.rows.length > previewLimit && <span className="preview-pager"><button aria-label="Previous preview page" title="Previous preview page" disabled={previewOffset === 0} onClick={() => patchRuntime(activeDocument.id, { previewOffset: Math.max(0, previewOffset - previewLimit) })}><ChevronLeft aria-hidden="true"/></button><button aria-label="Next preview page" title="Next preview page" disabled={previewOffset + previewLimit >= result.rows.length} onClick={() => patchRuntime(activeDocument.id, { previewOffset: Math.min(Math.max(0, result.rows.length - 1), previewOffset + previewLimit) })}><ChevronRight aria-hidden="true"/></button></span>}{result.rows.length > QUERY_PREVIEW_OPTIONS[0] && <label className="preview-size">Rows<select aria-label="Rows shown per result page" value={previewLimit} onChange={(event) => patchRuntime(activeDocument.id, { previewLimit: Number(event.target.value), previewOffset: 0 })}>{QUERY_PREVIEW_OPTIONS.map((value) => <option key={value} value={value}>{value.toLocaleString()}</option>)}</select></label>}</div>}
      <div id="query-results-preview" className="query-results-content" hidden={resultsCollapsed}>{resultsWindowError && <p className="field-error" role="alert">{resultsWindowError} <button onClick={() => void openResults()}>Try again</button></p>}{outcome && <WorkbookOutcome value={outcome}/>}<ResultGrid result={result} label="Query results" limit={previewLimit} offset={previewOffset}/></div>
      {!resultsCollapsed && <div className="results-actions" role="toolbar" aria-label="Result actions"><button aria-label="Open results in new window" title="Open results in new window" onClick={() => void openResults()}><Maximize2 aria-hidden="true"/></button><ResultsMore><button aria-label={result.truncated ? "Copy loaded preview" : "Copy results"} title={result.truncated ? "Copy the rows loaded for preview" : "Copy results"} onClick={() => void copy(resultTsv(result))}><Copy aria-hidden="true"/><span>{result.truncated ? "Copy loaded preview" : "Copy results"}</span></button></ResultsMore><button className="primary" title="Insert a static Excel table with no refresh connection" disabled={busy} onClick={() => void insert()}><TableProperties aria-hidden="true"/><span>Insert into Excel</span></button></div>}
    </div>}
  </section>;
}

function CatalogLoading(): React.JSX.Element {
  return <div className="catalog-loading" role="status" aria-live="polite"><LoaderCircle className="busy-spinner" aria-hidden="true"/><span>Loading catalog…</span></div>;
}

function CatalogPanel({ active, connection, busy, action, openQuery, copy }: { active: boolean; connection?: OfficeConnection; busy: boolean; action: Action; openQuery(sql: string): void; copy(value: string, success?: string): Promise<void> }): React.JSX.Element {
  const [loading, setLoading] = useState(false);
  const [objects, setObjects] = useState<CatalogObject[]>([]), [filter, setFilter] = useState(""), [selected, setSelected] = useState<CatalogObject | null>(null), [columns, setColumns] = useState<QueryResult | null>(null), [functions, setFunctions] = useState<CatalogFunction[]>([]);
  async function load(): Promise<void> { if (!connection) return; await action("Loading catalog…", async () => { setLoading(true); try { const backend = await resolveBackend(connection.name); const [tableResult, functionValues] = await Promise.all([backend.query("SELECT table_catalog AS catalog, table_schema AS schema, table_name AS name, CASE WHEN table_type='VIEW' THEN 'view' ELSE 'table' END AS kind FROM information_schema.tables WHERE table_schema NOT IN ('information_schema','pg_catalog') ORDER BY 1,2,3"), discoverFunctions(backend)]); setFunctions(functionValues); setObjects([...rows(tableResult).map((row) => ({ catalog: String(row.catalog), schema: String(row.schema), name: String(row.name), kind: String(row.kind), row })), ...functionValues.map((fn) => ({ catalog: fn.catalog, schema: fn.schema, name: fn.name, kind: fn.kind === "table" ? "table function" : "function", description: fn.description }))]); } finally { setLoading(false); } }, "Catalog loaded."); }
  useEffect(() => { setObjects([]); setSelected(null); setColumns(null); if (active && connection) void load(); }, [active, connection?.name]);
  async function select(value: CatalogObject): Promise<void> { setSelected(value); setColumns(null); if (!connection || value.kind.includes("function")) return; const backend = await resolveBackend(connection.name); const result = await action("Loading columns…", () => backend.query(`SELECT column_name, data_type, is_nullable FROM information_schema.columns WHERE table_catalog=${literal(value.catalog)} AND table_schema=${literal(value.schema)} AND table_name=${literal(value.name)} ORDER BY ordinal_position`), "Columns loaded."); if (result) setColumns(result); }
  const visible = objects.filter((value) => !filter.trim() || `${value.schema} ${value.name} ${value.kind} ${value.description ?? ""}`.toLowerCase().includes(filter.toLowerCase()));
  const fn = selected ? functions.find((value) => value.catalog === selected.catalog && value.schema === selected.schema && value.name === selected.name) : undefined;
  return <section className="catalog-panel"><div className="section-heading"><div><h2>Catalog View</h2><p>Browse schemas, tables, functions, arguments, and return types.</p></div></div><label className="catalog-search" htmlFor="catalog-filter">Search catalog<input id="catalog-filter" value={filter} onChange={(event) => setFilter(event.target.value)} placeholder="Function, table, description…"/></label><div className="catalog-browser"><CatalogTree loading={loading} onRefresh={() => void load()} refreshDisabled={busy || !connection} objects={visible} catalog={connection?.catalog ?? connection?.name ?? "Catalog"} forceExpanded={!!filter.trim()} onSelect={(value) => void select(value)}/><div className="catalog-inspector">{selected ? <div className="catalog-detail"><div className="detail-heading"><h3>{selected.catalog}.{selected.schema}.{selected.name}</h3><div className="compact-actions"><button onClick={() => void copy(`${selected.catalog}.${selected.schema}.${selected.name}`, "Name copied.")}>Copy name</button><button className="primary" onClick={() => openQuery(fn ? functionQuery(fn) : `SELECT *\nFROM ${qualified(selected)}\nLIMIT 100;`)}>Insert into query</button></div></div><div className="catalog-badges"><span>{selected.kind}</span>{fn?.returnType && <span>returns {fn.returnType}</span>}</div>{selected.description && <p>{selected.description}</p>}{fn && <><h4>Signature</h4><pre className="catalog-signature"><code>{functionSignature(fn)}</code></pre><h4>Arguments</h4><ResultGrid result={functionArguments(fn)} label={`Arguments for ${fn.name}`}/></>}{columns && <><h4>Columns</h4><ResultGrid result={columns} label={`Columns for ${selected.name}`}/></>}</div> : loading ? null : <div className="empty-detail"><h3>Select a catalog object</h3><p>Choose an object to inspect it and place a starter statement in Query.</p></div>}</div></div></section>;
}

function OfficeTreeBranch({ level, label, expanded, onToggle, children }: { level: number; label: string; expanded: boolean; onToggle(): void; children: React.ReactNode }): React.JSX.Element { return <li role="none" className="tree-node"><button type="button" role="treeitem" aria-level={level} aria-expanded={expanded} className="tree-branch" onClick={onToggle} onKeyDown={(event) => { if (event.key === "ArrowRight" && !expanded) { event.preventDefault(); onToggle(); } else if (event.key === "ArrowLeft" && expanded) { event.preventDefault(); onToggle(); } }}><span className="tree-chevron" aria-hidden="true">{expanded ? <ChevronDown/> : <ChevronRight/>}</span><span className="tree-folder" aria-hidden="true">{expanded ? <FolderOpen/> : <Folder/>}</span><span className="tree-label">{label}</span></button>{expanded && <ul role="group" aria-label={label}>{children}</ul>}</li>; }
function CatalogTree({ loading = false, onRefresh, refreshDisabled = false, objects, catalog, forceExpanded, onSelect }: { loading?: boolean; onRefresh?(): void; refreshDisabled?: boolean; objects: CatalogObject[]; catalog: string; forceExpanded: boolean; onSelect(value: CatalogObject): void }): React.JSX.Element {
  const [showHidden, setShowHidden] = useState(false);
  const model = useMemo(() => officeTreeModel(objects, showHidden), [objects, showHidden]), completeModel = useMemo(() => officeTreeModel(objects, true), [objects]), hiddenCount = useMemo(() => objects.filter((value) => value.name.includes("$")).length, [objects]), rootId = `catalog:${catalog}`, allIds = useMemo(() => [rootId, ...model.flatMap((schema) => [schema.id, ...schema.folders.map((folder) => `${schema.id}:${folder.id}`)])], [rootId, model]), completeIds = useMemo(() => [rootId, ...completeModel.flatMap((schema) => [schema.id, ...schema.folders.map((folder) => `${schema.id}:${folder.id}`)])], [rootId, completeModel]);
  const [expanded, setExpanded] = useState<Set<string>>(new Set()), [selected, setSelected] = useState<string | null>(null);
  useEffect(() => { const initial = new Set([rootId]); if (model[0]) { initial.add(model[0].id); for (const folder of model[0].folders) initial.add(`${model[0].id}:${folder.id}`); } setExpanded(initial); setSelected(null); setShowHidden(false); }, [rootId, objects.length === 0]);
  const shown = forceExpanded ? new Set(allIds) : expanded, toggle = (id: string) => setExpanded((current) => { const next = new Set(current); if (next.has(id)) next.delete(id); else next.add(id); return next; });
  const keyboard = (event: React.KeyboardEvent<HTMLDivElement>) => { if (!["ArrowDown", "ArrowUp", "Home", "End"].includes(event.key)) return; const items = [...event.currentTarget.querySelectorAll<HTMLElement>("[role='treeitem']")], current = items.indexOf(document.activeElement as HTMLElement), index = event.key === "Home" ? 0 : event.key === "End" ? items.length - 1 : Math.max(0, Math.min(items.length - 1, current + (event.key === "ArrowDown" ? 1 : -1))); if (items[index]) { event.preventDefault(); items[index].focus(); } };
  return <aside className="catalog-tree" aria-label="Catalog object sidebar"><div className="catalog-tree-toolbar"><div className="tree-toolbar-primary"><strong>Object tree</strong>{hiddenCount > 0 && <button type="button" className="tree-hidden-toggle" aria-pressed={showHidden} title={showHidden ? "Hide catalog items whose names contain $" : `Show ${hiddenCount} catalog items whose names contain $`} onClick={() => { if (showHidden) setShowHidden(false); else { setShowHidden(true); setExpanded(new Set(completeIds)); } }}>{showHidden ? <EyeOff aria-hidden="true"/> : <Eye aria-hidden="true"/>}<span>{showHidden ? "Hide hidden" : `Show hidden (${hiddenCount})`}</span></button>}{onRefresh && <button type="button" className="tree-refresh" title="Refresh catalog" disabled={refreshDisabled} onClick={onRefresh}><RefreshCw aria-hidden="true"/><span>Refresh</span></button>}</div></div>{loading ? <CatalogLoading/> : <div className="catalog-tree-scroll" role="tree" aria-label={`${catalog} catalog objects`} onKeyDown={keyboard}><ul className="tree-root" role="none"><OfficeTreeBranch level={1} label={catalog} expanded={shown.has(rootId)} onToggle={() => toggle(rootId)}>{model.map((schema) => <OfficeTreeBranch key={schema.id} level={2} label={schema.name} expanded={shown.has(schema.id)} onToggle={() => toggle(schema.id)}>{schema.folders.map((folder) => { const folderId = `${schema.id}:${folder.id}`; return <OfficeTreeBranch key={folderId} level={3} label={folder.label} expanded={shown.has(folderId)} onToggle={() => toggle(folderId)}>{folder.values.map((value) => { const key = `${value.catalog}.${value.schema}.${value.name}.${value.kind}`; return <li role="none" key={key}><button type="button" role="treeitem" aria-level={4} aria-selected={selected === key} aria-label={`${value.name}, ${value.kind}`} className="tree-leaf" title={`${value.schema}.${value.name}`} onClick={() => { setSelected(key); onSelect(value); }}><span className="tree-spacer"/><span className="object-icon" aria-hidden="true">{value.kind === "table" || value.kind === "view" ? <Table2/> : <Braces/>}</span><span className="tree-label">{value.name}</span></button></li>; })}</OfficeTreeBranch>; })}</OfficeTreeBranch>)}</OfficeTreeBranch></ul>{model.length === 0 && <p className="tree-empty">No matching catalog objects.</p>}</div>}</aside>;
}

function AgentPanel({ active, settingsActive, settingsTarget, onConfigure, connection, busy, setBusy, setNotice, copy, onCreateQuery }: { active: boolean; settingsActive: boolean; settingsTarget: HTMLDivElement | null; onConfigure(): void; connection?: OfficeConnection; busy: boolean; setBusy(value: boolean): void; setNotice(value: NoticeValue): void; copy(value: string, success?: string): Promise<void>; onCreateQuery(value: { name: string; sql: string }, navigate?: boolean): void }): React.JSX.Element {
  const clarification = useClarification();
  const activeAgent = useRef<AbortController | null>(null);
  const [stopping, setStopping] = useState(false);
  useEffect(() => () => { activeAgent.current?.abort(); }, []);
  function stop(): void { if (activeAgent.current && !activeAgent.current.signal.aborted) { setStopping(true); activeAgent.current.abort(); } }
  useEffect(() => {
    if (!active) return;
    const onEscape = (event: KeyboardEvent) => {
      if (event.key !== "Escape" || event.defaultPrevented || event.isComposing || event.repeat || !activeAgent.current || activeAgent.current.signal.aborted) return;
      event.preventDefault(); stop();
    };
    window.addEventListener("keydown", onEscape);
    return () => window.removeEventListener("keydown", onEscape);
  }, [active]);
  const scope = connection?.name ?? "default";
  const [workspaceId, setWorkspaceId] = useState(() => sessionStorage.getItem("cupola.anthropic-workspace-id") ?? "");
  const [apiKey, setApiKey] = useState(() => sessionStorage.getItem("vgi.excel.anthropic-key") ?? "");
  const [conversations, setConversations] = useState<OfficeAgentConversationState>(() => loadOfficeAgentConversationState(scope, DEFAULT_AI_MODEL));
  const [controller, setController] = useState<AbortController | null>(null), [runningId, setRunningId] = useState<string | null>(null);
  const sessions = useRef(new Map<string, OfficeAgentSession>()), scroll = useRef<HTMLDivElement>(null);
  const activeConversation = conversations.documents.find((value) => value.id === conversations.activeId) ?? conversations.documents[0];
  const { model, draft: prompt, displayMessages: messages, staged = null, outcome = null } = activeConversation;
  const modelCatalog = useModelCatalog(apiKey, workspaceId, settingsActive);
  const modelInfo = modelCatalog.models.find(value => value.id === model);
  const effortLevels = modelEfforts(model, modelInfo);
  useEffect(() => { saveOfficeAgentConversationState(scope, conversations); }, [scope, conversations]);
  useEffect(() => { scroll.current?.scrollTo({ top: scroll.current.scrollHeight, behavior: "smooth" }); }, [messages]);
  function updateConversation(id: string, change: (value: OfficeAgentConversation) => OfficeAgentConversation): void { setConversations((state) => ({ ...state, documents: state.documents.map((value) => value.id === id ? { ...change(value), updatedAt: Date.now() } : value) })); }
  function updateLast(id: string, change: (value: OfficeChatMessage) => OfficeChatMessage): void { updateConversation(id, (value) => ({ ...value, displayMessages: value.displayMessages.map((message, index) => index === value.displayMessages.length - 1 ? change(message) : message) })); }
  function sessionFor(value: OfficeAgentConversation): OfficeAgentSession { let session = sessions.current.get(value.id); if (!session) { session = new OfficeAgentSession(); session.restore(value.agentMessages); sessions.current.set(value.id, session); } return session; }
  async function ask(): Promise<void> {
    if (clarification.question && runningId === activeConversation.id && prompt.trim()) { clarification.answer(prompt); updateConversation(activeConversation.id, value => ({ ...value, draft: "" })); return; }
    if (activeAgent.current || !apiKey || !prompt.trim() || !connection) return;
    sessionStorage.setItem("vgi.excel.anthropic-key", apiKey);
    const conversationId = activeConversation.id, question = prompt.trim(), session = sessionFor(activeConversation);
    updateConversation(conversationId, (value) => ({ ...value,
      name: /^Conversation \d+$/.test(value.name) ? officeConversationTitle(question) : value.name,
      draft: "", outcome: null,
      displayMessages: [...value.displayMessages, { role: "user", text: question }, { role: "assistant", modelId: model, modelName: modelInfo?.name || model, text: "", timeline: [], tools: [], streaming: true, activity: "Thinking…" }],
    }));
    const abort = new AbortController(); activeAgent.current = abort; setStopping(false);
    setController(abort); setRunningId(conversationId); setBusy(true); setNotice(null);
    try {
      const backend = await resolveBackend(connection.name, "agent");
      const answer = await session.run(apiKey, question, agentContext(backend, onCreateQuery), abort.signal, { model, modelInfo, workspaceId, effort: normalizeEffort(activeConversation.effort), maxTokens: activeConversation.maxTokens, connection: { name: connection.name, catalog: connection.catalog || connection.name, authentication: connection.authentication ?? "anonymous" } }, {
        onStage: activity => updateLast(conversationId, (value) => ({ ...value, activity })),
        onClarification: async (question, signal) => {
          updateLast(conversationId, (value) => ({ ...value, clarification: question, activity: "Waiting for your answer…" }));
          const answer = await clarification.request(question, signal);
          updateLast(conversationId, (value) => ({ ...value, clarification: { ...question, answer }, activity: "Thinking…" }));
          return answer;
        },
        onQueryResult: result => updateLast(conversationId, (value) => ({ ...value, results: [...(value.results ?? []), { ...result, catalog: connection.catalog ?? connection.name }] })),
        onThinking: text => updateLast(conversationId, (value) => ({ ...value, timeline: appendTranscript(value.timeline, { type: "thinking", text }), activity: "Thinking…" })),
        onText: (chunk) => updateLast(conversationId, (value) => ({ ...value, text: value.text + chunk, timeline: appendTranscript(value.timeline, { type: "text", text: chunk }), activity: undefined })),
        onTool: (name, state, detail, id) => updateLast(conversationId, (value) => { const tools = [...(value.tools ?? [])], key = id ?? crypto.randomUUID(), index = tools.findIndex((tool) => tool.id === key); if (index >= 0) tools[index] = { ...tools[index], state, detail, sql: tools[index].sql ?? (name === "run_sql" && state === "running" ? detail : undefined) }; else tools.push({ id: key, name, state, detail, sql: name === "run_sql" && state === "running" ? detail : undefined }); return { ...value, tools, timeline: appendTranscript(value.timeline, { type: "tool", id: key }), activity: state === "done" ? "Preparing answer…" : toolStage(name) }; }),
        onRetry: (activity) => updateLast(conversationId, value => ({ ...value, activity: activity ?? "Thinking…" })),
        onResult: (value) => updateConversation(conversationId, (conversation) => ({ ...conversation, staged: value })),
      });
      if (answer.stagedResult) updateConversation(conversationId, (value) => ({ ...value, staged: answer.stagedResult }));
    } catch (error) {
      if ((error as Error).name === "AbortError") updateLast(conversationId, (value) => ({ ...value, stopped: true, activity: undefined, tools: value.tools?.map(tool => tool.state === "running" || tool.state === "writing" ? { ...tool, state: "stopped" as const } : tool) }));
      else { captureError(error, "agent.run"); updateLast(conversationId, (value) => ({ ...value, text: `${value.text}${value.text ? "\n\n" : ""}Could not finish: ${message(error)} You can refine your question and try again.` })); }
    } finally {
      updateConversation(conversationId, (value) => ({ ...value, agentMessages: session.snapshot(), displayMessages: value.displayMessages.map((message, index) => index === value.displayMessages.length - 1 ? { ...message, streaming: false, activity: undefined } : message) }));
      activeAgent.current = null; setStopping(false);
      setController(null); setRunningId(null); setBusy(false);
    }
  }
  async function insert(): Promise<void> { if (!staged || !connection) return; const value = await insertResult(staged, "VGI_Agent_Result"); if (value) updateConversation(activeConversation.id, (conversation) => ({ ...conversation, outcome: value })); }
  function clearCurrent(): void { sessionFor(activeConversation).reset(); updateConversation(activeConversation.id, (value) => ({ ...value, displayMessages: [], agentMessages: [], staged: null, outcome: null })); }
  function closeConversation(id: string): void { if (id === runningId) return; sessions.current.delete(id); setConversations((state) => removeOfficeAgentConversation(state, id, model)); }
  return <section className="agent"><QueryTabs documents={conversations.documents} activeId={activeConversation.id} runningId={runningId} label="AI conversations" itemName="conversation" onSelect={(id) => setConversations((state) => activateOfficeAgentConversation(state, id))} onAdd={() => setConversations((state) => addOfficeAgentConversation(state, model))} onClose={closeConversation} onRename={(id, name) => setConversations((state) => renameOfficeAgentConversation(state, id, name))}/><div className="agent-toolbar"><div><span className={`health-dot ${apiKey ? "configured" : ""}`}/><span>{apiKey ? "Conversations saved locally" : "AI setup required"}</span></div><div className="compact-actions"><button disabled={!messages.length || !!controller} onClick={clearCurrent}>Clear current</button></div></div>{settingsTarget && createPortal(<div className="agent-settings"><label>Anthropic API key<input type="password" value={apiKey} onChange={(event) => { setApiKey(event.target.value); sessionStorage.setItem("vgi.excel.anthropic-key", event.target.value); }} autoComplete="off"/></label><p className="ai-conversation-scope">Model and response options for <strong>{activeConversation.name}</strong>.</p><ModelCatalog key={activeConversation.id} catalog={modelCatalog} disabled={!apiKey.trim()} selectionDisabled={!!controller} value={model} onChange={model => updateConversation(activeConversation.id, value => ({ ...value, model }))}/>{supportsEffort(model, modelInfo) && <label>Thinking effort<select value={normalizeEffort(activeConversation.effort)} disabled={!!controller} onChange={(event) => updateConversation(activeConversation.id, value => ({ ...value, effort: normalizeEffort(event.target.value) }))}>{!effortLevels.includes(normalizeEffort(activeConversation.effort)) && <option value={normalizeEffort(activeConversation.effort)} disabled>{normalizeEffort(activeConversation.effort)} (unavailable)</option>}{effortLevels.map(effort => <option key={effort} value={effort}>{effort}</option>)}</select></label>}<p className="hint">Your API key is kept only for this Excel session.</p><button onClick={() => { sessionStorage.removeItem("vgi.excel.anthropic-key"); setApiKey(""); }}>Forget key</button>{modelInfo && ((activeConversation.maxTokens ?? 16384) > clampMaxTokens(model, activeConversation.maxTokens, modelInfo) || (effortLevels.length > 0 && !effortLevels.includes(normalizeEffort(activeConversation.effort)))) && <p className="hint">This model uses a supported thinking level and limits output to {clampMaxTokens(model, activeConversation.maxTokens, modelInfo).toLocaleString()} tokens. Your saved options are unchanged.</p>}<details className="ai-advanced"><summary>Advanced options</summary><label>Anthropic workspace ID (optional)<input value={workspaceId} disabled={!!controller} onChange={(event) => { setWorkspaceId(event.target.value); sessionStorage.setItem("cupola.anthropic-workspace-id", event.target.value); }} placeholder="wrkspc_…" autoComplete="off"/></label><label>Max output tokens<input type="number" min="1" max={modelInfo?.maxTokens ?? 128000} value={activeConversation.maxTokens ?? 16384} disabled={!!controller} onChange={(event) => updateConversation(activeConversation.id, value => ({ ...value, maxTokens: Number(event.target.value) }))}/></label></details></div>, settingsTarget)}<div className="chat" ref={scroll}>{!messages.length && <div className="empty"><img className="empty-mark" src="./cupola-mark.svg" alt=""/><h2>Ask AI about your data</h2><p>This conversation is saved locally for the selected VGI connection. Cupola explores the catalog and runs read-only SQL; workbook insertion requires confirmation.</p>{!apiKey && <button onClick={onConfigure}>Configure AI</button>}</div>}{messages.map((value, index) => <article className={`message ${value.role}`} key={index}><strong>{value.role === "user" ? "You" : value.modelName || value.modelId || "AI assistant"}</strong> {value.role === "assistant" ? <AgentTranscript message={value} renderTool={tool => <ToolCall value={tool} onOpenQuery={sql => onCreateQuery({ name: "AI query", sql }, true)}/>} renderResult={result => <AgentResult key={result.id} value={result} disabled={busy || !!controller} openQuery={sql => onCreateQuery({ name: "AI query", sql }, true)} load={async result => { setBusy(true); try { const outcome = await insertResult(result.result, "VGI_Agent_Result"); return outcome ? "Snapshot inserted." : "Insertion cancelled."; } finally { setBusy(false); } }} loadLabel="Insert into Excel"/>} onOpenQuery={sql => onCreateQuery({ name: "AI query", sql }, true)}/> : <div className="message-text">{value.text}</div>}{value.clarification && <Clarification value={value.clarification} enabled={!!clarification.question && !!value.streaming && runningId === activeConversation.id} choose={draft => updateConversation(activeConversation.id, value => ({ ...value, draft }))}/>}{value.stopped && <div className="agent-stopped" role="status">Stopped</div>}</article>)}</div><div className="composer">{controller && runningId === activeConversation.id && <div className="agent-run-status"><AgentProgress waiting={!!clarification.question} stage={stopping ? "Stopping…" : clarification.question ? "Waiting for your answer…" : messages[messages.length - 1]?.activity ?? (messages[messages.length - 1]?.text ? "Writing answer…" : "Thinking…")}/><button className="danger" title="Stop (Esc)" aria-keyshortcuts="Escape" disabled={stopping} onClick={stop}>{stopping ? "Stopping…" : "Stop"}</button></div>}<label className="sr-only" htmlFor="agent-prompt">Ask AI</label><div className="composer-input-row"><textarea id="agent-prompt" rows={2} value={prompt} onChange={(event) => updateConversation(activeConversation.id, (value) => ({ ...value, draft: event.target.value }))} onKeyDown={(event) => { if (event.key === "Enter" && !event.shiftKey && !event.nativeEvent.isComposing) { event.preventDefault(); void ask(); } }} placeholder="Ask a question about the active catalog…"/><button className="primary" disabled={(busy && !(clarification.question && runningId === activeConversation.id)) || !connection || !apiKey || !prompt.trim()} onClick={() => void ask()}>{clarification.question && runningId === activeConversation.id ? "Continue" : controller && runningId === activeConversation.id ? "Working…" : "Send"}</button></div>{controller && runningId !== activeConversation.id && runningId && <button onClick={() => setConversations((state) => activateOfficeAgentConversation(state, runningId))}>View running</button>}{outcome && <WorkbookOutcome value={outcome}/>}</div></section>;
}

function WorkbookPanel({ connection, disabled, action }: { connection?: OfficeConnection; disabled: boolean; action: Action }): React.JSX.Element {
  const [tables, setTables] = useState<Array<{ name: string; worksheet: string; rows: number }>>([]), [snapshots, setSnapshots] = useState<ManagedSnapshot[]>([]);
  async function load(): Promise<void> { const [tableValues, snapshotValues] = await Promise.all([listWorkbookTables(), listManagedSnapshots()]); setTables(tableValues); setSnapshots(snapshotValues); }
  useEffect(() => { void load().catch(() => undefined); }, []);
  return <div className="workbook-panel"><h3>Cupola tables</h3><p>Refresh these tables here in Cupola. They do not participate in Excel Refresh All. New snapshots are static Excel tables.</p><ul className="card-list">{snapshots.map((snapshot) => <li key={snapshot.table}><div><strong>{snapshot.table}</strong><small>{snapshot.connection} · updated {new Date(snapshot.updatedAt).toLocaleString()}</small></div><div className="row-actions"><button disabled={disabled} onClick={async () => { if (!await confirmAction(`Refresh table “${snapshot.table}” and replace its contents?`, "Refresh table", "Refresh table")) return; void action("Refreshing table…", async () => { await refreshSnapshot(snapshot); await load(); }, `${snapshot.table} refreshed.`); }}>Refresh table</button><button onClick={() => void goToTable(snapshot.table)}>Go to table</button><button disabled={disabled} onClick={async () => { if (!await confirmAction(`Keep “${snapshot.table}” as a static table? Its data will remain, but Cupola will no longer refresh it.`, "Keep as static table", "Keep as static table")) return; void action("Removing refresh metadata…", async () => { await forgetSnapshot(snapshot.table); await load(); }, "The table is now static."); }}>Keep as static table</button></div></li>)}{!snapshots.length && <li>No Cupola tables.</li>}</ul><h3>Make workbook data available to Cupola queries</h3><p>Imported workbook data appears as an in-memory snapshot in the <code>excel</code> schema for the current session.</p><div className="actions"><button disabled={disabled || !connection} onClick={() => void action("Importing selection…", importSelection, "Selection is available as excel.selection.")}>Make current selection available</button></div><ul className="card-list">{tables.map((table) => <li key={table.name}><div><strong>{table.name}</strong><small>{table.worksheet} · {table.rows} rows</small></div><button disabled={disabled || !connection} onClick={() => void action(`Importing ${table.name}…`, () => importWorkbookTable(table.name), `${table.name} is available in the excel schema.`)}>Make available</button></li>)}</ul><details className="advanced"><summary>Advanced workbook functions</summary><p>Add named formulas that call scalar VGI functions through <code>VGI.CALL</code>.</p><button disabled={disabled || !connection} onClick={() => void action("Adding workbook functions…", async () => createFunctionWrappers(await discoverFunctions(await resolveBackend(connection!.name))), "VGI functions added to this workbook.")}>Add VGI functions to this workbook</button></details></div>;
}

function ConnectionsPanel({ onSignIn, connections, defaultName, disabled, onChange, onDefault, action }: { onSignIn(name: string): Promise<void>; connections: OfficeConnection[]; defaultName: string; disabled: boolean; onChange(value: OfficeConnection[]): void; onDefault(value: string): void; action: Action }): React.JSX.Element {
  const catalogInput = useRef<HTMLInputElement>(null);
  const findCatalogsButton = useRef<HTMLButtonElement>(null);
  const catalogSelect = useRef<HTMLSelectElement>(null);
  const blank = (): OfficeConnection => ({ name: "", catalog: "", location: "", authentication: "anonymous", attachOptions: {} }); const active = connections.find((value) => value.name === defaultName) ?? connections[0]; const [form, setForm] = useState<OfficeConnection>(() => active ? { ...active } : blank()), [originalName, setOriginalName] = useState(active?.name ?? ""), [attachOptionsText, setAttachOptionsText] = useState(() => formatAttachOptionsJson(active?.attachOptions)), [validation, setValidation] = useState(""), [status, setStatus] = useState("");
  useEffect(() => { if (!originalName && active) { setForm({ ...active }); setNameEdited(true); setOriginalName(active.name); setAttachOptionsText(formatAttachOptionsJson(active.attachOptions)); } }, [active?.name]);
  const [catalogs, setCatalogs] = useState<string[]>([]), [manualCatalog, setManualCatalog] = useState(true), [nameEdited, setNameEdited] = useState(!!active);
  const [errorAt, setErrorAt] = useState<"discovery" | "connection">("connection");
  useEffect(() => { if (disabled) return; if (validation && errorAt === "discovery") findCatalogsButton.current?.focus(); else if (catalogs.length > 1) catalogSelect.current?.focus(); }, [disabled, validation, errorAt, catalogs]);
  const needsSignIn = useSyncExternalStore(connectionSignIn.subscribe, connectionSignIn.snapshot);
  useEffect(() => {
    if (!needsSignIn.includes(originalName)) setValidation(value => value === "Sign-in wasn’t completed. Try again when you’re ready." ? "" : value);
  }, [needsSignIn, originalName]);
  const signedIn = !!form.location && !!getServiceToken(form.location);
  function definition(): OfficeConnection { return { ...form, authentication: signedIn ? "oauth" : form.authentication ?? "anonymous", name: form.name.trim(), catalog: form.catalog?.trim() ?? "", location: form.location.trim(), attachOptions: parseAttachOptionsJson(attachOptionsText) }; }
  const nameError = connectionNameError(form.name, originalName, connections);
  async function discover(): Promise<void> {
    setErrorAt("discovery");
    setValidation(""); setStatus("Finding catalogs…"); setCatalogs([]);
    try {
      const found = await BrowserBackend.discoverCatalogs(form.location.trim(), setStatus);
      setCatalogs(found); setManualCatalog(!found.length);
      if (found.length === 1) setForm(value => catalogDraft(value, found[0], nameEdited));
      else if (found.length && !found.includes(form.catalog ?? "")) setForm(value => catalogDraft(value, "", nameEdited));
      setStatus(found.length ? `${found.length} catalog${found.length === 1 ? "" : "s"} available.` : "No catalogs were listed. Enter a catalog manually.");
    } catch (error) { setManualCatalog(true); setStatus(""); setValidation(catalogDiscoveryError(error)); }
  }
  let attachOptionsError = ""; try { parseAttachOptionsJson(attachOptionsText); } catch (error) { attachOptionsError = message(error); }
  const valid = !nameError && !!form.catalog?.trim() && !!form.name.trim() && !!form.location.trim() && !attachOptionsError;
  function save(): void { setErrorAt("connection"); if (nameError) { setValidation(nameError); return; } let prepared: OfficeConnection; try { prepared = definition(); assertHttpsConnection(prepared); setValidation(""); } catch (error) { setValidation(message(error)); return; } const next = [...connections.filter((item) => item.name !== originalName && item.name !== prepared.name), prepared]; onChange(next); setOriginalName(prepared.name); setNameEdited(true); setForm(prepared); setAttachOptionsText(formatAttachOptionsJson(prepared.attachOptions)); setStatus("Connection saved."); }
  async function test(): Promise<void> { setErrorAt("connection"); setStatus("Testing connection…"); try { const prepared = definition(); assertHttpsConnection(prepared); setValidation(""); await BrowserBackend.testConnection(prepared); const authenticated = !!getServiceToken(form.location); setForm((value) => ({ ...value, attachOptions: prepared.attachOptions, authentication: authenticated ? "oauth" : "anonymous" })); setStatus(authenticated ? "Connected · signed in securely." : "Connected · no sign-in required."); } catch (error) { setStatus(""); setValidation(message(error)); captureError(error, "connection.test"); } }
  function change(patch: Partial<OfficeConnection>): void { setForm({ ...form, ...patch }); setStatus(""); setValidation(""); }
  return <div className="connection-layout">
    <fieldset className="connection-sidebar" disabled={disabled}><button className="primary new-connection" onClick={() => { setForm(blank()); setCatalogs([]); setManualCatalog(true); setNameEdited(false); setOriginalName(""); setAttachOptionsText(""); setStatus(""); setValidation(""); }}>New connection</button><div className="connection-list">{connections.map((connection) => { const authenticated = !!getServiceToken(connection.location); return <button key={connection.name} className={`${connection.name === originalName ? "connection selected-connection" : "connection"} ${connection.name === defaultName ? "active-connection" : ""}`} onClick={() => { setForm({ ...connection }); setCatalogs([]); setManualCatalog(true); setNameEdited(true); setOriginalName(connection.name); setAttachOptionsText(formatAttachOptionsJson(connection.attachOptions)); setValidation(""); setStatus(""); setValidation(""); }} title={connection.location}><span className="health-dot configured"/><span><strong>{connection.name}{connection.name === defaultName ? " · active" : ""}</strong><small>{connection.catalog ?? connection.name} · {needsSignIn.includes(connection.name) ? "Sign-in needed" : authenticated ? "Sign-in saved" : "Ready"}</small></span></button>; })}</div></fieldset>
    <fieldset className="connection-form" disabled={disabled}>
      <h3>{originalName ? `Edit ${originalName}` : "New connection"}</h3>
      <label>Server address<input value={form.location} onChange={(event) => { setForm({ ...form, location: event.target.value, authentication: "anonymous" }); setCatalogs([]); setManualCatalog(true); setStatus(""); setValidation(""); }} placeholder="https://vgi.example.com/"/></label>
      <div className="actions"><button ref={findCatalogsButton} disabled={disabled || !form.location.trim()} onClick={() => void action("Finding catalogs…", discover, null)}>Find catalogs</button></div>
      {validation && errorAt === "discovery" && <p className="field-error" role="alert">{validation}</p>}
      <small className="connection-address-help">Paste the HTTPS address provided by your data service, then find its catalogs.</small>
      <div className="catalog-choice-row">
      {catalogs.length > 0 && !manualCatalog ? <label>Catalog<select ref={catalogSelect} aria-label="Catalog" value={form.catalog ?? ""} onChange={event => { setForm(value => catalogDraft(value, event.target.value, nameEdited)); setStatus(""); setValidation(""); }}><option value="">Select a catalog</option>{catalogs.map(catalog => <option key={catalog} value={catalog}>{catalog}</option>)}</select></label> : <label>Catalog<input ref={catalogInput} value={form.catalog ?? ""} onChange={event => { setForm(value => catalogDraft(value, event.target.value, nameEdited)); setStatus(""); setValidation(""); }} placeholder="open_meteo"/></label>}
      {catalogs.length > 0 && <button className="catalog-choice-action" type="button" onClick={() => setManualCatalog(!manualCatalog)}>{manualCatalog ? "Choose from list" : "Enter catalog manually"}</button>}
      </div>
      <small className="connection-auth-hint">If this service requires authentication, Cupola opens a secure sign-in window when it connects.</small>
      <label>Connection name<input value={form.name} onChange={(event) => (setNameEdited(true), change({ name: event.target.value }))}/></label>
      <small className="connection-field-help">The name you’ll use to find this connection in Cupola and Excel.</small>
      <details className="advanced connection-advanced"><summary>Advanced options</summary><label htmlFor="office-attach-options">Options as JSON<textarea id="office-attach-options" value={attachOptionsText} onChange={(event) => { setAttachOptionsText(event.target.value); setStatus(""); setValidation(""); }} rows={4} spellCheck={false} placeholder={'{"region":"us-east"}'}/></label><small>Non-secret string, number, boolean, or null values. Cupola manages TYPE, LOCATION, and OAuth credentials.</small>{attachOptionsError && <p className="field-error" role="alert">{attachOptionsError}</p>}</details>
      {validation && errorAt === "connection" && <p className="field-error" role="alert">{validation}</p>}
      {nameError && <p className="field-error" role="alert">{nameError}</p>}
      <div className="actions connection-actions"><button disabled={disabled || !valid} onClick={() => void action("Testing connection…", test, null)}>{disabled && status.startsWith("Testing") ? "Testing…" : "Test connection"}</button><button className="primary" disabled={disabled || !valid} onClick={() => { try { save(); } catch (error) { setStatus(""); setValidation(message(error)); captureError(error, "connection.save"); } }}>Save changes</button>{originalName && originalName !== defaultName && <button onClick={() => { onDefault(originalName); setStatus(`${originalName} is now active.`); }}>Use as active</button>}</div>
      {status && <p className="connection-status" role="status">{status}</p>}
      {originalName && (form.authentication === "oauth" || signedIn) && <button disabled={disabled} onClick={() => { setValidation(""); void onSignIn(originalName).then(() => setStatus("Signed in. You can run your query again.")).catch(() => setValidation("Sign-in wasn’t completed. Try again when you’re ready.")); }}>Sign in again</button>}
      {signedIn && <div className="oauth-card"><div><strong>{needsSignIn.includes(originalName) ? "Sign-in needed" : "Sign-in saved"}</strong><small>Tokens are held for this Excel web session and used automatically.</small></div><button onClick={() => { sessionStorage.removeItem(sessionTokenKey(form.location)); resetRuntime(); setForm({ ...form, authentication: "oauth" }); setStatus("Signed out. Cupola will prompt again if this service requires authentication."); }}>Sign out</button></div>}
      {originalName && <div className="danger-zone"><button className="danger" onClick={async () => { if (!await confirmAction(`Remove “${originalName}” and its saved OAuth session?`, "Remove connection", "Remove connection")) return; sessionStorage.removeItem(sessionTokenKey(form.location)); const next = connections.filter((item) => item.name !== originalName); onChange(next); if (defaultName === originalName && next[0]) onDefault(next[0].name); setForm(blank()); setCatalogs([]); setManualCatalog(true); setNameEdited(false); setOriginalName(""); setAttachOptionsText(""); setStatus(""); setValidation(""); }}>Remove connection</button></div>}
    </fieldset>
  </div>;
}

function AgentActivity({ value }: { value: string }): React.JSX.Element { return <div className="agent-activity" role="status" aria-live="polite"><span className="agent-activity-dot" aria-hidden="true"/><span>{value}</span></div>; }
function ToolCall({ value, onOpenQuery }: { value: ToolEvent; onOpenQuery(sql: string): void }): React.JSX.Element { return <details className={`tool ${value.state}`} open={value.state === "error"}><summary><span className="tool-dot"/> {toolLabel(value.name, value.state)}</summary>{value.detail && <pre>{value.detail}</pre>}{value.name === "run_sql" && (value.sql || value.state === "done") && <button onClick={() => onOpenQuery(value.sql || value.detail || "")}>Open in Query Editor</button>}</details>; }
function WorkbookOutcome({ value }: { value: WorkbookWriteOutcome }): React.JSX.Element { return <div className="workbook-outcome" role="status"><div><strong>Snapshot inserted</strong><span>{value.table} · {value.sheet}!{value.address} · {value.rows.toLocaleString()} rows</span></div><button onClick={() => void goToTable(value.table)}>Go to table</button></div>; }
function ResultGrid({ result, label = "Results", onCopy, limit = 200, offset = 0 }: { result: QueryResult | null; label?: string; onCopy?(): void | Promise<unknown>; limit?: number; offset?: number }): React.JSX.Element | null { if (!result) return null; const shown = result.rows.slice(offset, offset + limit); return <div className="result-wrap"><div className="result-meta"><span>{result.rowCount.toLocaleString()} rows{result.elapsedMs != null ? ` · ${Math.round(result.elapsedMs)} ms` : ""}{shown.length < result.rows.length ? " · preview limited" : ""}</span>{onCopy && <button onClick={() => void onCopy()}>Copy results</button>}</div><table aria-label={label}><caption className="sr-only">{label}</caption><thead><tr>{result.columns.map((column) => <th key={column.name} scope="col"><span>{column.name}</span><small>{column.type}</small></th>)}</tr></thead><tbody>{shown.map((row, rowIndex) => <tr key={offset + rowIndex}>{row.map((cell, columnIndex) => <td key={columnIndex}>{cell == null ? <em>NULL</em> : String(cell)}</td>)}</tr>)}</tbody></table></div>; }

function message(error: unknown): string { return error instanceof Error ? error.message : String(error); }
function literal(value: string): string { return `'${value.replaceAll("'", "''")}'`; }
function qualified(value: CatalogObject): string { return [value.catalog, value.schema, value.name].map((part) => `"${part.replaceAll('"', '""')}"`).join("."); }
function functionSignature(value: CatalogFunction): string { return `${value.catalog}.${value.schema}.${value.name}(${value.parameters.map((parameter) => `${parameter.name}${parameter.kind === "named" ? " := " : " "}${parameter.type}`).join(", ")})`; }
function functionQuery(value: CatalogFunction): string { const args = value.parameters.filter((parameter) => parameter.default == null).map((parameter) => parameter.kind === "named" ? `${parameter.name} := NULL::${parameter.type}` : `NULL::${parameter.type} /* ${parameter.name} */`).join(", "); return `SELECT *\nFROM ${[value.catalog, value.schema, value.name].map((part) => `"${part.replaceAll('"', '""')}"`).join(".")}(${args});`; }
function functionArguments(value: CatalogFunction): QueryResult { const columns = ["name", "type", "kind", "default", "description"]; return { columns: columns.map((name) => ({ name, type: "VARCHAR" })), rows: value.parameters.map((parameter) => [parameter.name, parameter.type, parameter.kind ?? null, parameter.default == null ? null : JSON.stringify(parameter.default), parameter.description ?? null]), rowCount: value.parameters.length }; }
function toolLabel(name: string, state: ToolEvent["state"]): string { if (name === "ask_clarification") return state === "running" ? "Waiting for your answer" : state === "writing" ? "Preparing a question" : state === "done" ? "Clarification · answered" : `Clarification · ${state}`; const names: Record<string, string> = { run_sql: "SQL query", list_tables: "Catalog inventory", list_functions: "Function inventory", describe_table: "Table description", create_query_tab: "Creating query tab" }; return `${names[name] ?? name} · ${state === "writing" ? "preparing" : state === "running" ? "running" : state === "done" ? "complete" : state === "stopped" ? "stopped" : "error"}`; }
