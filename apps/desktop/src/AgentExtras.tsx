import { agentQueryName } from "@query-farm/vgi-excel-core";
import { Fragment, type ReactNode } from "react";
import type { AgentTranscriptPart } from "@query-farm/vgi-excel-core";
import { ChatMarkdown } from "./ChatMarkdown";
import { readAIRequestDiagnostics, readAIToolDiagnostics, clearAIRequestDiagnostics } from "@query-farm/vgi-excel-core";
import { useEffect, useRef, useState } from "react";
import type { AgentClarification, AgentResultCard } from "@query-farm/vgi-excel-core";
import { openResultsWindow } from "./results-window";

export function useClarification() {
  const [question, setQuestion] = useState<AgentClarification | null>(null);
  const pending = useRef<((answer: string) => void) | null>(null);
  return { question,
    answer(value: string) { if (value.trim()) pending.current?.(value.trim()); },
    request(value: AgentClarification, signal?: AbortSignal): Promise<string> {
      signal?.throwIfAborted(); setQuestion(value);
      return new Promise<string>((resolve, reject) => {
        const cleanup = () => { pending.current = null; setQuestion(null); signal?.removeEventListener("abort", abort); };
        const abort = () => { cleanup(); reject(new DOMException("Stopped", "AbortError")); };
        pending.current = answer => { cleanup(); resolve(answer); };
        signal?.addEventListener("abort", abort, { once: true });
        if (signal?.aborted) abort();
      });
    },
  };
}
export function Clarification({ value, enabled, choose }: { value: AgentClarification; enabled: boolean; choose(value: string): void }): React.JSX.Element {
  return <div className="agent-clarification"><strong>{value.question}</strong>{value.answer ? <p>Your answer: {value.answer}</p> : <><div className="clarification-choices">{value.options.map(option => <button key={option} disabled={!enabled} onClick={() => choose(option)}>{option}</button>)}</div>{enabled && <p>Choose an option or type your answer below, then Continue.</p>}</>}</div>;
}
export function AgentProgress({ stage, waiting }: { stage: string; waiting: boolean }): React.JSX.Element {
  const [seconds, setSeconds] = useState(0);
  const activeMs = useRef(0);
  useEffect(() => { if (waiting) return; const start = performance.now(); const timer = setInterval(() => setSeconds(Math.floor((activeMs.current + performance.now() - start) / 1000)), 1000); return () => { activeMs.current += performance.now() - start; clearInterval(timer); }; }, [waiting]);
  return <span className="agent-progress"><span role="status">{stage}</span>{!waiting && <span aria-label="Active elapsed time"> · {seconds}s</span>}</span>;
}
export function AgentResult({ value, disabled, openQuery, load, loadLabel, hideLoad = false }: { hideLoad?: boolean; value: AgentResultCard; disabled: boolean; openQuery(sql: string): void; load(value: AgentResultCard): Promise<string | void>; loadLabel: string }): React.JSX.Element {
  const [status, setStatus] = useState("");
  const [busy, setBusy] = useState(false);
  const incomplete = !!value.result.truncated || value.result.rows.length < value.result.rowCount;
  async function insert(): Promise<void> {
    if (!window.confirm(loadLabel === "Load into Excel" ? `Create a refreshable Excel table named “${agentQueryName(value.queryName)}”? This reruns the query using the saved connection. Excel will add a suffix if the name is already in use.` : `Insert ${value.result.rows.length.toLocaleString()} loaded rows as a static Excel table at the current selection?`)) return;
    setBusy(true); setStatus("");
    try { const message = await load(value); setStatus(message || (loadLabel === "Load into Excel" ? "Excel load requested. Check Queries & Connections for refresh status." : "Snapshot inserted.")); }
    catch (error) { setStatus(error instanceof Error ? error.message : "Could not insert the result. Try again."); }
    finally { setBusy(false); }
  }
  return <section className="agent-result" aria-label="AI query result"><strong>Query result</strong><p>Connection: {value.connection}{value.catalog && ` · Catalog: ${value.catalog}`}</p><p>Showing {Math.min(5, value.result.rows.length)} of {value.result.rows.length.toLocaleString()} loaded rows.{value.result.truncated ? " The loaded result is incomplete." : ""} {value.result.rowCount > value.result.rows.length && ` Query reports ${value.result.rowCount.toLocaleString()} rows.`} Initial AI sample: {Math.min(20, value.result.rows.length)} {value.result.rows.length === 1 ? "row" : "rows"}.</p><details><summary>Scope and SQL</summary><p>AI-described scope — verify against the SQL:</p><dl><dt>Date interpretation</dt><dd>{value.scope.dateRange || "Not specified"}</dd><dt>Filters</dt><dd>{value.scope.filters || "Not specified"}</dd><dt>Assumptions</dt><dd>{value.scope.assumptions || "Not specified"}</dd></dl><pre><code>{value.sql}</code></pre></details><div className="agent-result-grid" tabIndex={0} aria-label="Scrollable AI preview"><table><thead><tr>{value.result.columns.map((column, index) => <th key={index}>{column.name}</th>)}</tr></thead><tbody>{value.result.rows.slice(0, 5).map((row, index) => <tr key={index}>{row.map((cell, col) => <td key={col}>{cell == null ? "NULL" : String(cell)}</td>)}</tr>)}</tbody></table>{!value.result.rows.length && <p>No rows returned.</p>}</div><div className="agent-result-actions"><button onClick={() => void openResultsWindow({ title: "AI query result", result: value.result }).catch(error => setStatus(error instanceof Error ? error.message : "Could not open results."))}>Open window</button><button onClick={() => openQuery(value.sql)}>Edit query</button>{!hideLoad && <button className="primary" disabled={disabled || busy || !value.result.rows.length || (loadLabel === "Insert into Excel" && incomplete)} onClick={() => void insert()}>{busy ? "Loading…" : loadLabel}</button>}</div>{loadLabel === "Insert into Excel" && <small>{incomplete ? "Incomplete preview: use Edit query to load the full result into Excel." : "Static table · no refresh connection"}</small>}{status && <p role="status">{status}</p>}</section>;
}

export function AIRequestDiagnostics(): React.JSX.Element {
  const snapshot = () => ({ modelRequests: readAIRequestDiagnostics(), toolCalls: readAIToolDiagnostics() });
  const [requests, setRequests] = useState(snapshot);
  return <details className="ai-local-diagnostics"><summary>AI performance diagnostics</summary><p>Local to this session. Latest 50 model requests and 100 tool calls; no conversation content is recorded here. Model times include retries and planning. Tool times measure execution; clarification time includes waiting for your answer.</p><button onClick={() => setRequests(snapshot())}>Refresh diagnostics</button><button onClick={() => { clearAIRequestDiagnostics(); setRequests(snapshot()); }}>Clear diagnostics</button><pre>{JSON.stringify(requests, null, 2)}</pre></details>;
}

export function AgentTranscript<T extends { id: string; name: string; state: string }>({ message, renderTool, renderResult, onOpenQuery }: {
  message: { text: string; timeline?: AgentTranscriptPart[]; tools?: T[]; results?: AgentResultCard[]; streaming?: boolean };
  renderTool(tool: T): ReactNode; renderResult(result: AgentResultCard): ReactNode; onOpenQuery(sql: string): void;
}): React.JSX.Element {
  const parts = message.timeline;
  const text = (value: string) => <ChatMarkdown content={value} streaming={message.streaming} onOpenQuery={onOpenQuery}/>;
  if (!parts?.length) return <>{message.tools?.map(tool => <Fragment key={tool.id}>{renderTool(tool)}</Fragment>)}{message.text && text(message.text)}{message.results?.map(result => <Fragment key={result.id}>{renderResult(result)}</Fragment>)}</>;
  const shownText = parts.filter(part => part.type === "text").map(part => "text" in part ? part.text : "").join("");
  const remaining = message.text.startsWith(shownText) ? message.text.slice(shownText.length) : "";
  return <>{parts.map((part, index) => {
    if (part.type === "text") return <Fragment key={index}>{text(part.text)}</Fragment>;
    if (part.type === "thinking") return <details className="thinking-summary" key={index}><summary>{message.streaming && index === parts.length - 1 ? "Thinking…" : "Thinking summary"}</summary><div>{part.text}</div></details>;
    const tool = message.tools?.find(tool => tool.id === part.id);
    const results = message.results?.filter(result => result.toolCallId === part.id) ?? [];
    return <div className="agent-tool-step" key={index}>{tool && renderTool(tool)}{results.map(result => <Fragment key={result.id}>{renderResult(result)}</Fragment>)}{tool?.name === "run_sql" && tool.state === "done" && !results.length && !message.streaming && <p className="hint">This preview is no longer available. Open the query in Query Editor to run it again.</p>}</div>;
  })}{remaining && text(remaining)}{message.results?.filter(result => !result.toolCallId || !parts.some(part => part.type === "tool" && part.id === result.toolCallId)).map(result => <Fragment key={result.id}>{renderResult(result)}</Fragment>)}</>;
}
