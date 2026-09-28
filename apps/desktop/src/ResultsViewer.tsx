import { notifyRendered } from "./rendered";
import { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import type { QueryResult } from "@query-farm/vgi-excel-core";
import { ChevronLeft, ChevronRight, Copy, Maximize2, X } from "lucide-react";
import { receiveResults, closeResults, maximizeResults, nativeResultsWindow } from "./results-window";
import "./results-viewer.css";

export interface ResultsSnapshot { title: string; result: QueryResult }
function ResultsViewer(): React.JSX.Element {
  const [snapshot, setSnapshot] = useState<ResultsSnapshot | null>(null);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("");
  const [offset, setOffset] = useState(0);
  const [limit, setLimit] = useState(200);
  useEffect(() => receiveResults(value => { setSnapshot(value); document.title = `${value.title} — Results — Cupola for Excel by Query Farm`; }, () => setError("Results could not be loaded. Close this window and open it again from the query editor.")), []);
  useEffect(notifyRendered, []);
  const result = snapshot?.result;
  async function copy(): Promise<void> {
    if (!result) return;
    try {
      await navigator.clipboard.writeText([result.columns.map(c => c.name), ...result.rows].map(row => row.map(cell => cell == null ? "" : String(cell).replaceAll("\t", " ").replaceAll("\n", " ")).join("\t")).join("\n"));
      setStatus("Loaded results copied.");
    } catch { setStatus("Could not copy results. Check clipboard access and try again."); }
  }
  return <main className="results-window">
    <header><div className="window-title"><img src="./cupola-mark.svg" alt=""/><div><h1>{snapshot?.title ?? "Query results"} — Results</h1><p>Results captured when this window opened.</p></div></div><div className="window-actions">
      {result && <button onClick={() => void copy()}><Copy aria-hidden="true"/>Copy loaded results</button>}
      {(nativeResultsWindow || document.fullscreenEnabled) && <button onClick={() => void maximizeResults().catch(() => setStatus("Full screen is unavailable. You can still resize this window."))}><Maximize2 aria-hidden="true"/>{nativeResultsWindow ? "Maximize window" : "Full screen"}</button>}
      <button onClick={closeResults}><X aria-hidden="true"/>Close</button>
    </div></header>
    <p role="status" className="viewer-status">{status}</p>
    {error ? <p role="alert">{error}</p> : !result ? <p role="status">Loading results…</p> : <>
      <div className="viewer-toolbar"><p>{result.truncated || result.rows.length < result.rowCount ? `${result.rows.length.toLocaleString()} of ${result.rowCount.toLocaleString()} rows loaded. This window shows only the loaded preview.` : `${result.rowCount.toLocaleString()} rows loaded.`}</p><div className="paging">
        <label>Rows per page<select value={limit} onChange={event => { setLimit(Number(event.target.value)); setOffset(0); }}>{[200, 500, 1000, 2000].map(value => <option key={value} value={value}>{value.toLocaleString()}</option>)}</select></label>
        <button aria-label="Previous page" disabled={offset === 0} onClick={() => setOffset(Math.max(0, offset - limit))}><ChevronLeft/></button>
        <span>{result.rows.length ? offset + 1 : 0}–{Math.min(offset + limit, result.rows.length)} of {result.rows.length.toLocaleString()}</span>
        <button aria-label="Next page" disabled={offset + limit >= result.rows.length} onClick={() => setOffset(offset + limit)}><ChevronRight/></button>
      </div></div>
      <div className="viewer-grid" tabIndex={0} aria-label="Scrollable query results"><table aria-label="Query results"><thead><tr>{result.columns.map((column, i) => <th key={i} scope="col">{column.name}<small>{column.type}</small></th>)}</tr></thead><tbody>{result.rows.slice(offset, offset + limit).map((row, i) => <tr key={offset + i}>{row.map((cell, j) => <td key={j}>{cell == null ? <em>NULL</em> : String(cell)}</td>)}</tr>)}</tbody></table>{!result.rows.length && <p className="empty-results">This query returned no rows.</p>}</div>
    </>}
  </main>;
}
createRoot(document.getElementById("root")!).render(<ResultsViewer/>);
