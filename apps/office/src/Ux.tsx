import { useRef, type ReactNode } from "react";

export type NoticeValue = { kind: "progress" | "success" | "error" | "info"; message: string } | null;

export function Notice({ value, onDismiss, onRetry, onDiagnostics }: { value: NoticeValue; onDismiss?(): void; onRetry?(): void; onDiagnostics?(): void }): React.JSX.Element | null {
  if (!value) return null;
  return <div className={`notice ${value.kind}`} role={value.kind === "error" ? "alert" : "status"} aria-live={value.kind === "error" ? "assertive" : "polite"}><span className="notice-icon" aria-hidden="true">{value.kind === "error" ? "!" : value.kind === "success" ? "✓" : value.kind === "progress" ? "…" : "i"}</span><span>{value.message}</span><div className="notice-actions">{onRetry && <button onClick={onRetry}>Retry</button>}{value.kind === "error" && onDiagnostics && <button onClick={onDiagnostics}>Copy diagnostics</button>}{onDismiss && value.kind !== "progress" && <button aria-label="Dismiss message" onClick={onDismiss}>×</button>}</div></div>;
}

export function Onboarding({ onConnect }: { onConnect(): void }): React.JSX.Element {
  return <section className="onboarding"><img src="./cupola-mark.svg" alt=""/><p className="eyebrow">Cupola for Excel</p><h2>Connect a VGI data source to begin</h2><p>Explore a catalog, run read-only SQL, and place reviewed snapshots into your workbook over HTTPS.</p><div className="actions"><button className="primary" onClick={onConnect}>Add connection</button><a className="button-link" href="https://query.farm" target="_blank" rel="noreferrer">Learn about VGI connections</a></div></section>;
}

export function WorkspaceTabs<T extends string>({ value, tabs, onChange }: { value: T; tabs: Array<{ id: T; label: string; icon?: ReactNode }>; onChange(value: T): void }): React.JSX.Element {
  const refs = useRef<Array<HTMLButtonElement | null>>([]);
  return <nav className="workspace-tabs" role="tablist" aria-label="Cupola workspaces">{tabs.map((tab, index) => <button key={tab.id} ref={(node) => { refs.current[index] = node; }} role="tab" id={`tab-${tab.id}`} aria-label={tab.label} aria-selected={value === tab.id} aria-controls={`panel-${tab.id}`} tabIndex={value === tab.id ? 0 : -1} className={value === tab.id ? "active" : ""} onClick={() => onChange(tab.id)} onKeyDown={(event) => { if (event.key !== "ArrowLeft" && event.key !== "ArrowRight" && event.key !== "Home" && event.key !== "End") return; event.preventDefault(); const next = event.key === "Home" ? 0 : event.key === "End" ? tabs.length - 1 : (index + (event.key === "ArrowRight" ? 1 : -1) + tabs.length) % tabs.length; onChange(tabs[next].id); refs.current[next]?.focus(); }}>{tab.icon}<span>{tab.label}</span></button>)}</nav>;
}

export function TabPanel({ id, active, children, busy }: { id: string; active: boolean; children: React.ReactNode; busy?: boolean }): React.JSX.Element {
  return <div role="tabpanel" id={`panel-${id}`} aria-labelledby={`tab-${id}`} aria-busy={busy || undefined} hidden={!active} className="workspace-panel">{children}</div>;
}

export function formatSql(sql: string): string {
  return sql.trim().replace(/\s+/g, " ").replace(/\b(SELECT|FROM|WHERE|GROUP BY|ORDER BY|HAVING|LIMIT|UNION ALL|UNION|WITH|JOIN|LEFT JOIN|RIGHT JOIN|INNER JOIN|ON)\b/gi, (word) => `\n${word.toUpperCase()}`).replace(/ *\n */g, "\n").trim();
}

export function resultTsv(result: { columns: Array<{ name: string }>; rows: unknown[][] }): string {
  return [result.columns.map((column) => column.name), ...result.rows].map((row) => row.map((cell) => cell == null ? "" : String(cell).replaceAll("\t", " ").replaceAll("\n", " ")).join("\t")).join("\n");
}
