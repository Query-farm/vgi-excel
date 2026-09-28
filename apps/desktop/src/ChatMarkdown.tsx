import { tableClipboard } from "@query-farm/vgi-excel-core";
import { useState, useRef, type ReactNode } from "react";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { host } from "./bridge";

export function ChatMarkdown({ content, streaming = false, onOpenQuery }: { content: string; streaming?: boolean; onOpenQuery?(sql: string): void }): React.JSX.Element {
  const markdown = streaming ? completeStreamingMarkdown(content) : content;
  return <div className="message-markdown">
    <ReactMarkdown
      remarkPlugins={[remarkGfm]}
      skipHtml
      components={{
        h1: ({ children }) => <h2>{children}</h2>,
        h2: ({ children }) => <h3>{children}</h3>,
        h3: ({ children }) => <h4>{children}</h4>,
        h4: ({ children }) => <h4>{children}</h4>,
        a: ({ href, children }) => href
          ? <a href={href} target="_blank" rel="noopener noreferrer">{children}</a>
          : <span>{children}</span>,
        img: ({ alt }) => <span className="markdown-image">[Image: {alt || "image"}]</span>,
        code: ({ className, children }) => {
          if (!className) return <code>{children}</code>;
          return <CodeBlock className={className} value={String(children).replace(/\n$/, "")} onOpenQuery={streaming ? undefined : onOpenQuery}/>;
        },
        pre: ({ children }) => <>{children}</>,
        table: ({ children }) => <CopyableTable disabled={streaming}>{children}</CopyableTable>,
      }}
    >{markdown}</ReactMarkdown>
  </div>;
}

function CodeBlock({ className, value, onOpenQuery }: { className: string; value: string; onOpenQuery?(sql: string): void }): React.JSX.Element {
  const [copied, setCopied] = useState(false);
  const sql = /(^|\s)language-sql(\s|$)/i.test(className);
  async function copy(): Promise<void> {
    try {
      if (navigator.clipboard?.writeText) await navigator.clipboard.writeText(value);
      else await host.copyText(value);
    } catch {
      await host.copyText(value);
    }
    setCopied(true);
    globalThis.setTimeout(() => setCopied(false), 1_500);
  }
  return <pre className="markdown-code"><button type="button" className="copy-code" aria-label={`Copy ${sql ? "SQL" : "code"}`} onClick={() => void copy()}>{copied ? "Copied" : sql ? "Copy SQL" : "Copy"}</button><code className={className}>{value}</code>{sql && onOpenQuery && <button type="button" className="open-ai-query" onClick={() => onOpenQuery(value)}>Open in Query Editor</button>}</pre>;
}

export function completeStreamingMarkdown(content: string): string {
  let value = content.endsWith("\n") ? content : `${content}\n`;
  const fences = value.match(/^\s*```/gm)?.length ?? 0;
  if (fences % 2 === 1) value += "```\n";
  return value;
}

function CopyableTable({ children, disabled }: { children: ReactNode; disabled: boolean }): React.JSX.Element {
  const table = useRef<HTMLTableElement>(null);
  const [status, setStatus] = useState("");
  async function copy(): Promise<void> {
    if (!table.current) return;
    const payload = tableClipboard(Array.from(table.current.rows, row => Array.from(row.cells, cell => cell.textContent ?? "")));
    try {
      let rich = false;
      if (typeof ClipboardItem !== "undefined" && navigator.clipboard?.write) {
        try { await navigator.clipboard.write([new ClipboardItem({ "text/plain": new Blob([payload.text], { type: "text/plain" }), "text/html": new Blob([payload.html], { type: "text/html" }) })]); rich = true; } catch { /* Fall back to text for WebViews without HTML clipboard support. */ }
      }
      if (!rich) {
        try { await navigator.clipboard.writeText(payload.text); }
        catch (error) { await host.copyText(payload.text); }
      }
      setStatus("Table copied");
    } catch { setStatus("Could not copy the table. Select the cells and copy them manually."); }
  }
  return <div className="markdown-table-copy"><div className="markdown-table-actions"><button type="button" disabled={disabled} aria-label="Copy table" onClick={() => void copy()}>Copy table</button>{status && <span role="status">{status}</span>}</div><div className="markdown-table"><table ref={table}>{children}</table></div></div>;
}
