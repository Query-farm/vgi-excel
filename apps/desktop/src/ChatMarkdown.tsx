import { useState } from "react";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { host } from "./bridge";

export function ChatMarkdown({ content, streaming = false }: { content: string; streaming?: boolean }): React.JSX.Element {
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
          return <CodeBlock className={className} value={String(children).replace(/\n$/, "")}/>;
        },
        pre: ({ children }) => <>{children}</>,
        table: ({ children }) => <div className="markdown-table"><table>{children}</table></div>,
      }}
    >{markdown}</ReactMarkdown>
  </div>;
}

function CodeBlock({ className, value }: { className: string; value: string }): React.JSX.Element {
  const [copied, setCopied] = useState(false);
  const sql = /(^|\s)language-sql(\s|$)/.test(className);
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
  return <pre className="markdown-code"><button type="button" className="copy-code" aria-label={`Copy ${sql ? "SQL" : "code"}`} onClick={() => void copy()}>{copied ? "Copied" : sql ? "Copy SQL" : "Copy"}</button><code className={className}>{value}</code></pre>;
}

export function completeStreamingMarkdown(content: string): string {
  let value = content.endsWith("\n") ? content : `${content}\n`;
  const fences = value.match(/^\s*```/gm)?.length ?? 0;
  if (fences % 2 === 1) value += "```\n";
  return value;
}
