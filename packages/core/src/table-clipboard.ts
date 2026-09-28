/** Plain text + HTML clipboard representations, adapted from vgi-web-frontend 7d5ef8e. */
export function tableClipboard(rows: readonly (readonly string[])[]): { text: string; html: string } {
  const escape = (text: string) => text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
  const tsv = (text: string) => /[\t\r\n"]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
  return {
    text: rows.map(row => row.map(tsv).join("\t")).join("\n"),
    html: `<table>${rows.map(row => `<tr>${row.map(cell => `<td>${escape(cell)}</td>`).join("")}</tr>`).join("")}</table>`,
  };
}
