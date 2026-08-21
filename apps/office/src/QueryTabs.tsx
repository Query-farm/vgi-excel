import { useEffect, useRef, useState } from "react";
import { Plus, X } from "lucide-react";
type TabDocument = { id: string; name: string };

export function QueryTabs({ documents, activeId, runningId, label = "Saved queries", itemName = "query", onSelect, onAdd, onClose, onRename }: {
  documents: TabDocument[];
  activeId: string;
  runningId?: string | null;
  label?: string;
  itemName?: string;
  onSelect(id: string): void;
  onAdd(): void;
  onClose(id: string): void;
  onRename(id: string, name: string): void;
}): React.JSX.Element {
  const [editingId, setEditingId] = useState<string | null>(null);
  const [draft, setDraft] = useState("");
  const input = useRef<HTMLInputElement | null>(null);
  const tabs = useRef<Array<HTMLButtonElement | null>>([]);

  useEffect(() => { if (editingId) { input.current?.focus(); input.current?.select(); } }, [editingId]);

  function begin(document: TabDocument): void { setEditingId(document.id); setDraft(document.name); }
  function commit(): void {
    if (editingId && draft.trim()) onRename(editingId, draft);
    setEditingId(null);
  }
  function keyboard(event: React.KeyboardEvent<HTMLButtonElement>, index: number): void {
    if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return;
    event.preventDefault();
    const next = event.key === "Home" ? 0 : event.key === "End" ? documents.length - 1 : (index + (event.key === "ArrowRight" ? 1 : -1) + documents.length) % documents.length;
    onSelect(documents[next].id);
    tabs.current[next]?.focus();
  }

  return <div className="query-tabs" role="tablist" aria-label={label}>
    {documents.map((document, index) => {
      const active = document.id === activeId;
      return <div className={`query-tab${active ? " active" : ""}`} role="presentation" key={document.id}>
        {editingId === document.id ? <input ref={input} value={draft} aria-label={`Rename ${document.name}`} onChange={(event) => setDraft(event.target.value)} onBlur={commit} onKeyDown={(event) => { if (event.key === "Enter") commit(); else if (event.key === "Escape") setEditingId(null); }}/> : <button ref={(node) => { tabs.current[index] = node; }} type="button" role="tab" aria-selected={active} tabIndex={active ? 0 : -1} title={`${document.name} · saved locally · double-click to rename`} onClick={() => onSelect(document.id)} onDoubleClick={() => begin(document)} onKeyDown={(event) => keyboard(event, index)}><span>{document.name}</span>{runningId === document.id && <i aria-label="Query running" title="Query running"/>}</button>}
        <button type="button" className="close-query-tab" aria-label={`Close ${document.name}`} title={`Close ${itemName} tab`} onClick={(event) => { event.stopPropagation(); onClose(document.id); }}><X aria-hidden="true"/></button>
      </div>;
    })}
    <button type="button" className="add-query-tab" aria-label={`New ${itemName} tab`} title={`New ${itemName} tab`} onClick={onAdd}><Plus aria-hidden="true"/></button>
  </div>;
}
