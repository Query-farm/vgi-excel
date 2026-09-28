import { useEffect, useId, useRef, useState } from "react";
import { AI_MODELS, MODEL_CACHE_AGE, listAIModels, modelCacheScope, readModelCache, saveModelCache, type AIModelInfo } from "@query-farm/vgi-excel-core";

export function useModelCatalog(apiKey: string, workspace: string, active: boolean) {
  const [models, setModels] = useState<AIModelInfo[]>([]);
  const [status, setStatus] = useState("");
  const [loading, setLoading] = useState(false);
  const [revision, refresh] = useState(0);
  const previousRevision = useRef(0);
  useEffect(() => {
    const controller = new AbortController();
    const force = revision !== previousRevision.current;
    previousRevision.current = revision;
    setModels([]); setStatus(""); setLoading(false);
    if (!apiKey.trim()) return;
    // Let users finish pasting/editing credentials before issuing a request.
    const timer = setTimeout(() => void (async () => {
      let cached = false;
      try {
        const scope = await modelCacheScope(apiKey, workspace);
        if (controller.signal.aborted) return;
        const saved = readModelCache(localStorage, scope);
        if (saved) { setModels(saved.models); cached = true; }
        if (!active || (!force && saved && Date.now() - saved.updatedAt >= 0 && Date.now() - saved.updatedAt < MODEL_CACHE_AGE)) return;
        setLoading(true); setStatus("Updating models…");
        const values = await listAIModels(apiKey, workspace, { signal: controller.signal });
        if (controller.signal.aborted) return;
        saveModelCache(localStorage, { scope, models: values, updatedAt: Date.now() });
        setModels(values); setStatus("");
      } catch (error) {
        if (!controller.signal.aborted) setStatus(`${error instanceof Error ? error.message : "Could not update models. Try again."} ${cached ? "Showing the saved list." : "You can still enter a model ID manually."}`);
      } finally { if (!controller.signal.aborted) setLoading(false); }
    })(), force ? 0 : 500);
    return () => { clearTimeout(timer); controller.abort(); };
  }, [apiKey, workspace, active, revision]);
  return { models, status, loading, refresh: () => refresh(value => value + 1) };
}

export function ModelCatalog({ catalog, disabled, selectionDisabled, value, onChange }: { catalog: ReturnType<typeof useModelCatalog>; disabled: boolean; selectionDisabled: boolean; value: string; onChange(value: string): void }): React.JSX.Element {
  const modelId = useId();
  const [custom, setCustom] = useState(false);
  const [customDraft, setCustomDraft] = useState(value);
  const models = catalog.models.length ? catalog.models : AI_MODELS.map(id => ({ id, name: id }));
  return <div className="model-catalog">
    <div className="model-choice"><label htmlFor={modelId}>Model</label><select id={modelId} value={custom ? "" : value} disabled={selectionDisabled} onChange={event => {
      if (event.target.value === "") { setCustom(true); onChange(customDraft); }
      else { setCustom(false); onChange(event.target.value); }
    }}>
      {!models.some(model => model.id === value) && value && <option value={value}>{value} (saved model)</option>}
      {models.map(model => <option key={model.id} value={model.id}>{model.name}</option>)}
      <option value="">Custom model ID…</option>
    </select></div>
    {custom && <label>Custom model ID<input value={customDraft} disabled={selectionDisabled} autoComplete="off" spellCheck={false} onChange={event => { setCustomDraft(event.target.value); onChange(event.target.value); }}/></label>}
    <button type="button" disabled={disabled || catalog.loading} onClick={catalog.refresh}>Refresh models</button>
    {catalog.status && <p className="hint" role="status">{catalog.status}</p>}
    <p className="hint">Your conversation keeps its selected model.</p>
  </div>;
}
