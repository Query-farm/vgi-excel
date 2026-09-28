export interface AIModelInfo { id: string; name: string; maxTokens?: number; adaptive?: boolean; efforts?: string[] }
export interface AIModelCache { scope: string; updatedAt: number; models: AIModelInfo[] }
export const MODEL_CACHE_KEY = "cupola.ai.models.v1";
export const MODEL_CACHE_AGE = 24 * 60 * 60 * 1000;
class ModelDiscoveryError extends Error {}
const efforts = ["low", "medium", "high", "xhigh", "max"];

export async function modelCacheScope(key: string, workspace = ""): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(JSON.stringify([key.trim(), workspace.trim()])));
  return Array.from(new Uint8Array(digest), b => b.toString(16).padStart(2, "0")).join("");
}
export function readModelCache(storage: Pick<Storage, "getItem">, scope: string): AIModelCache | null {
  try {
    const value = JSON.parse(storage.getItem(MODEL_CACHE_KEY) ?? "null");
    if (value?.scope !== scope || !Number.isFinite(value.updatedAt) || !Array.isArray(value.models) || !value.models.every((m: AIModelInfo) => typeof m.id === "string" && typeof m.name === "string" && (m.maxTokens === undefined || (Number.isSafeInteger(m.maxTokens) && m.maxTokens > 0)) && (m.adaptive === undefined || typeof m.adaptive === "boolean") && (m.efforts === undefined || (Array.isArray(m.efforts) && m.efforts.every(level => efforts.includes(level)))))) return null;
    return value;
  } catch { return null; }
}
export function saveModelCache(storage: Pick<Storage, "setItem">, value: AIModelCache): void {
  try { storage.setItem(MODEL_CACHE_KEY, JSON.stringify(value)); } catch { /* Discovery still works if local storage is unavailable. */ }
}

export async function listAIModels(key: string, workspace = "", options: { signal?: AbortSignal; fetchImpl?: typeof fetch; timeoutMs?: number } = {}): Promise<AIModelInfo[]> {
  if (!key.trim()) throw new ModelDiscoveryError("Enter an API key to update the model list.");
  const controller = new AbortController();
  const abort = () => controller.abort();
  options.signal?.addEventListener("abort", abort, { once: true });
  if (options.signal?.aborted) abort();
  const timer = setTimeout(abort, options.timeoutMs ?? 15_000);
  const models = new Map<string, AIModelInfo>();
  const cursors = new Set<string>();
  let cursor = "";
  try {
    for (let page = 0; page < 20; page++) {
      controller.signal.throwIfAborted();
      const url = new URL("https://api.anthropic.com/v1/models");
      url.searchParams.set("limit", "1000");
      if (cursor) url.searchParams.set("after_id", cursor);
      const response = await (options.fetchImpl ?? fetch)(url.toString(), { signal: controller.signal, headers: {
        "x-api-key": key.trim(), "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true",
        ...(workspace.trim() ? { "anthropic-workspace-id": workspace.trim() } : {}),
      } });
      if (!response.ok) throw new ModelDiscoveryError(response.status === 401 || response.status === 403 ? "Could not update models. Check your API key and workspace." : "Could not update models. Try again.");
      const value = await response.json();
      if (!Array.isArray(value.data)) throw new ModelDiscoveryError("The model list was invalid. Try again.");
      for (const item of value.data) {
        if (typeof item?.id !== "string" || !item.id.trim()) continue;
        const capabilities = item.capabilities;
        models.set(item.id, { id: item.id, name: typeof item.display_name === "string" ? item.display_name : item.id,
          ...(Number.isInteger(item.max_tokens) && item.max_tokens > 0 ? { maxTokens: item.max_tokens } : {}),
          ...(capabilities ? { adaptive: capabilities.thinking?.types?.adaptive?.supported === true,
            efforts: capabilities.effort?.supported === true ? efforts.filter(level => capabilities.effort[level]?.supported === true) : [] } : {}),
        });
      }
      if (!value.has_more) {
        if (!models.size) throw new ModelDiscoveryError("No models were returned. You can enter a model ID manually.");
        return [...models.values()];
      }
      if (typeof value.last_id !== "string" || !value.last_id || cursors.has(value.last_id)) throw new ModelDiscoveryError("The model list was incomplete. Try again.");
      cursor = value.last_id; cursors.add(cursor);
    }
    throw new ModelDiscoveryError("The model list was incomplete. Try again.");
  } catch (error) {
    if (options.signal?.aborted) throw new DOMException("Cancelled", "AbortError");
    if (controller.signal.aborted) throw new ModelDiscoveryError("Updating models timed out. Try again.");
    // Never surface response bodies, request headers, or provider error payloads.
    if (error instanceof ModelDiscoveryError) throw error;
    throw new ModelDiscoveryError("Could not update models. Check your connection and try again.");
  } finally { clearTimeout(timer); options.signal?.removeEventListener("abort", abort); }
}
