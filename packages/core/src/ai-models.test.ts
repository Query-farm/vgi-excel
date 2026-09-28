import { expect, it, vi } from "vitest";
import { listAIModels, modelCacheScope, readModelCache, saveModelCache } from "./ai-models.js";
import { anthropicRequestBody } from "./anthropic.js";

const metadata = { id: "future-model", display_name: "Future model", max_tokens: 50000, capabilities: { thinking: { types: { adaptive: { supported: true } } }, effort: { supported: true, low: { supported: true }, high: { supported: true } } } };
it("discovers paginated models and applies their capabilities to requests", async () => {
  const fetchImpl = vi.fn().mockResolvedValueOnce(new Response(JSON.stringify({ data: [metadata], has_more: true, last_id: "future-model" }))).mockResolvedValueOnce(new Response(JSON.stringify({ data: [{ id: "basic", capabilities: { effort: { supported: false } } }], has_more: false })));
  const models = await listAIModels("key", "workspace", { fetchImpl });
  expect(models).toHaveLength(2);
  expect(fetchImpl.mock.calls[1][0]).toContain("after_id=future-model");
  expect(fetchImpl.mock.calls[0][1].headers["anthropic-workspace-id"]).toBe("workspace");
  const body = anthropicRequestBody("future-model", "", [], [], { modelInfo: models[0], maxTokens: 100000, effort: "max" });
  expect(body).toMatchObject({ max_tokens: 50000, thinking: { type: "adaptive" }, output_config: { effort: "high" } });
  expect(anthropicRequestBody("basic", "", [], [], { modelInfo: models[1] })).not.toHaveProperty("thinking");
});
it("stores metadata without credentials and isolates keys and workspaces", async () => {
  const values = new Map<string, string>();
  const storage = { getItem: (key: string) => values.get(key) ?? null, setItem: (key: string, value: string) => values.set(key, value) };
  const scope = await modelCacheScope("secret-key", "private-workspace");
  saveModelCache(storage, { scope, updatedAt: 42, models: [{ id: "future-model", name: "Future" }] });
  expect(readModelCache(storage, scope)?.updatedAt).toBe(42);
  expect(readModelCache(storage, await modelCacheScope("other-key", "private-workspace"))).toBeNull();
  expect(readModelCache(storage, await modelCacheScope("secret-key", "other-workspace"))).toBeNull();
  expect([...values.values()].join()).not.toMatch(/secret-key|private-workspace/);
});
it("bounds stalled discovery and never exposes provider error bodies", async () => {
  const fetchImpl = vi.fn((_url, init) => new Promise<Response>((_, reject) => init.signal.addEventListener("abort", () => reject(new DOMException("secret-key", "AbortError")))));
  await expect(listAIModels("key", "", { fetchImpl: fetchImpl as typeof fetch, timeoutMs: 10 })).rejects.toThrow("timed out");
  await expect(listAIModels("key", "", { fetchImpl: async () => new Response("secret-key SELECT private_data https://customer.example", { status: 401 }) })).rejects.toThrow("Check your API key");
});
it("does not return partial data if pagination is invalid", async () => {
  await expect(listAIModels("key", "", { fetchImpl: async () => new Response(JSON.stringify({ data: [metadata], has_more: true })) })).rejects.toThrow("incomplete");
});
