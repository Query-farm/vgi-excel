import { describe, expect, it } from "vitest";
import { addAgentConversation, loadAgentConversationState, removeAgentConversation, renameAgentConversation, saveAgentConversationState, titleFromPrompt, type ConversationStorage } from "./agent-conversations";

class MemoryStorage implements ConversationStorage {
  values = new Map<string, string>();
  getItem(key: string): string | null { return this.values.get(key) ?? null; }
  setItem(key: string, value: string): void { this.values.set(key, value); }
}

describe("persisted AI conversations", () => {
  it("creates, renames, activates, closes, and restores conversation tabs per connection", () => {
    const storage = new MemoryStorage();
    let state = loadAgentConversationState("weather", "model-a", storage);
    state.documents[0].displayMessages = [{ role: "user", text: "Forecast for Richmond" }, { role: "assistant", text: "It will be warm." }];
    state.documents[0].agentMessages = [{ role: "user", content: "Forecast for Richmond" }, { role: "assistant", content: "It will be warm." }];
    state = addAgentConversation(state, "model-b");
    state = renameAgentConversation(state, state.activeId, "Rain analysis");
    saveAgentConversationState("weather", state, storage);

    const restored = loadAgentConversationState("weather", "fallback", storage);
    expect(restored.documents.map((value) => value.name)).toEqual(["Conversation 1", "Rain analysis"]);
    expect(restored.documents[0].displayMessages[1].text).toBe("It will be warm.");
    expect(restored.documents[1].model).toBe("model-b");
    expect(loadAgentConversationState("other", "fallback", storage).documents).toHaveLength(1);

    const remaining = removeAgentConversation(restored, restored.activeId, "fallback");
    expect(remaining.documents).toHaveLength(1);
    expect(remaining.activeId).toBe(remaining.documents[0].id);
  });

  it("expires process-local result IDs and excludes staged workbook data", () => {
    const storage = new MemoryStorage();
    const state = loadAgentConversationState("weather", "model", storage);
    state.documents[0].agentMessages = [
      { role: "assistant", content: [{ type: "tool_use", id: "query", name: "run_sql", input: { sql: "SELECT 1" } }] },
      { role: "user", content: [{ type: "tool_result", tool_use_id: "query", content: '{"result_id":"secret-process-id","rows":[[1]]}' }] },
    ];
    state.documents[0].staged = { columns: [{ name: "value", type: "INTEGER" }], rows: [[1]], rowCount: 1 };
    saveAgentConversationState("weather", state, storage);
    const restored = loadAgentConversationState("weather", "model", storage);
    expect(JSON.stringify(restored.documents[0].agentMessages)).not.toContain("secret-process-id");
    expect(JSON.stringify(restored.documents[0].agentMessages)).toContain("result_expired");
    expect(restored.documents[0].staged).toBeUndefined();
  });

  it("derives concise names from the first prompt", () => {
    expect(titleFromPrompt("  Current weather in New York  ")).toBe("Current weather in New York");
    expect(titleFromPrompt("A very long question that should become a concise conversation tab title for display")).toMatch(/…$/);
  });

  it("compacts old history before reaching browser storage quotas", () => {
    const storage = new MemoryStorage();
    let state = loadAgentConversationState("weather", "model", storage);
    for (let index = 0; index < 20; index++) {
      if (index) state = addAgentConversation(state, "model");
      const active = state.documents.find((value) => value.id === state.activeId)!;
      active.displayMessages = Array.from({ length: 40 }, (_value, message) => ({ role: message % 2 ? "assistant" as const : "user" as const, text: "x".repeat(10_000) }));
      active.agentMessages = Array.from({ length: 40 }, (_value, message) => ({ role: message % 2 ? "assistant" as const : "user" as const, content: "y".repeat(10_000) }));
    }
    saveAgentConversationState("weather", state, storage);
    const payload = [...storage.values.values()][0];
    expect(payload.length).toBeLessThanOrEqual(2_500_000);
    expect(loadAgentConversationState("weather", "model", storage).documents).toHaveLength(20);
  });
});
