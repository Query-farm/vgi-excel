import { describe, expect, it } from "vitest";
import { activateOfficeAgentConversation, addOfficeAgentConversation, loadOfficeAgentConversationState, removeOfficeAgentConversation, renameOfficeAgentConversation, saveOfficeAgentConversationState, type OfficeConversationStorage } from "./agent-conversations";

class MemoryStorage implements OfficeConversationStorage {
  readonly values = new Map<string, string>();
  getItem(key: string): string | null { return this.values.get(key) ?? null; }
  setItem(key: string, value: string): void { this.values.set(key, value); }
}

describe("Microsoft 365 persisted AI conversations", () => {
  it("creates, renames, switches, closes, and restores conversation tabs per connection", () => {
    const storage = new MemoryStorage();
    let state = loadOfficeAgentConversationState("weather", "model", storage);
    const first = state.activeId;
    state.documents[0].displayMessages = [{ role: "user", text: "First question" }, { role: "assistant", text: "First answer" }];
    state.documents[0].agentMessages = [{ role: "user", content: "First question" }, { role: "assistant", content: [{ type: "text", text: "First answer" }] }];
    state = renameOfficeAgentConversation(state, first, "Close review");
    state = addOfficeAgentConversation(state, "model");
    const second = state.activeId;
    state.documents.find((value) => value.id === second)!.draft = "Second draft";
    saveOfficeAgentConversationState("weather", state, storage);

    const restored = loadOfficeAgentConversationState("weather", "other", storage);
    expect(restored.activeId).toBe(second);
    expect(restored.documents.map((value) => value.name)).toEqual(["Close review", "Conversation 1"]);
    expect(restored.documents[0].agentMessages).toHaveLength(2);
    expect(restored.documents[1].draft).toBe("Second draft");
    expect(loadOfficeAgentConversationState("ledger", "model", storage).documents).toHaveLength(1);
    expect(removeOfficeAgentConversation(activateOfficeAgentConversation(restored, first), first, "model").activeId).toBe(second);
  });

  it("does not persist API keys, staged results, or streaming state", () => {
    const storage = new MemoryStorage();
    const state = loadOfficeAgentConversationState("weather", "model", storage);
    state.documents[0].displayMessages = [{ role: "assistant", text: "Done", streaming: true }];
    state.documents[0].staged = { columns: [{ name: "amount", type: "DECIMAL(18,2)" }], rows: [[99.99]], rowCount: 1 };
    saveOfficeAgentConversationState("weather", state, storage);
    const payload = [...storage.values.values()].join("");
    expect(payload).not.toContain("api-key");
    expect(payload).not.toContain("99.99");
    const restored = loadOfficeAgentConversationState("weather", "model", storage);
    expect(restored.documents[0].staged).toBeUndefined();
    expect(restored.documents[0].displayMessages[0]).toMatchObject({ stopped: true });
    expect(restored.documents[0].displayMessages[0].streaming).toBeUndefined();
  });
});
