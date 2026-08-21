import { describe, expect, it } from "vitest";
import { sanitizeConversation } from "./agent-history";
import { recordToolCall } from "./agent-loop-guard";
import { parseStreamedToolInput } from "./tool-input";
import type { AgentMessage } from "./agent";

describe("Cupola-derived agent safeguards", () => {
  it("accepts empty tool input and reports malformed nonempty input", () => {
    expect(parseStreamedToolInput("  ")).toEqual({ input: {}, parseError: false });
    expect(parseStreamedToolInput("{bad")).toMatchObject({ input: { __parseError: "{bad" }, parseError: true });
  });

  it("repairs dangling tool calls and adjacent same-role messages", () => {
    const messages: AgentMessage[] = [
      { role: "user", content: "first" },
      { role: "assistant", content: [{ type: "tool_use", id: "dangling", name: "list_tables", input: {} }] },
      { role: "user", content: "second" },
      { role: "user", content: "third" },
    ];
    sanitizeConversation(messages);
    expect(JSON.stringify(messages)).not.toContain("dangling");
    expect(messages.map((item) => item.role)).toEqual(["user", "assistant", "user"]);
    expect(JSON.stringify(messages[2].content)).toContain("second");
    expect(JSON.stringify(messages[2].content)).toContain("third");
  });

  it("blocks a third identical deterministic metadata call regardless of key order", () => {
    const counts = new Map<string, number>();
    expect(recordToolCall(counts, "list_functions", { name: "forecast", catalog: "open_meteo" }).block).toBe(false);
    expect(recordToolCall(counts, "list_functions", { catalog: "open_meteo", name: "forecast" }).block).toBe(false);
    expect(recordToolCall(counts, "list_functions", { name: "forecast", catalog: "open_meteo" }).block).toBe(true);
    expect(recordToolCall(counts, "run_sql", { sql: "SELECT 1" }).block).toBe(false);
  });
});
