import { expect, it } from "vitest";
import { appendTranscript, savedTranscript, type AgentTranscriptPart } from "./agent-experience.js";

it("retains text and tool order without duplicating tool updates", () => {
  let parts: AgentTranscriptPart[] = [];
  for (const part of [
    { type: "thinking", text: "Check " }, { type: "thinking", text: "the scope." },
    { type: "text", text: "First query." }, { type: "tool", id: "first" },
    { type: "tool", id: "first" }, { type: "text", text: "Next query." },
    { type: "tool", id: "second" }, { type: "text", text: "Complete." },
  ] as AgentTranscriptPart[]) parts = appendTranscript(parts, part);
  expect(parts).toEqual([
    { type: "thinking", text: "Check the scope." }, { type: "text", text: "First query." },
    { type: "tool", id: "first" }, { type: "text", text: "Next query." },
    { type: "tool", id: "second" }, { type: "text", text: "Complete." },
  ]);
});

it("saves transcript presentation without signatures, redacted blocks or result rows", () => {
  expect(savedTranscript([
    { type: "thinking", text: "Provider summary", signature: "private-signature" },
    { type: "redacted_thinking", data: "private-redacted" },
    { type: "tool", id: "query", rows: [["private-row"]] },
    { type: "result", id: "process-local-result" },
  ])).toEqual([{ type: "thinking", text: "Provider summary" }, { type: "tool", id: "query" }]);
  expect(savedTranscript(undefined)).toBeUndefined();
});
