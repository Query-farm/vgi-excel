import type { AgentBlock, AgentMessage, ToolResultBlock } from "./agent";

type AnyBlock = AgentBlock | ToolResultBlock;

function blocks(content: AgentMessage["content"]): AnyBlock[] {
  return typeof content === "string" ? (content ? [{ type: "text", text: content }] : []) : content;
}

export function sanitizeConversation(messages: AgentMessage[]): void {
  for (let index = 0; index < messages.length; index++) {
    const message = messages[index];
    if (message.role !== "assistant" || !Array.isArray(message.content)) continue;
    const assistant = message.content as AgentBlock[];
    if (!assistant.some((item) => item.type === "tool_use")) continue;
    const next = messages[index + 1];
    const answered = new Set<string>();
    if (next?.role === "user" && Array.isArray(next.content)) {
      for (const item of next.content as ToolResultBlock[]) if (item.type === "tool_result") answered.add(item.tool_use_id);
    }
    const kept = assistant.filter((item) => item.type !== "tool_use" || answered.has(item.id));
    if (kept.length !== assistant.length) message.content = kept.length ? kept : [{ type: "text", text: "(stopped)" }];
  }
  for (let index = 1; index < messages.length; index++) {
    const current = messages[index], previous = messages[index - 1];
    if (current.role !== previous.role) continue;
    const merged = [...blocks(previous.content), ...blocks(current.content)];
    previous.content = merged.length ? merged : [{ type: "text", text: "(empty)" }];
    messages.splice(index, 1); index--;
  }
}
