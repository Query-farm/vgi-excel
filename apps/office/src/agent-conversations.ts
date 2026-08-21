import type { QueryResult } from "@query-farm/vgi-excel-core";
import type { OfficeAgentMessage } from "./anthropic";

export type OfficeToolEvent = { id: string; name: string; state: "writing" | "running" | "done" | "error"; detail?: string };
export type OfficeChatMessage = { role: "user" | "assistant"; text: string; tools?: OfficeToolEvent[]; streaming?: boolean; activity?: string; stopped?: boolean };

export interface OfficeAgentConversation {
  id: string;
  name: string;
  model: string;
  draft: string;
  displayMessages: OfficeChatMessage[];
  agentMessages: OfficeAgentMessage[];
  createdAt: number;
  updatedAt: number;
  /** Results and workbook outcomes are session-only. */
  staged?: QueryResult | null;
  outcome?: { sheet: string; table: string; address: string; rows: number } | null;
}

export interface OfficeAgentConversationState { version: 1; documents: OfficeAgentConversation[]; activeId: string }
export interface OfficeConversationStorage { getItem(key: string): string | null; setItem(key: string, value: string): void }

const STORAGE_PREFIX = "cupola.office.agent.conversations.v1";
const MAX_CONVERSATIONS = 20;
const MAX_DISPLAY_MESSAGES = 100;
const MAX_AGENT_MESSAGES = 80;
const MAX_TEXT_CHARS = 50_000;
const MAX_TOOL_DETAIL_CHARS = 20_000;
const MAX_TOOL_RESULT_CHARS = 30_000;
const MAX_STORAGE_CHARS = 2_500_000;

function id(): string { try { return crypto.randomUUID(); } catch { return `conversation-${Date.now()}-${Math.random().toString(36).slice(2)}`; } }
function key(scope: string): string { return `${STORAGE_PREFIX}::${encodeURIComponent(scope || "default")}`; }
function storage(): OfficeConversationStorage | undefined { try { return typeof localStorage === "undefined" ? undefined : localStorage; } catch { return undefined; } }
function nextName(documents: OfficeAgentConversation[]): string { let index = 1; const names = new Set(documents.map((value) => value.name.toLowerCase())); while (names.has(`conversation ${index}`)) index++; return `Conversation ${index}`; }
function document(name: string, model: string): OfficeAgentConversation { const now = Date.now(); return { id: id(), name, model, draft: "", displayMessages: [], agentMessages: [], createdAt: now, updatedAt: now, staged: null, outcome: null }; }
function fresh(model: string): OfficeAgentConversationState { const value = document("Conversation 1", model); return { version: 1, documents: [value], activeId: value.id }; }

export function loadOfficeAgentConversationState(scope: string, model: string, source: OfficeConversationStorage | undefined = storage()): OfficeAgentConversationState {
  if (!source) return fresh(model);
  try {
    const parsed = JSON.parse(source.getItem(key(scope)) ?? "null") as Partial<OfficeAgentConversationState> | null;
    if (!parsed || !Array.isArray(parsed.documents)) return fresh(model);
    const ids = new Set<string>();
    const documents = parsed.documents.filter((value): value is OfficeAgentConversation => !!value && typeof value.id === "string" && !ids.has(value.id) && (ids.add(value.id), true)).slice(-MAX_CONVERSATIONS).map((value, index) => sanitize({
      ...value,
      name: typeof value.name === "string" && value.name.trim() ? value.name.trim() : `Conversation ${index + 1}`,
      model: typeof value.model === "string" && value.model.trim() ? value.model : model,
      draft: typeof value.draft === "string" ? value.draft : "",
      createdAt: Number.isFinite(value.createdAt) ? value.createdAt : Date.now(),
      updatedAt: Number.isFinite(value.updatedAt) ? value.updatedAt : Date.now(),
      staged: null,
      outcome: null,
    }));
    if (!documents.length) return fresh(model);
    return { version: 1, documents, activeId: documents.some((value) => value.id === parsed.activeId) ? parsed.activeId! : documents[0].id };
  } catch { return fresh(model); }
}

export function saveOfficeAgentConversationState(scope: string, state: OfficeAgentConversationState, target: OfficeConversationStorage | undefined = storage()): void {
  if (!target) return;
  try {
    const documents = compact(state.documents.slice(-MAX_CONVERSATIONS).map(sanitize), state.activeId);
    const activeId = documents.some((value) => value.id === state.activeId) ? state.activeId : documents[0]?.id;
    if (activeId) target.setItem(key(scope), JSON.stringify({ version: 1, documents, activeId }));
  } catch { }
}

export function addOfficeAgentConversation(state: OfficeAgentConversationState, model: string): OfficeAgentConversationState { const value = document(nextName(state.documents), model); return { version: 1, documents: [...state.documents, value].slice(-MAX_CONVERSATIONS), activeId: value.id }; }
export function activateOfficeAgentConversation(state: OfficeAgentConversationState, conversationId: string): OfficeAgentConversationState { return state.documents.some((value) => value.id === conversationId) ? { ...state, activeId: conversationId } : state; }
export function renameOfficeAgentConversation(state: OfficeAgentConversationState, conversationId: string, name: string): OfficeAgentConversationState { const trimmed = name.trim(); return trimmed ? { ...state, documents: state.documents.map((value) => value.id === conversationId ? { ...value, name: trimmed.slice(0, 120), updatedAt: Date.now() } : value) } : state; }
export function removeOfficeAgentConversation(state: OfficeAgentConversationState, conversationId: string, model: string): OfficeAgentConversationState {
  const index = state.documents.findIndex((value) => value.id === conversationId); if (index < 0) return state;
  const documents = state.documents.filter((value) => value.id !== conversationId);
  if (!documents.length) return fresh(model);
  return { version: 1, documents, activeId: state.activeId === conversationId ? documents[Math.min(index, documents.length - 1)].id : state.activeId };
}
export function officeConversationTitle(prompt: string): string { const value = prompt.trim().replace(/\s+/g, " "); return value.length > 44 ? `${value.slice(0, 43).trimEnd()}…` : value || "New conversation"; }

function sanitize(value: OfficeAgentConversation): OfficeAgentConversation {
  const displayMessages = (Array.isArray(value.displayMessages) ? value.displayMessages : []).slice(-MAX_DISPLAY_MESSAGES).map((message) => ({
    role: message.role === "user" ? "user" as const : "assistant" as const,
    text: String(message.text ?? "").slice(0, MAX_TEXT_CHARS),
    tools: message.tools?.map((tool) => ({ id: String(tool.id), name: String(tool.name), state: tool.state === "writing" || tool.state === "running" ? "error" as const : tool.state, detail: tool.detail?.slice(0, MAX_TOOL_DETAIL_CHARS) })),
    stopped: message.stopped || message.streaming ? true : undefined,
  }));
  const agentMessages = JSON.parse(JSON.stringify((Array.isArray(value.agentMessages) ? value.agentMessages : []).slice(-MAX_AGENT_MESSAGES))) as OfficeAgentMessage[];
  for (const message of agentMessages) if (Array.isArray(message.content)) for (const block of message.content as Array<{ type?: string; content?: unknown }>) if (block.type === "tool_result" && typeof block.content === "string") block.content = block.content.slice(0, MAX_TOOL_RESULT_CHARS);
  while (agentMessages[0]?.role === "assistant") agentMessages.shift();
  return { ...value, draft: String(value.draft ?? "").slice(0, 20_000), displayMessages, agentMessages, staged: undefined, outcome: undefined };
}

function compact(input: OfficeAgentConversation[], activeId: string): OfficeAgentConversation[] {
  const documents = input.map((value) => ({ ...value, displayMessages: [...value.displayMessages], agentMessages: [...value.agentMessages] }));
  const serialized = () => JSON.stringify({ version: 1, documents, activeId });
  let payload = serialized();
  while (payload.length > MAX_STORAGE_CHARS) {
    const candidate = [...documents].filter((value) => value.displayMessages.length > 2 || value.agentMessages.length > 2).sort((left, right) => Number(left.id === activeId) - Number(right.id === activeId) || left.updatedAt - right.updatedAt)[0];
    if (!candidate) break;
    if (candidate.displayMessages.length > 2) candidate.displayMessages = candidate.displayMessages.slice(Math.min(10, candidate.displayMessages.length - 2));
    if (candidate.agentMessages.length > 2) candidate.agentMessages = candidate.agentMessages.slice(Math.min(10, candidate.agentMessages.length - 2));
    while (candidate.agentMessages[0]?.role === "assistant") candidate.agentMessages.shift();
    payload = serialized();
  }
  return documents;
}
