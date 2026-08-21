import type { QueryResult } from "@query-farm/vgi-excel-core";
import type { AgentMessage, WorkbookWriteAction } from "./agent";
import { sanitizeConversation } from "./agent-history";

export type ToolEvent = { id: string; name: string; state: "writing" | "running" | "done" | "error"; detail?: string; startedAt?: number; elapsedMs?: number };
export type StagedWorkbookAction = WorkbookWriteAction & { status: "pending" | "writing" | "done" | "error"; detail?: string };
export type ChatMessage = { role: "user" | "assistant"; text: string; tools?: ToolEvent[]; workbookActions?: StagedWorkbookAction[]; streaming?: boolean; activity?: string; stopped?: boolean };

export interface AgentConversationDocument {
  id: string;
  name: string;
  model: string;
  draft: string;
  displayMessages: ChatMessage[];
  agentMessages: AgentMessage[];
  createdAt: number;
  updatedAt: number;
  /** Session-only: query result IDs are process-local and cannot survive reload. */
  staged?: QueryResult | null;
}

export interface AgentConversationState { version: 1; documents: AgentConversationDocument[]; activeId: string }
export interface ConversationStorage { getItem(key: string): string | null; setItem(key: string, value: string): void }

const STORAGE_PREFIX = "cupola.agent.conversations.v1";
const MAX_CONVERSATIONS = 20;
const MAX_DISPLAY_MESSAGES = 100;
const MAX_AGENT_MESSAGES = 80;
const MAX_TEXT_CHARS = 50_000;
const MAX_TOOL_DETAIL_CHARS = 20_000;
const MAX_TOOL_RESULT_CHARS = 30_000;
const MAX_STORAGE_CHARS = 2_500_000;

function id(): string { try { return crypto.randomUUID(); } catch { return `conversation-${Date.now()}-${Math.random().toString(36).slice(2)}`; } }
function key(scope: string): string { return `${STORAGE_PREFIX}::${encodeURIComponent(scope || "default")}`; }
function storage(): ConversationStorage | undefined { try { return typeof localStorage === "undefined" ? undefined : localStorage; } catch { return undefined; } }
function nextName(documents: AgentConversationDocument[]): string {
  let index = 1;
  const names = new Set(documents.map((value) => value.name.toLocaleLowerCase()));
  while (names.has(`conversation ${index}`)) index++;
  return `Conversation ${index}`;
}
function document(name: string, model: string): AgentConversationDocument {
  const now = Date.now();
  return { id: id(), name, model, draft: "", displayMessages: [], agentMessages: [], createdAt: now, updatedAt: now, staged: null };
}
function fresh(model: string): AgentConversationState { const value = document("Conversation 1", model); return { version: 1, documents: [value], activeId: value.id }; }

export function loadAgentConversationState(scope: string, model: string, source: ConversationStorage | undefined = storage()): AgentConversationState {
  if (!source) return fresh(model);
  try {
    const raw = source.getItem(key(scope));
    if (!raw) return fresh(model);
    const parsed = JSON.parse(raw) as Partial<AgentConversationState>;
    const seen = new Set<string>();
    const documents = (Array.isArray(parsed.documents) ? parsed.documents : []).filter((value): value is AgentConversationDocument => {
      if (!value || typeof value.id !== "string" || seen.has(value.id) || !Array.isArray(value.displayMessages) || !Array.isArray(value.agentMessages)) return false;
      seen.add(value.id); return true;
    }).slice(-MAX_CONVERSATIONS).map((value, index) => sanitizeDocument({
      ...value,
      name: typeof value.name === "string" && value.name.trim() ? value.name.trim() : `Conversation ${index + 1}`,
      model: typeof value.model === "string" && value.model.trim() ? value.model : model,
      draft: typeof value.draft === "string" ? value.draft : "",
      createdAt: Number.isFinite(value.createdAt) ? value.createdAt : Date.now(),
      updatedAt: Number.isFinite(value.updatedAt) ? value.updatedAt : Date.now(),
      staged: null,
    }));
    if (!documents.length) return fresh(model);
    return { version: 1, documents, activeId: documents.some((value) => value.id === parsed.activeId) ? parsed.activeId! : documents[0].id };
  } catch { return fresh(model); }
}

export function saveAgentConversationState(scope: string, state: AgentConversationState, target: ConversationStorage | undefined = storage()): void {
  if (!target) return;
  try {
    const documents = compactDocuments(state.documents.slice(-MAX_CONVERSATIONS).map(sanitizeDocument), state.activeId);
    const activeId = documents.some((value) => value.id === state.activeId) ? state.activeId : documents[0]?.id;
    if (activeId) target.setItem(key(scope), JSON.stringify({ version: 1, documents, activeId }));
  } catch { }
}

export function addAgentConversation(state: AgentConversationState, model: string): AgentConversationState {
  const value = document(nextName(state.documents), model);
  const documents = [...state.documents, value].slice(-MAX_CONVERSATIONS);
  return { version: 1, documents, activeId: value.id };
}

export function removeAgentConversation(state: AgentConversationState, conversationId: string, model: string): AgentConversationState {
  const index = state.documents.findIndex((value) => value.id === conversationId);
  if (index < 0) return state;
  let documents = state.documents.filter((value) => value.id !== conversationId);
  if (!documents.length) { const value = document("Conversation 1", model); return { version: 1, documents: [value], activeId: value.id }; }
  const activeId = state.activeId === conversationId ? documents[Math.min(index, documents.length - 1)].id : state.activeId;
  return { version: 1, documents, activeId };
}

export function renameAgentConversation(state: AgentConversationState, conversationId: string, name: string): AgentConversationState {
  const trimmed = name.trim(); if (!trimmed) return state;
  return { ...state, documents: state.documents.map((value) => value.id === conversationId ? { ...value, name: trimmed.slice(0, 120), updatedAt: Date.now() } : value) };
}

export function activateAgentConversation(state: AgentConversationState, conversationId: string): AgentConversationState {
  return state.documents.some((value) => value.id === conversationId) ? { ...state, activeId: conversationId } : state;
}

export function titleFromPrompt(prompt: string): string {
  const value = prompt.trim().replace(/\s+/g, " ");
  return value.length > 44 ? `${value.slice(0, 43).trimEnd()}…` : value || "New conversation";
}

function sanitizeDocument(value: AgentConversationDocument): AgentConversationDocument {
  const displayMessages = value.displayMessages.slice(-MAX_DISPLAY_MESSAGES).map((message) => ({
    role: message.role,
    text: String(message.text ?? "").slice(0, MAX_TEXT_CHARS),
    tools: message.tools?.map((tool) => ({ id: tool.id, name: tool.name, state: tool.state === "writing" || tool.state === "running" ? "error" as const : tool.state, detail: tool.detail?.slice(0, MAX_TOOL_DETAIL_CHARS), elapsedMs: tool.elapsedMs })),
    stopped: message.stopped || message.streaming ? true : undefined,
  }));
  const agentMessages = JSON.parse(JSON.stringify(value.agentMessages.slice(-MAX_AGENT_MESSAGES))) as AgentMessage[];
  for (const message of agentMessages) {
    if (!Array.isArray(message.content)) continue;
    for (const block of message.content) {
      if (block.type !== "tool_result") continue;
      block.content = expireResultId(block.content).slice(0, MAX_TOOL_RESULT_CHARS);
    }
  }
  sanitizeAgentMessages(agentMessages);
  return { ...value, draft: String(value.draft ?? "").slice(0, 20_000), displayMessages, agentMessages, staged: undefined };
}

function sanitizeAgentMessages(messages: AgentMessage[]): void {
  sanitizeConversation(messages);
  while (messages[0]?.role === "assistant") messages.shift();
}

function compactDocuments(input: AgentConversationDocument[], activeId: string): AgentConversationDocument[] {
  const documents = input.map((value) => ({ ...value, displayMessages: [...value.displayMessages], agentMessages: [...value.agentMessages] }));
  const serialized = () => JSON.stringify({ version: 1, documents, activeId });
  let payload = serialized();
  while (payload.length > MAX_STORAGE_CHARS) {
    const candidate = [...documents].filter((value) => value.displayMessages.length > 2 || value.agentMessages.length > 2)
      .sort((left, right) => Number(left.id === activeId) - Number(right.id === activeId) || left.updatedAt - right.updatedAt)[0];
    if (!candidate) break;
    if (candidate.displayMessages.length > 2) candidate.displayMessages = candidate.displayMessages.slice(Math.min(10, candidate.displayMessages.length - 2));
    if (candidate.agentMessages.length > 2) candidate.agentMessages = candidate.agentMessages.slice(Math.min(10, candidate.agentMessages.length - 2));
    sanitizeAgentMessages(candidate.agentMessages);
    payload = serialized();
  }
  return documents;
}

function expireResultId(content: string): string {
  try {
    const value = JSON.parse(content) as Record<string, unknown>;
    if (value && typeof value === "object" && "result_id" in value) {
      delete value.result_id;
      value.result_expired = true;
      value.result_note = "The prior query result cache expired when the Cupola window closed. Run the SQL again before paging or staging it.";
      return JSON.stringify(value);
    }
  } catch { }
  return content;
}
