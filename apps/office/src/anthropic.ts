import { recordAITool } from "@query-farm/vgi-excel-core";
import { recordToolCall, repeatedCallMessage, clarificationInput, resultScope, type AgentClarification, type AgentResultCard } from "@query-farm/vgi-excel-core";
import { AGENT_TOOLS, agentRoundSeparator, appendAgentRoundText, buildExcelAgentSystemPrompt, executeAgentTool, type AgentCatalogObject, type AgentConnectionContext, type AgentToolContext, type QueryResult } from "@query-farm/vgi-excel-core";

import { DEFAULT_AI_MODEL, requestAnthropic, repairAgentHistory, type AIRequestOptions, type AssistantBlock, type AgentMessage, type ToolResultBlock } from "@query-farm/vgi-excel-core";
export type OfficeContentBlock = AssistantBlock;
type ContentBlock = OfficeContentBlock;
export type OfficeAgentMessage = AgentMessage;
type Message = OfficeAgentMessage;

export interface AgentAnswer {
  text: string;
  stagedResult?: QueryResult;
}

export interface AgentOptions extends AIRequestOptions {
  endpoint?: string;
  model?: string;
  fetchImpl?: typeof fetch;
  connection?: AgentConnectionContext;
}

export interface AgentCallbacks {
  onThinking?(text: string): void;
  onStage?(stage: string): void;
  onClarification?(question: AgentClarification, signal?: AbortSignal): Promise<string>;
  onQueryResult?(value: AgentResultCard): void;
  onText?(chunk: string): void;
  onTool?(name: string, state: "writing" | "running" | "done" | "error", detail?: string, id?: string): void;
  onResult?(result: QueryResult): void;
  onRetry?(message: string | null): void;
}

export class OfficeAgentSession {
  private readonly messages: Message[] = [];
  private readonly prompts = new Map<string, string>();

  reset(): void { this.messages.length = 0; this.prompts.clear(); }
  restore(messages: OfficeAgentMessage[]): void { this.prompts.clear(); this.messages.splice(0, this.messages.length, ...JSON.parse(JSON.stringify(messages))); repairAgentHistory(this.messages); }
  snapshot(): OfficeAgentMessage[] { const value = JSON.parse(JSON.stringify(this.messages)) as OfficeAgentMessage[]; repairAgentHistory(value); return value; }

  async run(apiKey: string, prompt: string, context: AgentToolContext, signal?: AbortSignal, options: AgentOptions = {}, callbacks: AgentCallbacks = {}): Promise<AgentAnswer> {
    if (!apiKey.trim()) throw new Error("An Anthropic API key is required.");
    if (!prompt.trim()) throw new Error("Enter a question for the agent.");
    signal?.throwIfAborted();
    this.messages.push({ role: "user", content: prompt });
    repairAgentHistory(this.messages);
    const repeated = new Map<string, number>();
    let sqlFailures = 0;
    callbacks.onStage?.("Exploring your catalog…");
    const promptKey = JSON.stringify(options.connection ?? {});
    let systemPrompt = this.prompts.get(promptKey);
    if (!systemPrompt) {
      const inventoryStarted = performance.now();
      try { systemPrompt = await loadSystemPrompt(context, options.connection, signal); signal?.throwIfAborted(); this.prompts.set(promptKey, systemPrompt); recordAITool("list_tables", performance.now() - inventoryStarted, "complete"); }
      catch (error) { recordAITool("list_tables", performance.now() - inventoryStarted, signal?.aborted ? "stopped" : "error"); throw error; }
    }
    let answer = "";
    let stagedResult: QueryResult | undefined;
    for (let round = 0; round < 20; round++) {
      let firstTextChunk = true;
      const roundCallbacks: AgentCallbacks = { ...callbacks, onText: (chunk) => {
        if (firstTextChunk && chunk) { firstTextChunk = false; const separator = agentRoundSeparator(answer, chunk); if (separator) callbacks.onText?.(separator); }
        callbacks.onText?.(chunk);
      } };
      callbacks.onStage?.("Thinking…");
      const response = await requestAnthropic(apiKey, options.model ?? DEFAULT_AI_MODEL, systemPrompt, AGENT_TOOLS, this.messages, roundCallbacks, signal, options);
      this.messages.push({ role: "assistant", content: response.content });
      const text = response.content.filter((block): block is Extract<ContentBlock, { type: "text" }> => block.type === "text").map((block) => block.text).join("");
      answer = appendAgentRoundText(answer, text);
      const calls = response.content.filter((block): block is Extract<ContentBlock, { type: "tool_use" }> => block.type === "tool_use");
      if (!calls.length) return { text: answer, stagedResult };
      const results: ToolResultBlock[] = [];
      this.messages.push({ role: "user", content: results });
      for (const call of calls) {
        signal?.throwIfAborted();
        callbacks.onTool?.(call.name, "running", summarizeInput(call.input), call.id);
        const started = performance.now();
        try {
          if (typeof call.input.__parseError === "string") throw new Error("The tool arguments were not valid JSON. Retry with a valid JSON object.");
          const repeat = recordToolCall(repeated, call.name, call.input);
          if (repeat.block || (call.name === "run_sql" && sqlFailures >= 3)) {
            const blocked = repeat.block ? repeatedCallMessage(call.name, repeat.count) : "The three-query failure budget for this turn has been reached. Explain what is missing instead of running another query.";
            callbacks.onTool?.(call.name, "error", blocked, call.id);
            results.push({ type: "tool_result", tool_use_id: call.id, content: blocked, is_error: true });
            continue;
          }
          const result = call.name === "ask_clarification" ? { content: JSON.stringify({ answer: await (callbacks.onClarification ? callbacks.onClarification(clarificationInput(call.input), signal) : Promise.reject(new Error("Ask your question in the answer and wait for the user's next message."))) }), queryResult: undefined } : await executeAgentTool(call.name, call.input, context, signal);
          signal?.throwIfAborted();
          recordAITool(call.name, performance.now() - started, "complete");
          if (result.queryResult) { stagedResult = result.queryResult; callbacks.onResult?.(result.queryResult); callbacks.onQueryResult?.({ toolCallId: call.id, id: call.id, sql: String(call.input.sql), result: result.queryResult, connection: options.connection?.name ?? "", catalog: options.connection?.catalog ?? "", scope: resultScope(call.input.scope) }); }
          callbacks.onTool?.(call.name, "done", call.name === "run_sql" ? String(call.input.sql) : `${Math.round(performance.now() - started)} ms`, call.id);
          results.push({ type: "tool_result", tool_use_id: call.id, content: result.content });
        } catch (error) {
          recordAITool(call.name, performance.now() - started, signal?.aborted || (error as Error).name === "AbortError" ? "stopped" : "error");
          if (signal?.aborted || (error as Error).name === "AbortError") { repairAgentHistory(this.messages); throw error; }
          if (call.name === "run_sql") sqlFailures++;
          callbacks.onTool?.(call.name, "error", message(error), call.id);
          results.push({ type: "tool_result", tool_use_id: call.id, is_error: true, content: message(error) + (call.name === "run_sql" ? ` Query failure ${sqlFailures} of 3. Inspect metadata before retrying; stop querying after three failures.` : "") });
        }
      }
    }
    throw new Error("The agent exceeded the maximum number of tool rounds.");
  }
}

export async function runAgent(
  apiKey: string,
  prompt: string,
  context: AgentToolContext,
  signal?: AbortSignal,
  options: AgentOptions = {},
): Promise<AgentAnswer> {
  return new OfficeAgentSession().run(apiKey, prompt, context, signal, options);
}

async function loadSystemPrompt(context: AgentToolContext, connection?: AgentConnectionContext, signal?: AbortSignal): Promise<string> {
  const active = connection ?? { name: "selected VGI connection", catalog: "selected catalog", authentication: "unspecified" };
  const objects: AgentCatalogObject[] = [];
  const errors: string[] = [];
  const [tables, functions] = await Promise.allSettled([context.listTables(signal), context.listFunctions(signal)]);
  signal?.throwIfAborted();
  if (tables.status === "fulfilled") {
    for (const row of records(tables.value)) objects.push({
      catalog: String(row.table_catalog ?? active.catalog), schema: String(row.table_schema ?? "main"),
      name: String(row.table_name ?? ""), kind: String(row.table_type ?? "table").toLowerCase(),
    });
  } else errors.push(message(tables.reason));
  if (functions.status === "fulfilled") {
    for (const row of records(functions.value)) objects.push({
      catalog: String(row.catalog ?? active.catalog), schema: String(row.schema ?? "main"), name: String(row.name ?? ""),
      kind: String(row.kind ?? "function"), description: row.description == null ? undefined : String(row.description),
    });
  } else errors.push(message(functions.reason));
  return buildExcelAgentSystemPrompt({ connection: active, objects: objects.filter((item) => item.name), inventoryError: errors.join("; ") || undefined });
}

function records(value: unknown): Record<string, unknown>[] {
  return Array.isArray(value) ? value.filter((item): item is Record<string, unknown> => !!item && typeof item === "object" && !Array.isArray(item)) : [];
}

function message(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function summarizeInput(input: Record<string, unknown>): string {
  if (typeof input.sql === "string") return input.sql;
  const value = JSON.stringify(input); return value.length > 500 ? value.slice(0, 500) + "…" : value;
}
