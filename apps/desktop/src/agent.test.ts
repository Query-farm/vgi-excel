import { afterEach, describe, expect, it, vi } from "vitest";

const realWindow = (globalThis as any).window;
const realFetch = globalThis.fetch;

afterEach(() => {
  (globalThis as any).window = realWindow;
  globalThis.fetch = realFetch;
  vi.restoreAllMocks();
  vi.resetModules();
});

function sse(events: unknown[]): Response {
  const body = events.map((event) => `event: ${(event as any).type}\ndata: ${JSON.stringify(event)}\n\n`).join("");
  return new Response(body, { status: 200, headers: { "content-type": "text/event-stream" } });
}

function toolTurn(id = "tool-1", name = "run_sql", input = '{"sql":"SELECT 42 AS value"}') {
  return sse([
    { type: "content_block_start", content_block: { type: "tool_use", id, name } },
    { type: "content_block_delta", delta: { type: "input_json_delta", partial_json: input } },
    { type: "content_block_stop" },
    { type: "message_delta", delta: { stop_reason: "tool_use" } },
  ]);
}

function textTurn(text = "The answer is 42.") {
  return sse([
    { type: "content_block_start", content_block: { type: "text", text: "" } },
    { type: "content_block_delta", delta: { type: "text_delta", text } },
    { type: "content_block_stop" },
    { type: "message_delta", delta: { stop_reason: "end_turn" } },
  ]);
}

function textAndToolTurn(text: string, id = "tool-1", name = "run_sql", input = '{"sql":"SELECT 42 AS value"}') {
  return sse([
    { type: "content_block_start", content_block: { type: "text", text: "" } },
    { type: "content_block_delta", delta: { type: "text_delta", text } },
    { type: "content_block_stop" },
    { type: "content_block_start", content_block: { type: "tool_use", id, name } },
    { type: "content_block_delta", delta: { type: "input_json_delta", partial_json: input } },
    { type: "content_block_stop" },
    { type: "message_delta", delta: { stop_reason: "tool_use" } },
  ]);
}

function installHost(query: { result?: unknown; error?: string } | ((request: any) => { result?: unknown; error?: string })) {
  const listeners: Array<(event: MessageEvent) => void> = [];
  const posted: any[] = [];
  (globalThis as any).window = { chrome: { webview: {
    addEventListener: (_type: string, listener: (event: MessageEvent) => void) => listeners.push(listener),
    postMessage: (request: any) => {
      posted.push(request);
      const response = request.method === "agent.trace" ? { result: true } : typeof query === "function" ? query(request) : query;
      queueMicrotask(() => listeners[0]({ data: { id: request.id, ...response } } as MessageEvent));
    },
  } } };
  return posted;
}

function callbacks() {
  return { onText: vi.fn(), onResult: vi.fn(), onRetry: vi.fn(), onTool: vi.fn(), onWorkbookAction: vi.fn(), onQueryDocument: vi.fn() };
}

const CONNECTION = { name: "weather", catalog: "open_meteo", location: "https://weather.test/vgi", authentication: "anonymous" as const };

describe("desktop agent", () => {
  it("restores persisted model history before continuing a conversation", async () => {
    installHost({ result: { columns: [], rows: [], rowCount: 0 } });
    const requests: RequestInit[] = [];
    globalThis.fetch = vi.fn(async (_url, init) => { requests.push(init!); return textTurn("Continuing the saved conversation."); }) as unknown as typeof fetch;
    const { AgentSession } = await import("./agent");
    const session = new AgentSession();
    session.restore([{ role: "user", content: "Earlier question" }, { role: "assistant", content: "Earlier answer" }]);
    await session.run("key", "model", "Follow-up question", CONNECTION, callbacks(), new AbortController().signal);
    const body = JSON.parse(String(requests[0].body));
    expect(body.messages.slice(0, 3)).toEqual([
      { role: "user", content: "Earlier question" },
      { role: "assistant", content: "Earlier answer" },
      { role: "user", content: "Follow-up question" },
    ]);
    expect(session.snapshot()).toHaveLength(4);
  });

  it("sends the active connection and live catalog inventory in the system prompt", async () => {
    const inventory = {
      columns: ["catalog", "schema", "name", "kind", "description"].map((name) => ({ name, type: "VARCHAR" })),
      rows: [["open_meteo", "main", "forecast_current", "table", "Current forecast"]], rowCount: 1, truncated: false,
    };
    installHost({ result: inventory });
    const requests: RequestInit[] = [];
    globalThis.fetch = vi.fn(async (_url, init) => { requests.push(init!); return textTurn("Ready"); }) as unknown as typeof fetch;
    const { AgentSession } = await import("./agent");
    await new AgentSession().run("test-key", "model", "what is connected?", CONNECTION, callbacks(), new AbortController().signal);

    const body = JSON.parse(String(requests[0].body));
    expect(body.system).toContain("Connection name: weather");
    expect(body.system).toContain("Attached catalog: open_meteo");
    expect(body.system).toContain("`open_meteo.main.forecast_current`");
    expect(body.system).not.toContain(CONNECTION.location);
  });

  it("streams text, executes a native read-only query, and stages the result", async () => {
    const result = { columns: [{ name: "value", type: "INTEGER" }], rows: [[42]], rowCount: 1, truncated: false, elapsedMs: 3 };
    const posted = installHost({ result });
    const responses = [toolTurn(), textTurn()];
    globalThis.fetch = vi.fn(async () => responses.shift()!) as unknown as typeof fetch;
    const { AgentSession } = await import("./agent");
    const events = callbacks();
    await new AgentSession().run("test-key", "test-model", "Find it", CONNECTION, events, new AbortController().signal);

    expect(events.onText.mock.calls.flat().join("")).toBe("The answer is 42.");
    expect(events.onResult).toHaveBeenCalledWith(result);
    expect(posted.find((item) => item.method === "query.run" && item.params.sql === "SELECT 42 AS value")).toMatchObject({ params: { agent: true, connection: "weather", sql: "SELECT 42 AS value" } });
    expect(events.onTool).toHaveBeenCalledWith("run_sql", "writing", undefined, "tool-1");
    expect(events.onTool).toHaveBeenCalledWith("run_sql", "done", "SELECT 42 AS value", "tool-1");
    expect(events.onTool.mock.calls.flat().join(" ")).not.toContain("result_id");
  });

  it("separates prose streamed before and after a tool call into paragraphs", async () => {
    installHost({ result: { columns: [{ name: "value", type: "INTEGER" }], rows: [[42]], rowCount: 1 } });
    const responses = [textAndToolTurn("Let me check."), textTurn("Got the result.")];
    globalThis.fetch = vi.fn(async () => responses.shift()!) as unknown as typeof fetch;
    const { AgentSession } = await import("./agent");
    const events = callbacks();
    await new AgentSession().run("test-key", "test-model", "Find it", CONNECTION, events, new AbortController().signal);
    expect(events.onText.mock.calls.flat().join("")).toBe("Let me check.\n\nGot the result.");
  });

  it("inspects the workbook through the read-only native bridge", async () => {
    const posted = installHost((request) => request.method === "excel.workbookOverview"
      ? { result: { workbook: "Book1.xlsx", activeSheet: "Forecast", selection: "$A$1", worksheets: [] } }
      : { result: { columns: [], rows: [], rowCount: 0 } });
    const responses = [toolTurn("workbook", "workbook_overview", "{}"), textTurn("The active sheet is Forecast.")];
    globalThis.fetch = vi.fn(async () => responses.shift()!) as unknown as typeof fetch;
    const { AgentSession } = await import("./agent");
    await new AgentSession().run("key", "model", "Inspect my workbook", CONNECTION, callbacks(), new AbortController().signal);
    expect(posted.some((item) => item.method === "excel.workbookOverview")).toBe(true);
  });

  it("stages a new worksheet write without mutating Excel", async () => {
    const queryResult = { columns: [{ name: "temperature", type: "DOUBLE" }], rows: [[21.5]], rowCount: 1, truncated: false };
    const posted = installHost((request) => ({ result: request.method === "query.run" && request.params.sql === "SELECT 21.5 AS temperature" ? queryResult : { columns: [], rows: [], rowCount: 0 } }));
    const responses = [
      toolTurn("query", "run_sql", '{"sql":"SELECT 21.5 AS temperature"}'),
      toolTurn("stage", "stage_result_to_new_sheet", '{"result_id":"RESULT_ID","sheet_name":"Revenue / Expense: August [Final]","table_name":"VGI_Forecast"}'),
      textTurn("The Forecast worksheet is ready for your confirmation."),
    ];
    const events = callbacks();
    globalThis.fetch = vi.fn(async (_url, init) => {
      if (responses.length === 2) {
        const body = JSON.parse(String(init?.body));
        const resultId = JSON.parse(body.messages[2].content[0].content).result_id;
        responses[0] = toolTurn("stage", "stage_result_to_new_sheet", JSON.stringify({ result_id: resultId, sheet_name: "Revenue / Expense: August [Final]", table_name: "VGI_Forecast" }));
      }
      return responses.shift()!;
    }) as unknown as typeof fetch;
    const { AgentSession } = await import("./agent");
    await new AgentSession().run("key", "model", "Put it on a new sheet", CONNECTION, events, new AbortController().signal);
    expect(events.onWorkbookAction).toHaveBeenCalledWith(expect.objectContaining({ mode: "new_sheet", sheetName: "Revenue - Expense - August - Fi", tableName: "VGI_Forecast", result: queryResult }));
    expect(posted.some((item) => item.method === "excel.writeResult")).toBe(false);
  });

  it("creates a saved query tab without executing or changing Excel", async () => {
    const posted = installHost({ result: { columns: [], rows: [], rowCount: 0 } });
    const responses = [
      toolTurn("query-tab", "create_query_tab", '{"name":"Daily forecast","sql":"SELECT * FROM open_meteo.main.forecast_daily()"}'),
      textTurn("I created the Daily forecast query tab."),
    ];
    globalThis.fetch = vi.fn(async () => responses.shift()!) as unknown as typeof fetch;
    const { AgentSession } = await import("./agent");
    const events = callbacks();
    await new AgentSession().run("key", "model", "Save a forecast query", CONNECTION, events, new AbortController().signal);

    expect(events.onQueryDocument).toHaveBeenCalledWith({ name: "Daily forecast", sql: "SELECT * FROM open_meteo.main.forecast_daily()" });
    expect(posted.filter((item) => item.method === "query.run" && !String(item.params.sql).includes("information_schema.tables"))).toHaveLength(0);
    expect(posted.some((item) => item.method.startsWith("excel."))).toBe(false);
  });

  it("turns malformed streamed tool JSON into an error result without executing it", async () => {
    const posted = installHost({ result: {} });
    const requests: RequestInit[] = [];
    const responses = [toolTurn("broken", "run_sql", '{"sql":look'), textTurn("I could not run that malformed call.")];
    globalThis.fetch = vi.fn(async (_url, init) => { requests.push(init!); return responses.shift()!; }) as unknown as typeof fetch;
    const { AgentSession } = await import("./agent");
    await new AgentSession().run("test-key", "model", "try", CONNECTION, callbacks(), new AbortController().signal);

    expect(posted.filter((item) => item.method === "query.run" && !String(item.params.sql).includes("information_schema.tables"))).toHaveLength(0);
    const secondBody = JSON.parse(String(requests[1].body));
    expect(JSON.stringify(secondBody.messages)).toContain("not valid JSON");
  });

  it("filters function inventory and returns DuckDB named-parameter guidance", async () => {
    const flat = {
      columns: ["database_name", "schema_name", "function_name", "function_type", "parameters", "parameter_types", "description"].map((name) => ({ name, type: "VARCHAR" })),
      rows: [["open_meteo", "main", "forecast_current", "table", '["latitude","longitude","temperature_unit"]', '["DOUBLE","DOUBLE","VARCHAR"]', "Current weather"]], rowCount: 1, truncated: false,
    };
    const rich = {
      columns: ["catalog_name", "schema_name", "function_name", "arg_position", "arg_name", "arg_type", "is_named", "is_positional", "arg_choices"].map((name) => ({ name, type: "VARCHAR" })),
      rows: [
        ["open_meteo", "main", "forecast_current", 0, "latitude", "DOUBLE", false, true, null],
        ["open_meteo", "main", "forecast_current", null, "temperature_unit", "VARCHAR", true, false, '["celsius","fahrenheit"]'],
      ], rowCount: 2, truncated: false,
    };
    const posted = installHost((request) => ({ result: String(request.params.sql).includes("vgi_function_arguments") ? rich : flat }));
    const requests: RequestInit[] = [];
    const responses = [toolTurn("functions", "list_functions", '{"catalog":"open_meteo","name":"forecast"}'), textTurn("Use named options.")];
    globalThis.fetch = vi.fn(async (_url, init) => { requests.push(init!); return responses.shift()!; }) as unknown as typeof fetch;
    const { AgentSession } = await import("./agent");
    await new AgentSession().run("test-key", "model", "forecast", CONNECTION, callbacks(), new AbortController().signal);

    const queries = posted.filter((item) => item.method === "query.run" && !String(item.params.sql).includes("information_schema.tables"));
    expect(queries).toHaveLength(2);
    expect(queries[0].params.sql).toContain("database_name='open_meteo'");
    expect(queries[1].params.sql).toContain("catalog_name='open_meteo'");
    expect(queries.every((query) => query.params.sql.includes("function_name ILIKE '%forecast%'"))).toBe(true);
    const secondBody = JSON.parse(String(requests[1].body));
    expect(JSON.stringify(secondBody.messages)).toContain("name := value");
    const inventory = JSON.parse(secondBody.messages[2].content[0].content);
    expect(inventory.functions[0].arguments[1]).toMatchObject({ name: "temperature_unit", kind: "named", choices: ["celsius", "fahrenheit"] });
  });

  it("stops executing SQL after three failures in one turn", async () => {
    const posted = installHost({ error: "Binder Error: wrong signature" });
    const responses = [
      toolTurn("q1", "run_sql", '{"sql":"SELECT bad1()"}'),
      toolTurn("q2", "run_sql", '{"sql":"SELECT bad2()"}'),
      toolTurn("q3", "run_sql", '{"sql":"SELECT bad3()"}'),
      toolTurn("q4", "run_sql", '{"sql":"SELECT bad4()"}'),
      textTurn("I need the exact signature."),
    ];
    globalThis.fetch = vi.fn(async () => responses.shift()!) as unknown as typeof fetch;
    const { AgentSession } = await import("./agent");
    const events = callbacks();
    await new AgentSession().run("test-key", "model", "keep trying", CONNECTION, events, new AbortController().signal);

    expect(posted.filter((item) => item.method === "query.run" && String(item.params.sql).startsWith("SELECT bad"))).toHaveLength(3);
    expect(events.onTool).toHaveBeenCalledWith("run_sql", "error", expect.stringContaining("failure budget"), "q4");
  });

  it("logs only prompt length and never sends the API key to the native trace", async () => {
    const posted = installHost({ result: {} });
    globalThis.fetch = vi.fn(async () => textTurn("Done")) as unknown as typeof fetch;
    const { AgentSession } = await import("./agent");
    await new AgentSession().run("sk-ant-super-secret", "model", "private question", CONNECTION, callbacks(), new AbortController().signal);
    await Promise.resolve();

    const traces = posted.filter((item) => item.method === "agent.trace");
    expect(traces.length).toBeGreaterThan(0);
    expect(JSON.stringify(traces)).not.toContain("sk-ant-super-secret");
    expect(traces[0].params.event).toMatchObject({ event: "run_start", prompt_chars: 16 });
    expect(traces[0].params.event).not.toHaveProperty("prompt");
  });

  it("redacts credentials embedded in tool input before native tracing", async () => {
    const result = { columns: [], rows: [], rowCount: 0, truncated: false };
    const posted = installHost({ result });
    const responses = [toolTurn("secret-query", "run_sql", '{"sql":"SELECT thing(api_key := \'sql-secret\')"}'), textTurn("Done")];
    globalThis.fetch = vi.fn(async () => responses.shift()!) as unknown as typeof fetch;
    const { AgentSession } = await import("./agent");
    await new AgentSession().run("key", "model", "run", CONNECTION, callbacks(), new AbortController().signal);
    await Promise.resolve();

    const traces = posted.filter((item) => item.method === "agent.trace");
    expect(JSON.stringify(traces)).not.toContain("sql-secret");
    expect(JSON.stringify(traces)).toContain("***");
  });

  it("does not retry an invalid API key", async () => {
    installHost({ result: true });
    globalThis.fetch = vi.fn(async () => new Response("bad key", { status: 401 })) as unknown as typeof fetch;
    const { AgentSession } = await import("./agent");
    await expect(new AgentSession().run("bad", "model", "hello", CONNECTION, callbacks(), new AbortController().signal)).rejects.toThrow(/401/);
    expect(globalThis.fetch).toHaveBeenCalledTimes(1);
  });
});
