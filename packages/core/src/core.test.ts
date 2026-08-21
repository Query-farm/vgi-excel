import { describe, expect, it } from "vitest";
import { activateQueryDocument, addQueryDocument, agentRoundSeparator, appendAgentRoundText, assertAgentReadOnlySql, assertHttpsConnection, broadcastArguments, buildExcelAgentSystemPrompt, excelWorksheetName, executeAgentTool, formatAttachOptionsJson, loadQueryDocumentState, parseAttachOptionsJson, removeQueryDocument, renameQueryDocument, resultMatrix, saveQueryDocumentState, updateQueryDocumentSql, wrapperFormula, wrapperName, type AgentToolContext, type QueryDocumentStorage } from "./index.js";

function memoryStorage(): QueryDocumentStorage {
  const values = new Map<string, string>();
  return { getItem: (key) => values.get(key) ?? null, setItem: (key, value) => { values.set(key, value); } };
}

describe("Excel worksheet names", () => {
  it("normalizes AI labels before confirmation", () => {
    expect(excelWorksheetName("  Revenue / Expense: August [Final]  ")).toBe("Revenue - Expense - August - Fi");
    expect(excelWorksheetName("'Forecast'" )).toBe("Forecast");
    expect(excelWorksheetName("\\/?*[]:")).toBe("- - - - - - -");
    expect(excelWorksheetName("History")).toBe("History Data");
    expect(Array.from(excelWorksheetName("😀".repeat(40))).length).toBe(31);
  });
});

describe("persisted query documents", () => {
  it("persists named tabs and the active document per connection alias", () => {
    const storage = memoryStorage();
    let state = loadQueryDocumentState("weather-prod", storage);
    state = updateQueryDocumentSql(state, state.activeId, "select 1");
    state = addQueryDocument(state, "select 2");
    state = renameQueryDocument(state, state.activeId, "Daily forecast");
    saveQueryDocumentState("weather-prod", state, storage);

    const restored = loadQueryDocumentState("weather-prod", storage);
    expect(restored.documents.map((value) => [value.name, value.sql])).toEqual([["Query 1", "select 1"], ["Daily forecast", "select 2"]]);
    expect(restored.activeId).toBe(restored.documents[1].id);
    expect(loadQueryDocumentState("weather-dev", storage).documents).toHaveLength(1);
  });

  it("chooses stable names, activates tabs, and always leaves one document", () => {
    let state = loadQueryDocumentState("test", undefined, "");
    state = addQueryDocument(state);
    state = addQueryDocument(state);
    const second = state.documents[1].id;
    state = removeQueryDocument(state, second);
    state = addQueryDocument(state);
    expect(state.documents.map((value) => value.name)).toEqual(["Query 1", "Query 3", "Query 2"]);
    state = activateQueryDocument(state, state.documents[0].id);
    expect(state.activeId).toBe(state.documents[0].id);
    for (const document of [...state.documents]) state = removeQueryDocument(state, document.id);
    expect(state.documents).toHaveLength(1);
    expect(state.documents[0].sql).toBe("");
  });
});

describe("HTTPS connections", () => {
  it("accepts HTTPS and rejects every other transport", () => {
    expect(() => assertHttpsConnection({ name: "weather", location: "https://vgi.example.com" })).not.toThrow();
    expect(() => assertHttpsConnection({ name: "weather", location: "http://vgi.example.com" })).toThrow(/HTTPS/);
    expect(() => assertHttpsConnection({ name: "weather", location: "uv run worker.py" })).toThrow(/HTTPS/);
  });

  it("rejects credentials embedded in URLs", () => {
    expect(() => assertHttpsConnection({ name: "weather", location: "https://user:secret@vgi.example.com" })).toThrow(/Credentials/);
    expect(() => assertHttpsConnection({ name: "weather", location: "https://vgi.example.com", attachOptions: { bearer_token: "secret" } })).toThrow(/sign-in/);
  });

  it("parses safe ATTACH options and rejects managed or secret options", () => {
    expect(parseAttachOptionsJson('{"region":"us-east","cache":true,"retries":2}')).toEqual({ region: "us-east", cache: true, retries: 2 });
    expect(formatAttachOptionsJson({ region: "us-east" })).toContain('"region": "us-east"');
    expect(() => parseAttachOptionsJson('{"TYPE":"vgi"}')).toThrow(/managed/);
    expect(() => parseAttachOptionsJson('{"oauth_refresh_token":"secret"}')).toThrow(/sign-in/);
    expect(() => parseAttachOptionsJson('{"bad-key":"value"}')).toThrow(/option name/);
  });
});

describe("agent response rounds", () => {
  it("inserts paragraph breaks between tool-assisted response rounds", () => {
    expect(agentRoundSeparator("Let me check.", "Got the result.")).toBe("\n\n");
    expect(appendAgentRoundText("Let me check.", "Got the result.")).toBe("Let me check.\n\nGot the result.");
    expect(appendAgentRoundText("Let me check.\n", "Got the result.")).toBe("Let me check.\n\nGot the result.");
    expect(appendAgentRoundText("Let me check.\n\n", "Got the result.")).toBe("Let me check.\n\nGot the result.");
  });
});

describe("agent SQL gate", () => {
  it("allows one read query", () => expect(() => assertAgentReadOnlySql("with x as (select 1) select * from x;")).not.toThrow());
  it("rejects mutations and stacked statements", () => {
    expect(() => assertAgentReadOnlySql("delete from x")).toThrow(/read-only/);
    expect(() => assertAgentReadOnlySql("with x as (delete from t returning *) select * from x")).toThrow(/change data/);
    expect(() => assertAgentReadOnlySql("with x as (select 1) insert into t select * from x")).toThrow(/change data/);
    expect(() => assertAgentReadOnlySql("select 1; drop table x")).toThrow(/one SQL/);
    expect(() => assertAgentReadOnlySql("select 'ATTACH';")).not.toThrow();
  });
});

describe("agent query documents", () => {
  it("creates a named read-only query without executing it", async () => {
    const created: Array<{ name: string; sql: string }> = [];
    const context: AgentToolContext = {
      backend: { async query() { throw new Error("must not execute"); }, async call() { throw new Error("must not execute"); } },
      async listTables() { return []; }, async listFunctions() { return []; }, async describeTable() { return []; },
      createQuery(value) { created.push(value); },
    };
    const result = await executeAgentTool("create_query_tab", { name: "Weather summary", sql: "SELECT 1 AS temperature" }, context);
    expect(created).toEqual([{ name: "Weather summary", sql: "SELECT 1 AS temperature" }]);
    expect(result.content).toContain("has not been executed");
    await expect(executeAgentTool("create_query_tab", { name: "Unsafe", sql: "DELETE FROM readings" }, context)).rejects.toThrow(/read-only/);
  });
});

describe("agent system prompt", () => {
  it("identifies the active connection and live catalog without leaking its URL", () => {
    const prompt = buildExcelAgentSystemPrompt({
      connection: { name: "weather", catalog: "open_meteo", authentication: "oauth" },
      objects: [{ catalog: "open_meteo", schema: "main", name: "forecast_current", kind: "table function", description: "Current forecast" }],
    });
    expect(prompt).toContain("Connection name: weather");
    expect(prompt).toContain("Attached catalog: open_meteo");
    expect(prompt).toContain("`open_meteo.main.forecast_current`");
    expect(prompt).toContain("Authentication: oauth");
    expect(prompt).not.toContain("https://");
  });
});

describe("formula values", () => {
  it("broadcasts scalar arguments", () => {
    expect(broadcastArguments([[[1], [2]], 10])).toEqual([[[1], [2]], [[10], [10]]]);
  });
  it("adds headers", () => {
    expect(resultMatrix({ columns: [{ name: "n", type: "INTEGER" }], rows: [[1]], rowCount: 1 })).toEqual([["n"], [1]]);
  });
});

describe("catalog wrappers", () => {
  const fn = { catalog: "weather", schema: "main", name: "to-c", kind: "scalar" as const, parameters: [{ name: "value", type: "DOUBLE" }] };
  it("creates a stable Excel name", () => expect(wrapperName(fn)).toBe("VGI_WEATHER_MAIN_TO_C"));
  it("creates a lambda", () => expect(wrapperFormula(fn)).toBe('=LAMBDA(value,VGI.CALL("weather.main.to-c",value))'));
});
