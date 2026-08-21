import { describe, expect, it } from "vitest";
import { DateDay, Table, TimeMicrosecond, TimeMillisecond, TimeNanosecond, TimeSecond, TimestampMicrosecond, TimestampMillisecond, TimestampNanosecond, TimestampSecond, makeVector, tableFromArrays, vectorFromArray } from "apache-arrow";
import { arrowResult, excelScalar, isExcelSafeDecimal, setResultTimeZone } from "./arrow";
import { configureTimeZone, sanitizeTableName } from "./browser-backend";
import { OfficeAgentSession, runAgent } from "./anthropic";
import { discoverFunctions } from "./catalog";
import { authorizationUrl, oidcMetadataUrl, protectedResourceMetadataUrl } from "./oauth-core";
import { EXCEL_MAX_DATA_ROWS, assertCompleteExcelResult, validateExcelRange } from "./excel-limits";
import type { AgentToolContext, QueryResult } from "@query-farm/vgi-excel-core";
import { excelColumnNumberFormat } from "./accounting-format";

describe("Arrow to Excel conversion", () => {
  it("preserves unsafe integers as text", () => {
    expect(excelScalar(9007199254740993n)).toBe("9007199254740993");
    expect(excelScalar(42n)).toBe(42);
  });

  it("creates a tabular query result", () => {
    const result = arrowResult(tableFromArrays({ city: ["Paris", "Oslo"], value: [1, 2] }));
    expect(result.columns.map((column) => column.name)).toEqual(["city", "value"]);
    expect(result.rows).toEqual([["Paris", 1], ["Oslo", 2]]);
  });

  it("formats zoned timestamps in the user's DuckDB timezone", () => {
    setResultTimeZone("America/New_York");
    const instant = new Date("2026-08-19T22:15:30.000Z");
    expect(excelScalar(instant, { timezone: "UTC" })).toBe("2026-08-19 18:15:30");
    expect(excelScalar(instant, { timezone: null })).toBe("2026-08-19 22:15:30");
  });

  it("round trips Arrow timestamps into the values sent to Excel without a VGI worker", () => {
    setResultTimeZone("America/New_York");
    const instant = new Date("2026-08-19T22:15:30.000Z");
    const table = new Table({
      zoned: vectorFromArray([instant], new TimestampMillisecond("UTC")),
      naive: vectorFromArray([instant], new TimestampMillisecond()),
    });
    const result = arrowResult(table);
    expect(result.columns.map((column) => column.type)).toEqual(["Timestamp<MILLISECOND, UTC>", "Timestamp<MILLISECOND>"]);
    expect(result.rows).toEqual([["2026-08-19 18:15:30", "2026-08-19 22:15:30"]]);
  });

  it("covers DuckDB date, time, and timestamp resolutions at the Arrow-to-Excel boundary", () => {
    setResultTimeZone("America/New_York");
    const second = new Date("2026-08-19T22:15:30.000Z");
    const millisecond = new Date("2026-08-19T22:15:30.123Z");
    const table = new Table({
      date: vectorFromArray([new Date("2024-02-29T00:00:00Z")], new DateDay()),
      time_s: vectorFromArray([86_399], new TimeSecond()),
      time_ms: vectorFromArray([86_399_123], new TimeMillisecond()),
      time_us: vectorFromArray([86_399_123_456n], new TimeMicrosecond()),
      time_ns: vectorFromArray([86_399_123_456_789n], new TimeNanosecond()),
      timestamp_s: vectorFromArray([second], new TimestampSecond()),
      timestamp_ms: vectorFromArray([millisecond], new TimestampMillisecond()),
      timestamp_us: makeVector({ type: new TimestampMicrosecond(), data: new BigInt64Array([1_787_177_730_123_456n]) }),
      timestamp_ns: makeVector({ type: new TimestampNanosecond(), data: new BigInt64Array([1_787_177_730_123_456_789n]) }),
      timestamp_tz: makeVector({ type: new TimestampMicrosecond("UTC"), data: new BigInt64Array([1_787_177_730_123_456n]) }),
    });
    expect(arrowResult(table).rows[0]).toEqual([
      "2024-02-29", "23:59:59", "23:59:59.123", "23:59:59.123456", "23:59:59.123456789",
      "2026-08-19 22:15:30", "2026-08-19 22:15:30.123", "2026-08-19 22:15:30.123456",
      "2026-08-19 22:15:30.123456789", "2026-08-19 18:15:30.123456",
    ]);
  });

  it("preserves the bounds of DuckDB numeric families for Excel", () => {
    const cases: Array<[unknown, unknown]> = [
      [-128, -128], [127, 127], [-32_768, -32_768], [32_767, 32_767],
      [-2_147_483_648, -2_147_483_648], [2_147_483_647, 2_147_483_647],
      [0, 0], [255, 255], [65_535, 65_535], [4_294_967_295, 4_294_967_295],
      [-(2n ** 53n - 1n), -(2 ** 53 - 1)], [2n ** 53n - 1n, 2 ** 53 - 1],
      [-(2n ** 53n), "-9007199254740992"], [2n ** 53n, "9007199254740992"],
      [-(2n ** 63n), "-9223372036854775808"], [2n ** 63n - 1n, "9223372036854775807"],
      [2n ** 64n - 1n, "18446744073709551615"],
      [2n ** 127n - 1n, "170141183460469231731687303715884105727"],
      [2n ** 128n - 1n, "340282366920938463463374607431768211455"],
    ];
    for (const [input, expected] of cases) expect(excelScalar(input)).toBe(expected);
    expect(excelScalar({ toString: () => "123456789012345678901234567890" }, { typeId: 7, scale: 18 }))
      .toBe("123456789012.345678901234567890");
    expect(excelScalar({ toString: () => "-1" }, { typeId: 7, scale: 18 })).toBe(-0.000000000000000001);
    expect(excelScalar(3.4028234663852886e38)).toBe(3.4028234663852886e38);
    expect(excelScalar(Number.MAX_VALUE)).toBe(Number.MAX_VALUE);
    expect(excelScalar(Number.MIN_VALUE)).toBe(Number.MIN_VALUE);
    expect(excelScalar(Number.NaN)).toBe("NaN");
    expect(excelScalar(Number.POSITIVE_INFINITY)).toBe("Infinity");
    expect(excelScalar(Number.NEGATIVE_INFINITY)).toBe("-Infinity");
  });

  it("keeps ordinary monetary decimals numeric without sacrificing unsafe precision", () => {
    const decimal = (raw: string, scale: number, numeric?: boolean) => excelScalar({ toString: () => raw }, { typeId: 7, scale }, numeric);
    expect(decimal("9999", 2)).toBe(99.99);
    expect(decimal("-123450", 2)).toBe(-1234.5);
    expect(decimal("999999999999999", 2)).toBe(9_999_999_999_999.99);
    expect(decimal("9999999999999999", 2)).toBe("99999999999999.99");
    expect(decimal("001200", 0, false)).toBe("001200");
    expect(isExcelSafeDecimal("9999999999999.99")).toBe(true);
    expect(isExcelSafeDecimal("99999999999999.99")).toBe(false);
  });

  it("uses accounting-friendly number formats while protecting text precision", () => {
    expect(excelColumnNumberFormat("Decimal[18e+2]", [[99.99], [-1234.5]], 0)).toBe("#,##0.00");
    expect(excelColumnNumberFormat("DECIMAL(38,18)", [["12345678901234567890.123456789012345678"]], 0)).toBe("@");
    expect(excelColumnNumberFormat("BIGINT", [[42]], 0)).toBe("#,##0");
    expect(excelColumnNumberFormat("BIGINT", [["9223372036854775807"]], 0)).toBe("@");
  });

  it("loads ICU and sets the IANA timezone before attaching", async () => {
    const statements: string[] = [];
    await configureTimeZone({ async query(sql) { statements.push(sql); return {}; } }, "America/New_York");
    expect(statements).toEqual(["INSTALL icu", "LOAD icu", "SET TimeZone='America/New_York'"]);
  });
});

describe("large Excel snapshots", () => {
  it("accepts complete large results and enforces worksheet bounds", () => {
    const result: QueryResult = { columns: [{ name: "value", type: "INTEGER" }], rows: Array.from({ length: 100_000 }, (_value, index) => [index]), rowCount: 100_000 };
    expect(() => assertCompleteExcelResult(result)).not.toThrow();
    expect(() => assertCompleteExcelResult({ ...result, rows: result.rows.slice(0, 10_000), truncated: true })).toThrow(/preview/);
    expect(() => validateExcelRange(0, 0, 100_001, 1)).not.toThrow();
    expect(() => validateExcelRange(1, 0, EXCEL_MAX_DATA_ROWS + 1, 1)).toThrow(/does not fit/);
  });
});

describe("Excel snapshot names", () => {
  it("normalizes table names", () => expect(sanitizeTableName("Sales - 2026")).toBe("Sales_2026"));
});

describe("OAuth URL policy", () => {
  it("builds HTTPS discovery and authorization URLs", () => {
    expect(protectedResourceMetadataUrl("https://vgi.example.com/path")).toBe("https://vgi.example.com/.well-known/oauth-protected-resource");
    expect(oidcMetadataUrl("https://login.example.com/tenant")).toBe("https://login.example.com/tenant/.well-known/openid-configuration");
    const url = new URL(authorizationUrl("https://login.example.com/authorize", {
      clientId: "excel", redirectUri: "https://addin.example.com/oauth-dialog.html", scope: "openid offline_access",
      state: "state", challenge: "challenge",
    }));
    expect(url.searchParams.get("code_challenge_method")).toBe("S256");
    expect(url.searchParams.get("state")).toBe("state");
  });

  it("rejects insecure discovery and authorization endpoints", () => {
    expect(() => protectedResourceMetadataUrl("http://vgi.example.com")).toThrow(/HTTPS/);
    expect(() => authorizationUrl("http://login.example.com/authorize", {
      clientId: "excel", redirectUri: "https://addin.example.com/oauth-dialog.html", scope: "openid", state: "s", challenge: "c",
    })).toThrow(/HTTPS/);
  });
});

describe("agent loop", () => {
  it("restores saved model history before a follow-up", async () => {
    const context: AgentToolContext = {
      backend: { async query() { return { columns: [], rows: [], rowCount: 0 }; }, async call() { return { columns: [], rows: [], rowCount: 0 }; } },
      async listTables() { return []; }, async listFunctions() { return []; }, async describeTable() { return []; },
    };
    const bodies: Array<{ messages: unknown[] }> = [];
    const session = new OfficeAgentSession();
    session.restore([{ role: "user", content: "Earlier question" }, { role: "assistant", content: [{ type: "text", text: "Earlier answer" }] }]);
    await session.run("key", "Follow-up question", context, undefined, { fetchImpl: async (_url, init) => {
      bodies.push(JSON.parse(String(init?.body)) as { messages: unknown[] });
      return new Response(JSON.stringify({ content: [{ type: "text", text: "Follow-up answer" }], stop_reason: "end_turn" }), { status: 200, headers: { "content-type": "application/json" } });
    } });
    expect(bodies[0].messages).toEqual([
      { role: "user", content: "Earlier question" },
      { role: "assistant", content: [{ type: "text", text: "Earlier answer" }] },
      { role: "user", content: "Follow-up question" },
    ]);
    expect(session.snapshot()).toHaveLength(4);
  });

  it("executes a mocked read-only tool call and stages its result", async () => {
    const queryResult: QueryResult = { columns: [{ name: "value", type: "INTEGER" }], rows: [[42]], rowCount: 1 };
    const queries: string[] = [];
    const context: AgentToolContext = {
      backend: {
        async query(sql) { queries.push(sql); return queryResult; },
        async call() { throw new Error("not used"); },
      },
      async listTables() { return []; }, async listFunctions() { return []; }, async describeTable() { return []; },
    };
    let request = 0;
    const bodies: Record<string, unknown>[] = [];
    const fetchImpl: typeof fetch = async (_url, init) => {
      bodies.push(JSON.parse(String(init?.body)) as Record<string, unknown>);
      return new Response(JSON.stringify(request++ === 0 ? {
      content: [{ type: "text", text: "Let me check." }, { type: "tool_use", id: "tool-1", name: "run_sql", input: { sql: "SELECT 42 AS value" } }],
      stop_reason: "tool_use",
    } : {
      content: [{ type: "text", text: "The value is 42." }], stop_reason: "end_turn",
    }), { status: 200, headers: { "content-type": "application/json" } });
    };
    const answer = await runAgent("test-key", "Find the value", context, undefined, {
      endpoint: "https://agent.test/messages", model: "test-model", fetchImpl,
      connection: { name: "weather", catalog: "open_meteo", authentication: "oauth" },
    });
    expect(queries).toEqual(["SELECT 42 AS value"]);
    expect(answer.text).toBe("Let me check.\n\nThe value is 42.");
    expect(answer.stagedResult).toEqual(queryResult);
    expect(String(bodies[0].system)).toContain("Connection name: weather");
    expect(String(bodies[0].system)).toContain("Attached catalog: open_meteo");
  });

  it("routes create_query_tab through the host context without running SQL", async () => {
    const created: Array<{ name: string; sql: string }> = [];
    const context: AgentToolContext = {
      backend: { async query() { throw new Error("must not execute"); }, async call() { throw new Error("must not execute"); } },
      async listTables() { return []; }, async listFunctions() { return []; }, async describeTable() { return []; },
      createQuery(value) { created.push(value); },
    };
    let request = 0;
    const fetchImpl: typeof fetch = async () => new Response(JSON.stringify(request++ === 0 ? {
      content: [{ type: "tool_use", id: "query-tab", name: "create_query_tab", input: { name: "Weather summary", sql: "SELECT 1 AS temperature" } }], stop_reason: "tool_use",
    } : { content: [{ type: "text", text: "The query is ready in Query Editor." }], stop_reason: "end_turn" }), { status: 200, headers: { "content-type": "application/json" } });

    await runAgent("test-key", "Create a saved query", context, undefined, { endpoint: "https://agent.test/messages", fetchImpl });
    expect(created).toEqual([{ name: "Weather summary", sql: "SELECT 1 AS temperature" }]);
  });
});

describe("VGI function discovery", () => {
  it("combines complete DuckDB inventory with rich VGI argument metadata", async () => {
    const backend: AgentToolContext["backend"] = {
      async query(sql) {
        if (sql.includes("vgi_function_arguments")) return {
          columns: ["catalog_name", "schema_name", "function_name", "arg_position", "arg_name", "arg_type", "is_named", "is_positional", "is_varargs", "arg_choices"].map((name) => ({ name, type: "VARCHAR" })),
          rows: [
            ["open_meteo", "main", "forecast_current", 0, "latitude", "DOUBLE", false, true, false, null],
            ["open_meteo", "main", "forecast_current", null, "temperature_unit", "VARCHAR", true, false, false, '["celsius","fahrenheit"]'],
          ], rowCount: 2,
        };
        return {
          columns: ["database_name", "schema_name", "function_name", "function_type", "parameters", "parameter_types", "return_type", "description"].map((name) => ({ name, type: "VARCHAR" })),
          rows: [
            ["open_meteo", "main", "forecast_current", "table", '["latitude","temperature_unit"]', '["DOUBLE","VARCHAR"]', null, "Current weather"],
            ["other", "main", "total", "aggregate", "[]", "[]", "BIGINT", "Aggregate"],
          ], rowCount: 2,
        };
      },
      async call() { throw new Error("not used"); },
    };
    const functions = await discoverFunctions(backend);
    expect(functions).toHaveLength(2);
    expect(functions[0].parameters[1]).toMatchObject({ name: "temperature_unit", kind: "named", choices: ["celsius", "fahrenheit"] });
    expect(functions[1].kind).toBe("aggregate");
  });
});
