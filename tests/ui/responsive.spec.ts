import { expect, test, type Page } from "@playwright/test";

const widths = [300, 320, 340, 400, 720, 1060];

async function expectContained(page: Page): Promise<void> {
  await expect.poll(() => page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  const tabs = page.locator(".workspace-tabs");
  for (const name of ["Query Editor", "Ask AI", "Catalog View"]) {
    const tab = tabs.getByRole("tab", { name, exact: true });
    await expect(tab.locator("span")).toBeVisible();
    expect(await tab.evaluate(el => el.scrollWidth <= el.clientWidth)).toBe(true);
  }
  const footer = page.locator(".product-footer");
  await expect(footer).toHaveCount(0);
  const metrics = await page.evaluate(() => ({
    viewportHeight: innerHeight,
    documentHeight: document.documentElement.scrollHeight,
    shellBottom: document.querySelector(".app-shell")?.getBoundingClientRect().bottom ?? -1,
  }));
  expect(metrics.documentHeight).toBeLessThanOrEqual(metrics.viewportHeight + 1);
  expect(Math.abs(metrics.shellBottom - metrics.viewportHeight)).toBeLessThanOrEqual(1);
}

for (const width of widths) {
  test(`desktop contains onboarding at ${width}px`, async ({ page }) => {
    await page.setViewportSize({ width, height: 760 });
    await page.addInitScript(() => {
      const webview = { postMessage(request: { id: number; method: string }) { const result = request.method === "connections.list" ? [] : request.method === "app.diagnostics" ? "test diagnostics" : request.method === "agent.key.load" ? null : true; setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0); }, addEventListener() {} };
      Object.defineProperty(window, "chrome", { value: { webview }, configurable: true });
    });
    await page.goto("http://127.0.0.1:4173/index.html");
    await expect(page.getByRole("heading", { name: "Connect a VGI data source to begin" })).toBeVisible();
    await expectContained(page);
  });

  test(`Microsoft 365 contains connection settings at ${width}px`, async ({ page }) => {
    await page.setViewportSize({ width, height: 760 });
    await page.route("https://appsforoffice.microsoft.com/**", (route) => route.abort());
    await page.addInitScript(() => {
      localStorage.setItem("vgi.excel.connections.v1", JSON.stringify([{ name: "weather", catalog: "open_meteo", location: "https://example.com", authentication: "anonymous" }]));
      localStorage.setItem("vgi.excel.default-connection.v1", "weather");
    });
    await page.goto("http://127.0.0.1:4174/taskpane.html");
    await page.getByRole("button", { name: "Settings" }).click();
    await expect(page.getByRole("heading", { name: "Settings" })).toBeVisible();
    await expectContained(page);
  });
}

test("desktop workspace follows a live window resize", async ({ page }) => {
  await page.setViewportSize({ width: 900, height: 760 });
  await page.addInitScript(() => {
    const connection = { name: "weather", catalog: "open_meteo", location: "https://example.com", authentication: "anonymous", isDefault: true, isSignedIn: true };
    const webview = { postMessage(request: { id: number; method: string }) { const result = request.method === "connections.list" ? [connection] : request.method === "agent.key.load" ? null : ["query.run", "query.editor", "query.agent"].includes(request.method) ? { columns: [], rows: [], rowCount: 0, truncated: false } : request.method === "app.diagnostics" ? "test diagnostics" : request.method === "agent.key.load" ? null : true; setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0); }, addEventListener() {} };
    Object.defineProperty(window, "chrome", { value: { webview }, configurable: true });
  });
  await page.goto("http://127.0.0.1:4173/index.html");
  await page.getByRole("tab", { name: "Ask AI" }).click();
  const tall = await page.locator(".agent").boundingBox();
  await expectContained(page);

  await page.setViewportSize({ width: 420, height: 480 });
  await expectContained(page);
  const compact = await page.locator(".agent").boundingBox();
  expect(compact?.height ?? 0).toBeLessThan((tall?.height ?? 0) - 150);
  expect(compact?.width ?? 0).toBeLessThan(tall?.width ?? 0);
});

test("desktop query editor uses the Cupola icon toolbar without a header selector", async ({ page }) => {
  await page.setViewportSize({ width: 900, height: 680 });
  await page.addInitScript(() => {
    const connection = { name: "weather", catalog: "open_meteo", location: "https://example.com", authentication: "anonymous", isDefault: true, isSignedIn: true };
    const webview = { postMessage(request: { id: number; method: string }) { const result = request.method === "connections.list" ? [connection] : request.method === "agent.key.load" ? null : ["query.run", "query.editor", "query.agent"].includes(request.method) ? { columns: [], rows: [], rowCount: 0, truncated: false } : true; setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0); }, addEventListener() {} };
    Object.defineProperty(window, "chrome", { value: { webview }, configurable: true });
  });
  await page.goto("http://127.0.0.1:4173/index.html");
  expect(await page.locator(".workspace-tabs").getByRole("tab").allTextContents()).toEqual(["Query Editor", "Ask AI", "Catalog View"]);
  await expect(page.locator("header select")).toHaveCount(0);
  for (const name of ["Run", "Format", "Copy"]) await expect(page.getByRole("button", { name, exact: true }).locator("svg")).toHaveCount(1);
  await expect(page.locator(".query-history-menu")).toHaveCount(0);
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await expect(page.getByRole("status")).toHaveText("Query completed. No rows returned.");
  await expect(page.locator(".results-toolbar")).toHaveCount(0);
  await expect(page.locator(".notice")).toHaveCount(0);
  await expectContained(page);
});

test("desktop query preview pages independently from complete snapshot insertion", async ({ page }) => {
  await page.setViewportSize({ width: 1000, height: 760 });
  await page.addInitScript(() => {
    const connection = { name: "open_meteo", catalog: "open_meteo", location: "https://example.com", authentication: "anonymous", isDefault: true, isSignedIn: true };
    const queryResult = { columns: [{ name: "range", type: "NUMBER" }], rows: Array.from({ length: 1_000 }, (_, index) => [index]), rowCount: 100_000, truncated: true };
    const webview = { postMessage(request: { id: number; method: string; params?: { sql?: string } }) {
      localStorage.setItem("cupola.test.lastMethod", request.method);
      const result = request.method === "connections.list" ? [connection]
        : request.method === "agent.key.load" ? null
        : ["query.run", "query.editor", "query.agent"].includes(request.method) ? queryResult
        : request.method === "excel.createPowerQuery" ? { query: "Query 1", loaded: true, sheet: "Query 1", table: "Query_1", message: "Power Query created and refresh started." }
        : request.method === "excel.insertQuery" ? { sheet: "Sheet1", table: "VGI_Result", address: "$A$1:$A$100001" }
        : true;
      setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
    }, addEventListener() {} };
    Object.defineProperty(window, "chrome", { value: { webview }, configurable: true });
  });
  await page.goto("http://127.0.0.1:4173/index.html");
  await expect(page.locator("header .brand small")).toHaveCount(0);

  await page.getByRole("button", { name: "Run", exact: true }).click();
  const splitter = page.getByRole("separator", { name: "Resize query editor and results" });
  await splitter.focus();
  await splitter.press("ArrowDown");
  await expect(splitter).toHaveAttribute("aria-valuenow", "35");
  await expect(page.locator(".results-toolbar")).toContainText("100,000 rows · 1,000 loaded");
  await page.getByLabel("Rows shown per result page").selectOption("500");
  await page.getByRole("button", { name: "Next preview page" }).click();
  await expect(page.locator(".results-paging")).toContainText("501–1,000 of 1,000 loaded");

  await page.getByRole("button", { name: "Load into Excel" }).click();
  await expect(page.locator(".query-results-pane")).toContainText("Power Query created and refresh started.");
  await expect.poll(() => page.evaluate(() => localStorage.getItem("cupola.test.lastMethod"))).toBe("excel.createPowerQuery");

  await page.getByText("More", { exact: true }).click();
  await page.getByRole("button", { name: "Insert static table" }).click();
  await expect(page.getByText("Snapshot inserted", { exact: true })).toBeVisible();
  await expect.poll(() => page.evaluate(() => localStorage.getItem("cupola.test.lastMethod"))).toBe("excel.insertQuery");
});

test("desktop query tabs persist SQL and keep session results with their document", async ({ page }) => {
  await page.setViewportSize({ width: 1000, height: 760 });
  await page.addInitScript(() => {
    const connection = { name: "weather-prod", catalog: "open_meteo", location: "https://example.com", authentication: "anonymous", isDefault: true, isSignedIn: true };
    const webview = { postMessage(request: { id: number; method: string; params?: { sql?: string } }) {
      const marker = request.params?.sql?.includes("second") ? "second result" : "first result";
      const result = request.method === "connections.list" ? [connection]
        : request.method === "agent.key.load" ? null
        : ["query.run", "query.editor", "query.agent"].includes(request.method) ? { columns: [{ name: "marker", type: "VARCHAR" }], rows: [[marker]], rowCount: 1, truncated: false }
        : true;
      setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
    }, addEventListener() {} };
    Object.defineProperty(window, "chrome", { value: { webview }, configurable: true });
  });
  await page.goto("http://127.0.0.1:4173/index.html");
  const editor = page.getByLabel("SQL query");
  await editor.fill("SELECT 'first' AS marker");
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await expect(page.locator(".query-results-pane")).toContainText("first result");

  await page.getByRole("button", { name: "New query tab" }).click();
  await editor.fill("SELECT 'second' AS marker");
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await expect(page.locator(".query-results-pane")).toContainText("second result");
  await page.getByRole("tab", { name: "Query 2" }).dblclick();
  await page.getByLabel("Rename Query 2").fill("Forecast review");
  await page.getByLabel("Rename Query 2").press("Enter");

  await page.getByRole("tab", { name: "Query 1" }).click();
  await expect(editor).toHaveValue("SELECT 'first' AS marker");
  await expect(page.locator(".query-results-pane")).toContainText("first result");
  await page.getByRole("tab", { name: "Forecast review" }).click();
  await expect(editor).toHaveValue("SELECT 'second' AS marker");

  await page.reload();
  await expect(page.getByRole("tab", { name: "Forecast review" })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByLabel("SQL query")).toHaveValue("SELECT 'second' AS marker");
  await page.getByRole("button", { name: "Close Forecast review" }).click();
  await expect(page.getByRole("tab", { name: "Forecast review" })).toHaveCount(0);
  await expect(page.getByRole("tab", { name: "Query 1" })).toHaveAttribute("aria-selected", "true");
});

test("Ask AI can create a saved Query Editor tab without executing it", async ({ page }) => {
  await page.setViewportSize({ width: 1000, height: 760 });
  await page.addInitScript(() => {
    const connection = { name: "weather-prod", catalog: "open_meteo", location: "https://example.com", authentication: "anonymous", isDefault: true, isSignedIn: true };
    const empty = { columns: [], rows: [], rowCount: 0, truncated: false };
    const webview = { postMessage(request: { id: number; method: string; params?: { sql?: string } }) {
      if (["query.run", "query.editor", "query.agent"].includes(request.method) && !request.params?.sql?.includes("information_schema.tables")) localStorage.setItem("cupola.test.agentExecutedSql", request.params?.sql ?? "unknown");
      const result = request.method === "connections.list" ? [connection] : request.method === "agent.key.load" ? "test-key" : ["query.run", "query.editor", "query.agent"].includes(request.method) ? empty : true;
      setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
    }, addEventListener() {} };
    Object.defineProperty(window, "chrome", { value: { webview }, configurable: true });
  });
  let request = 0;
  await page.route("https://api.anthropic.com/v1/messages", async (route) => {
    const events = request++ === 0 ? [
      { type: "content_block_start", content_block: { type: "tool_use", id: "query-tab", name: "create_query_tab" } },
      { type: "content_block_delta", delta: { type: "input_json_delta", partial_json: '{"name":"NYC forecast","sql":"SELECT * FROM open_meteo.main.forecast_daily()"}' } },
      { type: "content_block_stop" },
      { type: "message_delta", delta: { stop_reason: "tool_use" } },
    ] : [
      { type: "content_block_start", content_block: { type: "text", text: "" } },
      { type: "content_block_delta", delta: { type: "text_delta", text: "I created the NYC forecast query." } },
      { type: "content_block_stop" },
      { type: "message_delta", delta: { stop_reason: "end_turn" } },
    ];
    await route.fulfill({ status: 200, contentType: "text/event-stream", body: events.map((event) => `data: ${JSON.stringify(event)}\n\n`).join("") });
  });
  await page.goto("http://127.0.0.1:4173/index.html");
  await page.locator(".workspace-tabs").getByRole("tab", { name: "Ask AI" }).click();
  await page.getByRole("textbox", { name: "Ask AI" }).fill("Save a query for the NYC forecast");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText("I created the NYC forecast query.")).toBeVisible();
  await expect(page.locator(".workspace-tabs").getByRole("tab", { name: "Ask AI" })).toHaveAttribute("aria-selected", "true");

  await page.locator(".workspace-tabs").getByRole("tab", { name: "Query Editor" }).click();
  await expect(page.getByRole("tab", { name: "NYC forecast" })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByLabel("SQL query")).toHaveValue("SELECT * FROM open_meteo.main.forecast_daily()");
  await expect.poll(() => page.evaluate(() => localStorage.getItem("cupola.test.agentExecutedSql"))).toBeNull();
});

for (const loadOutcome of ["loaded", "query-only", "error"] as const) {
test(`Ask AI stages a refreshable query and handles ${loadOutcome} after later replies`, async ({ page }) => {
  await page.setViewportSize({ width: 440, height: 760 });
  await page.addInitScript(({ loadOutcome }) => {
    (window as any).writes = [];
    const connection = { name: "weather-prod", catalog: "open_meteo", location: "https://example.com", authentication: "anonymous", isDefault: true, isSignedIn: true };
    const empty = { columns: [], rows: [], rowCount: 0, truncated: false };
    const forecast = { columns: [{ name: "time", type: "TIMESTAMP" }, { name: "temperature", type: "DOUBLE" }], rows: [["2026-08-19T18:00:00", 81]], rowCount: 168, truncated: false };
    const webview = { postMessage(request: { id: number; method: string; params?: { sql?: string } }) {
      let result: unknown = true;
      if (request.method === "connections.list") result = [connection];
      else if (request.method === "agent.key.load") result = "test-key";
      else if (["query.run", "query.editor", "query.agent"].includes(request.method)) result = request.params?.sql === "SELECT * FROM hourly_forecast" ? forecast : empty;
      else if (request.method === "excel.writeResult") { throw new Error("Static insertion must not be used for refreshable output"); }
      else if (request.method === "excel.createPowerQuery") {
        (window as any).writes.push(request.params);
        localStorage.setItem("cupola.test.confirmedWrite", "true");
        if (loadOutcome === "error") { setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, error: "The driver is unavailable." }), 0); return; }
        result = loadOutcome === "loaded" ? { query: "Forecast query", loaded: true, sheet: "Glen Allen VA - Hourly Forecast", table: "VGI_Hourly_Forecast_2", message: "Refresh started." } : { query: "Forecast query", loaded: false, message: "The query was created. Finish loading in Queries & Connections." };
      }
      setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
    }, addEventListener() {} };
    Object.defineProperty(window, "chrome", { value: { webview }, configurable: true });
  }, { loadOutcome });
  let request = 0;
  await page.route("https://api.anthropic.com/v1/messages", async (route) => {
    const body = route.request().postDataJSON() as { messages?: Array<{ content?: unknown }> };
    const tool = (name: string, input: string) => [
      { type: "content_block_start", content_block: { type: "tool_use", id: `${name}-${request}`, name } },
      { type: "content_block_delta", delta: { type: "input_json_delta", partial_json: input } },
      { type: "content_block_stop" },
      { type: "message_delta", delta: { stop_reason: "tool_use" } },
    ];
    const text = (value: string) => [
      { type: "content_block_start", content_block: { type: "text", text: "" } },
      { type: "content_block_delta", delta: { type: "text_delta", text: value } },
      { type: "content_block_stop" },
      { type: "message_delta", delta: { stop_reason: "end_turn" } },
    ];
    let events;
    if (request === 0) events = tool("run_sql", '{"sql":"SELECT * FROM hourly_forecast"}');
    else if (request === 1) {
      const blocks = (body.messages ?? []).flatMap((message) => Array.isArray(message.content) ? message.content as Array<{ type?: string; content?: string }> : []);
      const result = [...blocks].reverse().find((block) => block.type === "tool_result");
      const resultId = JSON.parse(result?.content ?? "{}").result_id;
      events = tool("stage_result_to_new_sheet", JSON.stringify({ result_id: resultId, sheet_name: "Glen Allen VA - Hourly Forecast", table_name: "VGI_Hourly_Forecast" }));
    } else if (request === 2) events = text("The hourly forecast is ready for your confirmation.");
    else events = text("Here are some additional notes about the forecast.");
    request += 1;
    await route.fulfill({ status: 200, contentType: "text/event-stream", body: events.map((event) => `data: ${JSON.stringify(event)}\n\n`).join("") });
  });

  await page.goto("http://127.0.0.1:4173/index.html");
  await page.locator(".workspace-tabs").getByRole("tab", { name: "Ask AI" }).click();
  const prompt = page.getByRole("textbox", { name: "Ask AI" });
  await prompt.fill("Put the hourly forecast in Excel");
  await page.getByRole("button", { name: "Send" }).click();
  const tray = page.getByLabel("Workbook actions");
  await expect(tray).toBeVisible();
  await expect(page.locator(".agent > .chat + .workbook-action-tray")).toHaveCount(1);
  await expect(tray.getByText("Create refreshable Excel table")).toBeVisible();
  await expect(tray.getByText("Preview: 1 row")).toBeVisible();
  await expect(tray.getByText("New worksheet")).toBeVisible();
  await expect(tray.getByText("“Glen Allen VA - Hourly Forecast”")).toBeVisible();
  const queryTool = page.locator(".tool").filter({ hasText: "SQL query · complete" }).first();
  await queryTool.locator("summary").click();
  await expect(queryTool.getByText("Executed SQL")).toBeVisible();
  await expect(queryTool.locator("pre")).toHaveText("SELECT * FROM hourly_forecast");
  await expect(queryTool.locator("pre")).not.toContainText("result_id");
  await expect(queryTool.getByRole("button", { name: "Copy SQL" })).toBeVisible();

  await prompt.fill("Anything else I should know?");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText("Here are some additional notes about the forecast.")).toBeVisible();
  const confirm = tray.getByRole("button", { name: "Load into Excel", exact: true });
  await expect(confirm).toBeVisible();
  await page.setViewportSize({ width: 360, height: 600 });
  await expect(page.getByRole("button", { name: "Load into Excel", exact: true })).toHaveCount(1);
  await expect(confirm).toBeInViewport();
  if (loadOutcome === "loaded") await page.screenshot({ path: "/tmp/cupola-ai-refreshable.png" });
  const bounds = await confirm.boundingBox();
  expect(bounds && bounds.x + bounds.width).toBeLessThanOrEqual(360);
  expect(await page.evaluate(() => (window as any).writes)).toEqual([]);
  await confirm.click();
  await expect.poll(() => page.evaluate(() => localStorage.getItem("cupola.test.confirmedWrite"))).toBe("true");
  expect(await page.evaluate(() => (window as any).writes)).toEqual([{ sql: "SELECT * FROM hourly_forecast", connection: "weather-prod", name: "VGI_Hourly_Forecast", loadToWorksheet: true, sheetName: "Glen Allen VA - Hourly Forecast", tableName: "VGI_Hourly_Forecast" }]);
  if (loadOutcome === "loaded") await expect(tray.getByRole("button", { name: "Go to table" })).toBeVisible();
  else if (loadOutcome === "query-only") {
    await expect(tray.getByText("Query created — finish loading in Excel")).toBeVisible();
    await expect(tray.getByRole("button")).toHaveCount(0);
  } else {
    await expect(tray).toContainText("The driver is unavailable.");
    await expect(tray.getByRole("button", { name: "Try again" })).toBeVisible();
  }
});
}

test("Ask AI conversation tabs persist across closing and reopening the Workbench", async ({ page }) => {
  await page.setViewportSize({ width: 760, height: 760 });
  await page.addInitScript(() => {
    const connection = { name: "weather-prod", catalog: "open_meteo", location: "https://example.com", authentication: "anonymous", isDefault: true, isSignedIn: true };
    const empty = { columns: [], rows: [], rowCount: 0, truncated: false };
    const webview = { postMessage(request: { id: number; method: string }) {
      const result = request.method === "connections.list" ? [connection] : request.method === "agent.key.load" ? "test-key" : ["query.run", "query.editor", "query.agent"].includes(request.method) ? empty : true;
      setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
    }, addEventListener() {} };
    Object.defineProperty(window, "chrome", { value: { webview }, configurable: true });
  });
  let request = 0;
  await page.route("https://api.anthropic.com/v1/messages", async (route) => {
    const reply = request++ === 0 ? "Richmond is warm today." : "Rain is likely tomorrow.";
    const events = [
      { type: "content_block_start", content_block: { type: "text", text: "" } },
      { type: "content_block_delta", delta: { type: "text_delta", text: reply } },
      { type: "content_block_stop" },
      { type: "message_delta", delta: { stop_reason: "end_turn" } },
    ];
    await route.fulfill({ status: 200, contentType: "text/event-stream", body: events.map((event) => `data: ${JSON.stringify(event)}\n\n`).join("") });
  });

  await page.goto("http://127.0.0.1:4173/index.html");
  await page.locator(".workspace-tabs").getByRole("tab", { name: "Ask AI" }).click();
  const prompt = page.getByRole("textbox", { name: "Ask AI" });
  await prompt.fill("Weather in Richmond");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText("Richmond is warm today.")).toBeVisible();
  await expect(page.getByRole("tab", { name: "Weather in Richmond" })).toBeVisible();

  await page.getByRole("button", { name: "New conversation tab" }).click();
  await prompt.fill("Rain tomorrow");
  await page.getByRole("button", { name: "Send" }).click();
  await expect(page.getByText("Rain is likely tomorrow.")).toBeVisible();
  await expect(page.getByRole("tab", { name: "Rain tomorrow" })).toHaveAttribute("aria-selected", "true");

  await page.reload();
  await page.locator(".workspace-tabs").getByRole("tab", { name: "Ask AI" }).click();
  await expect(page.getByRole("tab", { name: "Weather in Richmond" })).toBeVisible();
  await expect(page.getByRole("tab", { name: "Rain tomorrow" })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByText("Rain is likely tomorrow.")).toBeVisible();
  await page.getByRole("tab", { name: "Weather in Richmond" }).click();
  await expect(page.getByText("Richmond is warm today.")).toBeVisible();
  await expect.poll(() => page.evaluate(() => [...Object.entries(localStorage)].filter(([key]) => key.startsWith("cupola.agent.conversations")).map(([, value]) => value).join(""))).not.toContain("test-key");
});

test("Microsoft 365 restores persisted AI conversation tabs without storing its API key", async ({ page }) => {
  await page.setViewportSize({ width: 720, height: 760 });
  await page.route("https://appsforoffice.microsoft.com/**", (route) => route.abort());
  await page.addInitScript(() => {
    localStorage.setItem("vgi.excel.connections.v1", JSON.stringify([{ name: "weather-prod", catalog: "open_meteo", location: "https://example.com", authentication: "anonymous" }]));
    localStorage.setItem("vgi.excel.default-connection.v1", "weather-prod");
    if (!localStorage.getItem("cupola.office.agent.conversations.v1::weather-prod")) localStorage.setItem("cupola.office.agent.conversations.v1::weather-prod", JSON.stringify({
      version: 1,
      activeId: "close-review",
      documents: [{
        id: "close-review", name: "Month-end close", model: "test-model", draft: "",
        displayMessages: [{ role: "user", text: "Reconcile cash" }, { role: "assistant", text: "The cash query is ready." }],
        agentMessages: [{ role: "user", content: "Reconcile cash" }, { role: "assistant", content: [{ type: "text", text: "The cash query is ready." }] }],
        createdAt: 1, updatedAt: 2,
      }],
    }));
  });
  await page.goto("http://127.0.0.1:4174/taskpane.html");
  await page.locator(".workspace-tabs").getByRole("tab", { name: "Ask AI" }).click();
  await expect(page.getByRole("tab", { name: "Month-end close" })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByText("The cash query is ready.")).toBeVisible();
  await page.getByRole("button", { name: "New conversation tab" }).click();
  await page.getByRole("textbox", { name: "Ask AI" }).fill("Draft for the next conversation");
  await expect(page.getByRole("tab", { name: "Conversation 1" })).toHaveAttribute("aria-selected", "true");
  await page.reload();
  await page.locator(".workspace-tabs").getByRole("tab", { name: "Ask AI" }).click();
  await expect(page.getByRole("tab", { name: "Month-end close" })).toBeVisible();
  await expect(page.getByRole("tab", { name: "Conversation 1" })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByRole("textbox", { name: "Ask AI" })).toHaveValue("Draft for the next conversation");
  await expect.poll(() => page.evaluate(() => [...Object.entries(localStorage)].filter(([key]) => key.startsWith("cupola.office.agent.conversations")).map(([, value]) => value).join(""))).not.toContain("api-key");
});

test("Microsoft 365 query tabs restore local SQL", async ({ page }) => {
  await page.setViewportSize({ width: 720, height: 760 });
  await page.route("https://appsforoffice.microsoft.com/**", (route) => route.abort());
  await page.addInitScript(() => {
    localStorage.setItem("vgi.excel.connections.v1", JSON.stringify([{ name: "weather-prod", catalog: "open_meteo", location: "https://example.com", authentication: "anonymous" }]));
    localStorage.setItem("vgi.excel.default-connection.v1", "weather-prod");
  });
  await page.goto("http://127.0.0.1:4174/taskpane.html");
  const editor = page.getByLabel("SQL query");
  await editor.fill("SELECT 'office first'");
  await page.getByRole("button", { name: "New query tab" }).click();
  await editor.fill("SELECT 'office second'");
  await page.getByRole("tab", { name: "Query 2" }).dblclick();
  await page.getByLabel("Rename Query 2").fill("Workbook query");
  await page.getByLabel("Rename Query 2").press("Enter");
  await page.reload();
  await expect(page.getByRole("tab", { name: "Workbook query" })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByLabel("SQL query")).toHaveValue("SELECT 'office second'");
});

test("connections discover authentication instead of asking users to choose it", async ({ page }) => {
  await page.setViewportSize({ width: 900, height: 680 });
  await page.addInitScript(() => {
    const connection = { name: "weather", catalog: "open_meteo", location: "https://example.com", authentication: "anonymous", isDefault: true, isSignedIn: false };
    const webview = { postMessage(request: { id: number; method: string }) { const result = request.method === "connections.list" ? [connection] : request.method === "agent.key.load" ? null : true; setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0); }, addEventListener() {} };
    Object.defineProperty(window, "chrome", { value: { webview }, configurable: true });
  });
  await page.goto("http://127.0.0.1:4173/index.html");
  await page.getByRole("button", { name: "Settings" }).click();
  await expect(page.getByLabel("Authentication")).toHaveCount(0);
  await expect(page.getByText("If this service requires authentication, Cupola opens your browser when it connects.")).toBeVisible();

  await page.route("https://appsforoffice.microsoft.com/**", (route) => route.abort());
  await page.addInitScript(() => {
    localStorage.setItem("vgi.excel.connections.v1", JSON.stringify([{ name: "weather", catalog: "open_meteo", location: "https://example.com", authentication: "anonymous" }]));
    localStorage.setItem("vgi.excel.default-connection.v1", "weather");
  });
  await page.goto("http://127.0.0.1:4174/taskpane.html");
  await page.getByRole("button", { name: "Settings" }).click();
  await expect(page.getByLabel("Authentication")).toHaveCount(0);
  await expect(page.getByText("If this service requires authentication, Cupola opens a secure sign-in window when it connects.")).toBeVisible();
});

test("desktop catalog keeps scrolling inside the schema and inspector panes", async ({ page }) => {
  await page.setViewportSize({ width: 900, height: 520 });
  await page.addInitScript(() => {
    const connection = { name: "weather", catalog: "open_meteo", location: "https://example.com", authentication: "anonymous", isDefault: true, isSignedIn: false };
    const catalog = {
      columns: ["catalog", "schema", "name", "object_type", "kind", "summary"].map((name) => ({ name, type: "VARCHAR" })),
      rows: Array.from({ length: 100 }, (_, index) => ["open_meteo", "main", `function_${index}`, "table", "BASE TABLE", "Test table"]),
      rowCount: 100,
      truncated: false,
    };
    const fields = { columns: [{ name: "column_name", type: "VARCHAR" }, { name: "data_type", type: "VARCHAR" }], rows: [["value", "DOUBLE"]], rowCount: 1, truncated: false };
    const webview = { postMessage(request: { id: number; method: string; params?: { sql?: string } }) { const result = request.method === "connections.list" ? [connection] : request.method === "agent.key.load" ? null : ["query.run", "query.editor", "query.agent"].includes(request.method) ? (request.params?.sql?.includes("duckdb_columns()") ? fields : catalog) : true; setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0); }, addEventListener() {} };
    Object.defineProperty(window, "chrome", { value: { webview }, configurable: true });
  });
  await page.goto("http://127.0.0.1:4173/index.html");
  await page.getByRole("tab", { name: "Catalog View" }).click();
  const panel = page.locator("#panel-catalog");
  const tree = panel.locator(".catalog-tree");
  await expect(tree.getByText("function_99")).toBeAttached();
  const layout = await panel.evaluate((element) => {
    const browser = element.querySelector(".catalog-browser")!;
    const tree = element.querySelector(".catalog-tree-scroll")!;
    const inspector = element.querySelector(".catalog-inspector")!;
    return {
      panelOverflow: getComputedStyle(element).overflowY,
      panelFits: element.scrollHeight <= element.clientHeight + 1,
      browserFits: browser.getBoundingClientRect().bottom <= element.getBoundingClientRect().bottom + 1,
      treeScrolls: tree.scrollHeight > tree.clientHeight,
      treeOverflow: getComputedStyle(tree).overflowY,
      inspectorOverflow: getComputedStyle(inspector).overflowY,
    };
  });
  expect(layout).toEqual({ panelOverflow: "hidden", panelFits: true, browserFits: true, treeScrolls: true, treeOverflow: "auto", inspectorOverflow: "auto" });
  await tree.getByText("function_99").click();
  await page.getByRole("button", { name: "Insert into query" }).click();
  await expect(page.getByRole("tab", { name: "Query 2" })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByLabel("SQL query")).toHaveValue(/FROM "open_meteo"\."main"\."function_99"/);
});

for (const host of ["desktop", "office"] as const) {
  test(`${host} can open Connections with an empty store and retain a saved connection after reload`, async ({ page }) => {
    await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
    await page.addInitScript(() => {
      const webview = { postMessage(request: { id: number; method: string; params?: { connection?: Record<string, unknown> } }) {
        let connections = JSON.parse(localStorage.getItem("test.connections") ?? "[]");
        if (request.method === "connections.save") {
          connections = [{ ...request.params?.connection, isDefault: true }];
          localStorage.setItem("test.connections", JSON.stringify(connections));
        }
        const result = request.method.startsWith("connections.") ? connections : request.method === "agent.key.load" ? null : true;
        setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
      }, addEventListener() {} };
      Object.defineProperty(window, "chrome", { value: { webview }, configurable: true });
    });
    await page.goto(host === "desktop" ? "http://127.0.0.1:4173/index.html" : "http://127.0.0.1:4174/taskpane.html");
    await page.getByRole("button", { name: "Settings", exact: true }).click();
    await expect(page.getByRole("heading", { name: "Settings", exact: true })).toBeVisible();
    await expect(page.getByRole("button", { name: "Settings", exact: true })).toHaveAttribute("aria-pressed", "true");
    await expect(page.locator(".workspace-tabs").getByRole("tab", { name: "Connections", exact: true })).toHaveCount(0);
    await expect(page.locator('.workspace-tabs [aria-selected="true"]')).toHaveCount(0);
    await page.getByRole("tab", { name: "Query Editor", exact: true }).focus();
    await page.keyboard.press("ArrowRight");
    await expect(page.getByRole("tab", { name: "Ask AI", exact: true })).toBeFocused();
    await page.getByRole("button", { name: "Settings", exact: true }).click();
    await page.getByRole("button", { name: "Back to workspace", exact: true }).click();
    await expect(page.getByRole("tab", { name: "Ask AI", exact: true })).toHaveAttribute("aria-selected", "true");
    await expect(page.getByRole("button", { name: "Settings", exact: true })).toHaveAttribute("aria-pressed", "false");
    await page.getByRole("button", { name: "Settings", exact: true }).click();
    await page.getByRole("tab", { name: "About", exact: true }).click();
    const about = page.getByRole("tabpanel", { name: "About", exact: true });
    await expect(about).toBeVisible();
    await expect(about.getByRole("link", { name: "Query.Farm" })).toHaveAttribute("href", "https://query.farm");
    await expect(about).toContainText("Version 0.5.0");
    await expect(page.getByRole("dialog")).toHaveCount(0);
    await page.getByRole("tab", { name: "Connections", exact: true }).click();
    await page.getByLabel("Connection name", { exact: true }).fill("saved-test");
    await page.locator(".connection-form").getByLabel("Catalog", { exact: true }).fill("sample");
    await page.getByLabel("Server address", { exact: true }).fill("https://example.com");
    await page.getByRole("button", { name: "Save changes", exact: true }).click();
    await page.reload();
    await page.getByRole("button", { name: "Settings", exact: true }).click();
    await expect(page.getByRole("button", { name: /saved-test/ })).toBeVisible();
    await expect(page.getByRole("heading", { name: "Cupola needs to restart" })).toHaveCount(0);
  });
}
