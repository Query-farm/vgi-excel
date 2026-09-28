import { expect, test } from "@playwright/test";

const endpoint = process.env.CUPOLA_LIVE_VGI_ENDPOINT ?? "https://vgi-open-meteo.rusty-bb6.workers.dev";

test("self-hosted Haybarn WASM attaches an HTTPS Catalog and executes SQL", async ({ page }) => {
  const engineAssets = new Set<string>();
  page.on("response", (response) => {
    if (response.url().includes("/haybarn/")) engineAssets.add(new URL(response.url()).pathname.split("/").pop() ?? "");
  });
  await page.route("https://appsforoffice.microsoft.com/**", (route) => route.abort());
  await page.addInitScript(({ location }) => {
    localStorage.setItem("vgi.excel.connections.v1", JSON.stringify([{
      name: "open-meteo-live",
      catalog: "open_meteo",
      location,
      authentication: "anonymous",
      attachOptions: {},
    }]));
    localStorage.setItem("vgi.excel.default-connection.v1", "open-meteo-live");
  }, { location: endpoint });

  await page.goto("https://127.0.0.1:4184/taskpane.html");
  expect(await page.evaluate(() => globalThis.crossOriginIsolated)).toBe(true);
  await expect(page.locator(".query-results-pane")).toHaveCount(0);
  await page.getByLabel("SQL query").fill("SELECT 42 AS wasm_answer, current_setting('TimeZone') AS local_time_zone;");
  await page.getByRole("button", { name: "Run", exact: true }).click();

  const results = page.getByRole("table", { name: "Query results" });
  await expect(results).toContainText("wasm_answer");
  await expect(results).toContainText("42");
  expect([...engineAssets].some((name) => name.endsWith(".wasm"))).toBe(true);
  expect([...engineAssets].some((name) => name.endsWith(".worker.js"))).toBe(true);

  await page.getByLabel("SQL query").fill("CREATE TEMP TABLE cancellation_probe AS SELECT 84 AS after_cancel");
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await expect(results).toContainText("Count");

  await page.getByLabel("SQL query").fill(`SELECT catalog FROM vgi_catalogs('${endpoint.replaceAll("'", "''")}');`);
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await expect(results).toContainText("open_meteo");

  // Interrupt real CPU work, retain the last result, and reuse the same editor session.
  await page.emulateMedia({ reducedMotion: "reduce" });
  await page.getByLabel("SQL query").fill("SELECT sum(sin(i)) FROM range(100000000000) t(i)");
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await expect(page.getByRole("button", { name: "Cancel query", exact: true })).toBeEnabled();
  await expect(page.locator(".query-running-status")).toHaveText("Running query…");
  await expect(page.locator(".query-toolbar .busy-spinner")).toHaveCSS("animation-name", "none");
  await page.waitForTimeout(300);
  await page.getByRole("button", { name: "Cancel query", exact: true }).click();
  await expect(page.getByRole("status")).toContainText("Query cancelled.", { timeout: 10_000 });
  await expect(results).toContainText("open_meteo");
  await expect(page.locator(".notice.error")).toHaveCount(0);
  await page.getByLabel("SQL query").fill("SELECT after_cancel FROM cancellation_probe");
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await expect(results).toContainText("84");

  const popupPromise = page.waitForEvent("popup");
  await page.getByRole("button", { name: "Open results in new window" }).click();
  const popup = await popupPromise;
  await expect(popup.getByRole("table", { name: "Query results" })).toContainText("84");
  await page.getByLabel("SQL query").fill("SELECT 123 AS next_result");
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await expect(results).toContainText("123");
  await expect(popup.getByRole("table")).toContainText("84");
  await popup.getByRole("button", { name: "Close", exact: true }).click();
  await expect.poll(() => popup.isClosed()).toBe(true);

  await page.getByRole("button", { name: "Hide results", exact: true }).click();
  await page.getByLabel("SQL query").fill("SELECT 456+i AS collapsed_result FROM range(3) t(i)");
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await expect(page.getByRole("status")).toHaveText("Query completed.");
  await expect(page.getByRole("button", { name: "Show results", exact: true })).toBeVisible();
  await page.getByRole("button", { name: "Show results", exact: true }).click();
  await expect(results).toContainText("456");
  await page.setViewportSize({ width: 300, height: 480 });
  const visibleCell = (await page.getByRole("cell", { name: "458", exact: true }).boundingBox())!;
  const visibleGrid = (await page.locator(".query-results-content .result-wrap").boundingBox())!;
  expect(visibleCell.y + visibleCell.height).toBeLessThanOrEqual(visibleGrid.y + visibleGrid.height);
  expect(visibleCell.y + visibleCell.height).toBeLessThanOrEqual(480);
  await page.screenshot({ path: "/tmp/cupola-expanded-office-loaded.png" });
  await page.getByRole("button", { name: "Hide results", exact: true }).click();
  await page.screenshot({ path: "/tmp/cupola-collapsed-office-loaded.png" });
  await page.getByRole("button", { name: "Show results", exact: true }).click();
  await page.setViewportSize({ width: 1280, height: 720 });


  await page.getByLabel("SQL query").fill("SELECT 1 WHERE false");
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await expect(page.getByRole("status")).toHaveText("Query completed. No rows returned.");
  await expect(page.locator(".query-results-pane")).toHaveCount(0);

  // Exercise the real metadata SQL through Office's AI host, without paid API calls.
  await page.route(/sentry\.io\//, route => route.abort());
  let toolResults: Array<{ is_error?: boolean; content: string }> = [];
  let requests = 0;
  await page.route("https://api.anthropic.com/v1/messages", async route => {
    const body = route.request().postDataJSON();
    if (requests++ === 0) {
      await route.fulfill({ json: { content: [
        { type: "tool_use", id: "functions", name: "list_functions", input: { catalog: "open_meteo", name: "forecast", limit: 2 } },
        { type: "tool_use", id: "function", name: "describe_function", input: { catalog: "open_meteo", schema: "main", function: "forecast_current" } },
        { type: "tool_use", id: "categories", name: "list_categories", input: { catalog: "open_meteo" } },
        { type: "tool_use", id: "tables", name: "list_tables", input: { catalog: "open_meteo" } },
        { type: "tool_use", id: "table", name: "describe_table", input: { catalog: "open_meteo", schema: "main", table: "__cupola_metadata_probe__" } },
      ], stop_reason: "tool_use" } });
    } else {
      toolResults = body.messages.at(-1).content;
      await route.fulfill({ json: { content: [{ type: "text", text: "Catalog metadata checked." }], stop_reason: "end_turn" } });
    }
  });
  await page.getByRole("tab", { name: "Ask AI", exact: true }).click();
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await page.getByRole("tab", { name: "AI settings", exact: true }).click();
  await page.getByLabel("Anthropic API key").fill("test-only-no-network");
  await page.getByRole("button", { name: "Back to workspace", exact: true }).click();
  await page.getByRole("textbox", { name: "Ask AI", exact: true }).fill("Inspect catalog metadata");
  await page.getByRole("button", { name: "Send", exact: true }).click();
  await expect(page.getByText("Catalog metadata checked.")).toBeVisible();
  expect(toolResults).toHaveLength(5);
  expect(toolResults.filter(result => result.is_error)).toEqual([]);
  const detail = JSON.parse(toolResults[1].content);
  expect(detail.objects[0].name).toBe("forecast_current");
  expect(detail.objects[0].arguments.some((arg: { name: string; kind: string }) => arg.name === "temperature_unit" && arg.kind === "named")).toBe(true);

  // Cancel real WASM work launched by the agent, then continue the same conversation.
  await page.unroute("https://api.anthropic.com/v1/messages");
  let cancellationRequests = 0;
  await page.route("https://api.anthropic.com/v1/messages", async route => {
    cancellationRequests++;
    await route.fulfill({ json: cancellationRequests === 1
      ? { content: [{ type: "tool_use", id: "slow-agent-query", name: "run_sql", input: { sql: "SELECT sum(sin(i)) FROM range(100000000000) t(i)" } }], stop_reason: "tool_use" }
      : cancellationRequests === 2
      ? { content: [{ type: "thinking", thinking: "Check the recovered calculation.", signature: "private-signature" }, { type: "tool_use", id: "recovered-agent-query", name: "run_sql", input: { sql: "SELECT 42 AS recovered" } }], stop_reason: "tool_use" }
      : { content: [{ type: "text", text: "Conversation recovered after cancellation." }], stop_reason: "end_turn" } });
  });
  const prompt = page.getByRole("textbox", { name: "Ask AI", exact: true });
  await prompt.fill("Run a long calculation");
  await page.getByRole("button", { name: "Send", exact: true }).click();
  await expect(page.locator(".tool.running")).toContainText("SQL query");
  await page.waitForTimeout(250);
  await prompt.fill("Keep this draft after stopping");
  await prompt.press("Escape");
  await expect(prompt).toHaveValue("Keep this draft after stopping");
  await expect(page.locator(".agent-stopped")).toHaveText("Stopped", { timeout: 10_000 });
  await expect(page.locator(".tool.stopped")).toContainText("SQL query · stopped");
  expect(cancellationRequests).toBe(1);
  await expect(page.locator(".notice.error")).toHaveCount(0);
  await prompt.fill("Continue this conversation");
  await page.getByRole("button", { name: "Send", exact: true }).click();
  await expect(page.getByText("Conversation recovered after cancellation.")).toBeVisible();
  expect(cancellationRequests).toBe(3);
  await page.getByText("Thinking summary", { exact: true }).click();
  await expect(page.locator(".thinking-summary")).toContainText("Check the recovered calculation.");
  await expect(page.locator(".agent-tool-step").filter({ has: page.getByRole("region", { name: "AI query result" }) })).toContainText("SELECT 42 AS recovered");
  await expect(page.locator(".chat")).not.toContainText("private-signature");
  await expect(page.locator(".tool.done").last()).toContainText("SQL query");
  // Clarification pauses model work; results and editor handoff work in the narrow Office pane.
  await page.unroute("https://api.anthropic.com/v1/messages");
  let experienceRequests = 0;
  await page.route("https://api.anthropic.com/v1/messages", async route => {
    const content = [
      [{ type: "tool_use", id: "period", name: "ask_clarification", input: { question: "Which reporting period?", options: ["This year", "Last year"] } }],
      [{ type: "tool_use", id: "scoped", name: "run_sql", input: { sql: "SELECT 7 AS scoped_answer", scope: { dateRange: "Calendar 2025", filters: "USD", assumptions: "Illustrative example" } } }],
      [{ type: "text", text: "Scoped answer.\n```sql\nSELECT 7 AS scoped_answer\n```" }],
      [{ type: "tool_use", id: "stop-question", name: "ask_clarification", input: { question: "Which division?", options: ["North", "South"] } }],
    ][experienceRequests++];
    await route.fulfill({ json: { content, stop_reason: "end_turn" } });
  });
  await page.setViewportSize({ width: 300, height: 600 });
  await prompt.fill("Calculate a scoped answer"); await page.getByRole("button", { name: "Send", exact: true }).click();
  await expect(page.getByText("Which reporting period?", { exact: true })).toBeVisible();
  await prompt.fill("Calendar 2025, USD"); expect(experienceRequests).toBe(1);
  await page.getByRole("button", { name: "Continue", exact: true }).click();
  await expect(page.getByText("Scoped answer.", { exact: true })).toBeVisible();
  const card = page.getByRole("region", { name: "AI query result" }).last();
  await expect(card).toContainText("7");
  await card.getByText("Scope and SQL", { exact: true }).click();
  await expect(card).toContainText("Calendar 2025");
  await card.getByRole("button", { name: "Edit query", exact: true }).click();
  await expect(page.getByLabel("SQL query")).toHaveValue("SELECT 7 AS scoped_answer");
  await page.getByRole("tab", { name: "Ask AI", exact: true }).click();
  await card.getByRole("button", { name: "Insert into Excel", exact: true }).click();
  await page.getByRole("dialog").getByRole("button", { name: "Cancel", exact: true }).click();
  await expect(page.getByRole("dialog")).toHaveCount(0);
  await expect(page.locator(".notice.error")).toHaveCount(0);
  await prompt.fill("Another scoped question"); await page.getByRole("button", { name: "Send", exact: true }).click();
  await expect(page.getByText("Which division?", { exact: true })).toBeVisible();
  await page.getByRole("button", { name: "Stop", exact: true }).click();
  await expect(page.locator(".agent-stopped")).toHaveCount(2);
  await expect(page.getByRole("region", { name: "AI query result" })).toHaveCount(2);
  expect(experienceRequests).toBe(4);
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  await page.setViewportSize({ width: 1280, height: 720 });

  // Editor session state survives cancellation on the separate agent connection.
  await page.getByRole("tab", { name: "Query Editor", exact: true }).click();
  await page.getByLabel("SQL query").fill("SELECT after_cancel FROM cancellation_probe");
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await expect(results).toContainText("84");

});
