import { expect, test } from "@playwright/test";

test("AI clarification, per-answer results, safe editor handoff and explicit Excel confirmation", async ({ page }) => {
  await page.setViewportSize({ width: 360, height: 600 });
  await page.addInitScript(() => {
    (window as any).hostCalls = [];
    Object.defineProperty(window, "chrome", { configurable: true, value: { webview: { addEventListener() {}, postMessage(request: any) {
      (window as any).hostCalls.push(request);
      const n = request.params?.sql === "SELECT 43 AS answer" ? 43 : 42;
      const result = request.method === "connections.list" ? [{ name: "test", catalog: "sample", location: "https://example.test", isDefault: true }] : request.method === "agent.key.load" ? "test-key" : request.method === "query.agent" ? { columns: [{ name: "answer", type: "INTEGER" }], rows: [[n]], rowCount: 1 } : true;
      setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
    } } } });
  });
  let requests = 0;
  let requestedModel = "";
  await page.route("https://api.anthropic.com/v1/messages", async route => {
    requestedModel = route.request().postDataJSON().model;
    const contents = [
      [{ type: "thinking", thinking: "I need the reporting period.", signature: "hidden-signature" }, { type: "redacted_thinking", data: "hidden-redacted" }, { type: "text", text: "Let me confirm the period first." }, { type: "tool_use", id: "clarify", name: "ask_clarification", input: { question: "Which period?", options: ["This year", "Last year"] } }],
      [42, 43].map(n => ({ type: "tool_use", id: `sql-${n}`, name: "run_sql", input: { sql: `SELECT ${n} AS answer`, query_name: `Annual revenue ${n}`, scope: { dateRange: "Last year", filters: "None", assumptions: "Illustrative calculation" } } })),
      [{ type: "text", text: "Two results.\n```sql\nSELECT 43 AS answer\n```" }],
      [{ type: "tool_use", id: "cancel-question", name: "ask_clarification", input: { question: "Which currency?", options: ["USD", "EUR"] } }],
    ];
    await route.fulfill({ json: { content: contents[requests++], stop_reason: "end_turn", usage: { input_tokens: 10, output_tokens: 4, cache_read_input_tokens: 90 } } });
  });
  await page.goto("http://127.0.0.1:4173/index.html");
  await page.getByLabel("SQL query").fill("SELECT original");
  await page.getByRole("tab", { name: "Ask AI", exact: true }).click();
  const prompt = page.getByRole("textbox", { name: "Ask AI", exact: true });
  await prompt.fill("Calculate revenue"); await page.getByRole("button", { name: "Send", exact: true }).click();
  await expect(page.getByText("Which period?", { exact: true })).toBeVisible();
  await expect(page.locator(".message.assistant > strong")).toHaveText(requestedModel);
  await page.getByText("Thinking summary", { exact: true }).click();
  await expect(page.locator(".thinking-summary")).toContainText("I need the reporting period.");
  await expect(page.locator(".chat")).not.toContainText("hidden-signature");
  await expect(page.locator(".chat")).not.toContainText("hidden-redacted");
  await expect(page.locator(".agent-progress")).toContainText("Waiting for your answer");
  await expect(page.locator(".agent-run-status").getByLabel("Active elapsed time")).toHaveCount(0);
  await expect(page.locator(".tool.running")).toContainText("Waiting for your answer");
  await page.screenshot({ path: "/tmp/cupola-ai-clarification.png" });
  await page.getByRole("button", { name: "Last year", exact: true }).click();
  expect(requests).toBe(1);
  await page.getByRole("button", { name: "Continue", exact: true }).click();
  const cards = page.getByRole("region", { name: "AI query result" });
  await expect(cards).toHaveCount(2);
  await expect(cards.nth(0)).toContainText("42"); await expect(cards.nth(1)).toContainText("43");
  await expect(page.locator(".agent-run-status")).toHaveCount(0);
  expect(await cards.nth(0).evaluate(el => {
    const step = el.closest(".agent-tool-step")!;
    const next = step.nextElementSibling!;
    return step.textContent!.includes("SELECT 42") && next.textContent!.includes("SELECT 43") && next.nextElementSibling!.textContent!.includes("Two results.");
  })).toBe(true);
  await cards.nth(1).scrollIntoViewIfNeeded();
  await page.screenshot({ path: "/tmp/cupola-ai-results.png" });
  await cards.nth(1).getByText("Scope and SQL", { exact: true }).click();
  await expect(cards.nth(1)).toContainText("AI-described scope");
  await expect(cards.nth(1)).toContainText("Last year");
  await cards.nth(1).getByRole("button", { name: "Edit query", exact: true }).click();
  await expect(page.getByLabel("SQL query")).toHaveValue("SELECT 43 AS answer");
  await expect(page.getByRole("tab", { name: "Annual revenue 43", exact: true })).toBeVisible();
  expect(await page.evaluate(() => (window as any).hostCalls.filter((c: any) => c.method === "query.editor").length)).toBe(0);
  await page.getByRole("tab", { name: "Ask AI", exact: true }).click();
  page.once("dialog", dialog => dialog.dismiss());
  await cards.nth(0).getByRole("button", { name: "Load into Excel", exact: true }).click();
  expect(await page.evaluate(() => (window as any).hostCalls.filter((c: any) => c.method === "excel.createPowerQuery").length)).toBe(0);
  page.once("dialog", async dialog => {
    expect(dialog.message()).toContain("Annual revenue 42");
    await dialog.accept();
  });
  await cards.nth(0).getByRole("button", { name: "Load into Excel", exact: true }).click();
  await expect.poll(() => page.evaluate(() => (window as any).hostCalls.find((c: any) => c.method === "excel.createPowerQuery")?.params)).toEqual({ sql: "SELECT 42 AS answer", connection: "test", name: "Annual revenue 42", loadToWorksheet: true });
  await cards.nth(1).getByRole("button", { name: "Open window", exact: true }).click();
  expect(await page.evaluate(() => (window as any).hostCalls.find((c: any) => c.method === "results.open").params.result.rows[0][0])).toBe(43);
  await page.locator(".markdown-code").getByRole("button", { name: "Open in Query Editor", exact: true }).click();
  await expect(page.getByLabel("SQL query")).toHaveValue("SELECT 43 AS answer");
  await page.getByRole("tab", { name: "Ask AI", exact: true }).click();
  await prompt.fill("Another question"); await page.getByRole("button", { name: "Send", exact: true }).click();
  await expect(page.getByText("Which currency?", { exact: true })).toBeVisible();
  await prompt.press("Escape");
  await expect(page.locator(".agent-stopped")).toHaveText("Stopped");
  await expect(cards).toHaveCount(2);
  expect(requests).toBe(4);
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
});
