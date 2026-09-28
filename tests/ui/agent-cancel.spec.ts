import { expect, test } from "@playwright/test";

test("desktop agent stops a tool, rejects duplicate sends, and resumes without losing its draft", async ({ page }) => {
  await page.addInitScript(() => {
    (window as any).cancelCalls = 0;
    let pending: { id: number; params: { queryId: string } } | undefined;
    const reply = (id: number, result: unknown) => window.vgiReceiveHostResponse?.({ id, result });
    const empty = { columns: [], rows: [], rowCount: 0 };
    const webview = { addEventListener() {}, postMessage(request: { id: number; method: string; params: { sql?: string; queryId: string } }) {
      if (request.method === "query.agent" && request.params.sql === "SELECT slow") { pending = request; return; }
      if (request.method === "query.cancel") {
        (window as any).cancelCalls++;
        if (pending?.params.queryId !== request.params.queryId) throw new Error("Wrong query cancelled");
        const target = pending;
        setTimeout(() => { reply(target.id, null); reply(request.id, true); }, 250); return;
      }
      const result = request.method === "connections.list" ? [{ name: "test", catalog: "sample", location: "https://example.test", isDefault: true }] : request.method === "agent.key.load" ? "test-key" : request.method === "query.agent" ? empty : true;
      setTimeout(() => reply(request.id, result), 0);
    }};
    Object.defineProperty(window, "chrome", { configurable: true, value: { webview } });
  });
  let requests = 0;
  await page.route("https://api.anthropic.com/v1/messages", async route => {
    await route.fulfill({ json: requests++ === 0
      ? { content: [{ type: "tool_use", id: "slow", name: "run_sql", input: { sql: "SELECT slow" } }], stop_reason: "tool_use" }
      : { content: [{ type: "text", text: "Recovered successfully." }], stop_reason: "end_turn" } });
  });
  await page.goto("http://127.0.0.1:4173/index.html");
  await page.getByRole("tab", { name: "Ask AI", exact: true }).click();
  const prompt = page.getByRole("textbox", { name: "Ask AI", exact: true });
  await prompt.fill("Run a query");
  await page.getByRole("button", { name: "Send", exact: true }).click();
  await expect(page.locator(".tool.running")).toBeVisible();
  await prompt.fill("My next question");
  await prompt.press("Enter");
  await prompt.dispatchEvent("keydown", { key: "Escape", isComposing: true });
  await expect(page.locator(".tool.running")).toBeVisible();
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await page.keyboard.press("Escape");
  expect(await page.evaluate(() => (window as any).cancelCalls)).toBe(0);
  await page.getByRole("button", { name: "Back to workspace", exact: true }).click();
  await page.getByRole("tablist", { name: "AI conversations" }).getByRole("tab", { selected: true }).dblclick();
  await page.getByLabel("Rename Run a query").press("Escape");
  expect(await page.evaluate(() => (window as any).cancelCalls)).toBe(0);
  await prompt.press("Escape");
  await prompt.press("Escape");
  await expect(page.getByRole("button", { name: "Stopping…", exact: true })).toBeDisabled();
  await expect(page.locator(".agent-stopped")).toHaveText("Stopped");
  await expect(page.locator(".tool.stopped")).toContainText("SQL query · stopped");
  await expect(prompt).toHaveValue("My next question");
  await expect(page.locator(".notice.error")).toHaveCount(0);
  expect(requests).toBe(1);
  expect(await page.evaluate(() => (window as any).cancelCalls)).toBe(1);
  await page.getByRole("button", { name: "Send", exact: true }).click();
  await expect(page.getByText("Recovered successfully.")).toBeVisible();
  expect(requests).toBe(2);
});
