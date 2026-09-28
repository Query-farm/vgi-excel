import { expect, test } from "@playwright/test";

for (const host of ["desktop", "office"] as const) {
  test(`${host} AI settings preserve models, persist effort, and copy tables at narrow widths`, async ({ page }) => {
    await page.setViewportSize({ width: host === "office" ? 300 : 360, height: 480 });
    await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
      await page.route("https://api.anthropic.com/v1/models?**", route => route.fulfill({ status: 503, body: "Unavailable" }));
    await page.route(/sentry\.io\//, route => route.abort());
    await page.addInitScript(({ host }) => {
      const connection = { name: "weather", catalog: "open_meteo", location: "https://example.com", authentication: "anonymous", isDefault: true };
      localStorage.setItem("vgi.excel.connections.v1", JSON.stringify([connection]));
      localStorage.setItem("vgi.excel.default-connection.v1", "weather");
      const storageKey = host === "office" ? "cupola.office.agent.conversations.v1::weather" : "cupola.agent.conversations.v1::weather";
      if (!localStorage.getItem(storageKey)) localStorage.setItem(storageKey, JSON.stringify({ version: 1, activeId: "chat", documents: [{ id: "chat", name: "Forecast", model: "claude-sonnet-4-6", draft: "", createdAt: 1, updatedAt: 1, displayMessages: [{ role: "user", text: "Weather?" }, { role: "assistant", modelId: "claude-sonnet-4-6", modelName: "Claude Sonnet 4.6", text: "| City | Temp |\n| --- | --- |\n| Paris | 21 |" }], agentMessages: [] }] }));
      Object.defineProperty(navigator, "clipboard", { configurable: true, value: { writeText: async (text: string) => { (window as any).__clipboard = text; } } });
      const webview = { postMessage(request: { id: number; method: string }) {
        const result = request.method === "connections.list" ? [connection] : request.method === "agent.key.load" ? "test-key" : { columns: [], rows: [], rowCount: 0 };
        setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
      }, addEventListener() {} };
      Object.defineProperty(window, "chrome", { value: { webview }, configurable: true });
    }, { host });
    await page.goto(host === "desktop" ? "http://127.0.0.1:4173" : "http://127.0.0.1:4174/taskpane.html");
    await page.getByRole("tab", { name: "Ask AI", exact: true }).click();
    await page.getByRole("button", { name: "Settings", exact: true }).click();
    await page.getByRole("tab", { name: "AI settings", exact: true }).click();
    await expect(page.getByLabel("Model", { exact: true })).toHaveValue("claude-sonnet-4-6");
    await page.getByLabel("Model", { exact: true }).selectOption("claude-sonnet-5");
    await page.getByLabel("Thinking effort").selectOption("low");
    await page.getByRole("tabpanel", { name: "AI settings", exact: true }).getByText("Advanced options", { exact: true }).click();
    await page.getByLabel("Anthropic workspace ID (optional)").fill("wrkspc_example");
    await page.getByRole("button", { name: "Back to workspace", exact: true }).click();
    await expect(page.getByRole("button", { name: "Send", exact: true })).toBeInViewport();
    await page.reload();
    await page.getByRole("tab", { name: "Ask AI", exact: true }).click();
    await page.getByRole("button", { name: "Settings", exact: true }).click();
    await page.getByRole("tab", { name: "AI settings", exact: true }).click();
    await expect(page.getByLabel("Thinking effort")).toHaveValue("low");
    await page.getByRole("tabpanel", { name: "AI settings", exact: true }).getByText("Advanced options", { exact: true }).click();
    await expect(page.getByLabel("Anthropic workspace ID (optional)")).toHaveValue("wrkspc_example");
    await page.getByLabel("Model", { exact: true }).selectOption("claude-haiku-4-5-20251001");
    await expect(page.getByLabel("Thinking effort")).toHaveCount(0);
    await page.getByRole("button", { name: "Back to workspace", exact: true }).click();
    await expect(page.locator(".message.assistant > strong")).toHaveText("Claude Sonnet 4.6");
    await page.getByRole("button", { name: "Copy table", exact: true }).click();
    await expect(page.getByRole("status").filter({ hasText: "Table copied" })).toBeVisible();
    expect(await page.evaluate(() => (window as any).__clipboard)).toBe("City\tTemp\nParis\t21");
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  });
}
