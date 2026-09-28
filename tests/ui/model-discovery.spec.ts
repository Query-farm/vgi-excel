import { expect, test } from "@playwright/test";
for (const host of ["desktop", "office"]) {
  test(`${host} discovers models, caches them, retains selection and reports refresh failures inline`, async ({ page }) => {
    await page.setViewportSize({ width: host === "office" ? 300 : 360, height: 600 });
    await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
    await page.addInitScript(() => {
      sessionStorage.setItem("vgi.excel.anthropic-key", "test-model-key");
      Object.defineProperty(window, "chrome", { configurable: true, value: { webview: { addEventListener() {}, postMessage(request: { id: number; method: string }) {
        const result = request.method === "connections.list" ? [] : request.method === "agent.key.load" ? "test-model-key" : true;
        setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
      } } } });
    });
    let requests = 0;
    await page.route("https://api.anthropic.com/v1/models?**", async route => {
      requests++;
      if (requests > 1) { await route.fulfill({ status: 503, body: "secret-provider-error" }); return; }
      await route.fulfill({ json: { data: [{ id: "future-model", display_name: "Future model", max_tokens: 32000, capabilities: { thinking: { types: { adaptive: { supported: true } } }, effort: { supported: true, low: { supported: true }, high: { supported: true } } } }, { id: "test-opus", display_name: "Claude Opus" }, { id: "test-haiku", display_name: "Claude Haiku" }], has_more: false } });
    });
    await page.goto(host === "office" ? "http://127.0.0.1:4174/taskpane.html" : "http://127.0.0.1:4173/index.html");
    await page.getByRole("button", { name: "Settings", exact: true }).click();
    await page.getByRole("tab", { name: "AI settings", exact: true }).click();
    const model = page.getByLabel("Model", { exact: true });
    const original = await model.inputValue();
    await expect(model.locator('option[value="future-model"]')).toHaveCount(1);
    await expect(model).toHaveValue(original);
    await expect(model).toHaveJSProperty("tagName", "SELECT");
    await expect(model.locator("option")).toContainText([original, "Future model", "Claude Opus", "Claude Haiku", "Custom model ID…"]);
    await model.selectOption("test-opus");
    await expect(model).toHaveValue("test-opus");
    await model.selectOption("future-model");
    await expect(page.getByLabel("Thinking effort").locator("option")).toHaveCount(2);
    await page.getByRole("tab", { name: "About", exact: true }).click();
    await page.getByRole("tab", { name: "AI settings", exact: true }).click();
    await expect(model.locator('option[value="future-model"]')).toHaveCount(1);
    expect(requests).toBe(1);
    await page.getByRole("button", { name: "Refresh models", exact: true }).click();
    await expect(page.getByText(/Showing the saved list/)).toBeVisible();
    await expect(model).toHaveValue("future-model");
    await expect(model.locator('option[value="future-model"]')).toHaveCount(1);
    await expect(page.locator(".notice.error")).toHaveCount(0);
    await expect(page.locator("body")).not.toContainText("secret-provider-error");
    expect(await page.evaluate(() => localStorage.getItem("cupola.ai.models.v1"))).not.toContain("test-model-key");
    await model.selectOption({ label: "Custom model ID…" });
    await page.getByLabel("Custom model ID", { exact: true }).fill("my-manual-model");
    await page.getByRole("tab", { name: "About", exact: true }).click();
    await page.getByRole("tab", { name: "AI settings", exact: true }).click();
    await expect(page.getByLabel("Custom model ID", { exact: true })).toHaveValue("my-manual-model");
    await model.selectOption("test-haiku");
    await model.selectOption({ label: "Custom model ID…" });
    await expect(page.getByLabel("Custom model ID", { exact: true })).toHaveValue("my-manual-model");
    await page.getByRole("button", { name: "Refresh models", exact: true }).click();
    await expect(page.getByText(/Showing the saved list/)).toBeVisible();
    await expect(page.getByLabel("Custom model ID", { exact: true })).toHaveValue("my-manual-model");
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  });
}
