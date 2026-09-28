import { expect, test } from "@playwright/test";

for (const host of ["desktop", "office"]) {
  for (const connected of [false, true]) {
    test(`${host} unified settings preserve AI drafts with ${connected ? "a connection" : "no connection"}`, async ({ page }) => {
      await page.setViewportSize({ width: host === "office" ? 300 : 360, height: 480 });
      await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
      await page.route("https://api.anthropic.com/v1/models?**", route => route.fulfill({ status: 503, body: "Unavailable" }));
      await page.addInitScript(({ connected }) => {
        const connections = connected ? [{ name: "test", catalog: "sample", location: "https://example.test", isDefault: true, authentication: "anonymous" }] : [];
        localStorage.setItem("vgi.excel.connections.v1", JSON.stringify(connections));
        if (connected) localStorage.setItem("vgi.excel.default-connection.v1", "test");
        const webview = { addEventListener() {}, postMessage(request: { id: number; method: string }) {
          const result = request.method === "connections.list" ? connections : request.method === "agent.key.load" ? null : true;
          setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
        }};
        Object.defineProperty(window, "chrome", { configurable: true, value: { webview } });
      }, { connected });
      await page.goto(host === "office" ? "http://127.0.0.1:4174/taskpane.html" : "http://127.0.0.1:4173/index.html");
      await page.getByRole("tab", { name: "Ask AI", exact: true }).click();
      await page.getByRole("textbox", { name: "Ask AI", exact: true }).fill("Keep this draft");
      await page.getByRole("button", { name: "Configure AI", exact: true }).click();
      await expect(page.getByRole("heading", { name: "Settings", exact: true })).toBeVisible();
      await expect(page.getByRole("tab", { name: "AI settings", exact: true })).toHaveAttribute("aria-selected", "true");
      await page.getByLabel("Anthropic API key", { exact: true }).fill("test-only-key");
      const keyInput = page.getByLabel("Anthropic API key", { exact: true });
      await keyInput.focus();
      const focusBounds = await keyInput.evaluate(input => {
        const rect = input.getBoundingClientRect();
        const panel = input.closest(".settings-content")!.getBoundingClientRect();
        const style = getComputedStyle(input);
        const ring = parseFloat(style.outlineWidth) + parseFloat(style.outlineOffset);
        return { left: rect.left - ring - panel.left, right: panel.right - rect.right - ring };
      });
      expect(focusBounds.left).toBeGreaterThanOrEqual(0);
      expect(focusBounds.right).toBeGreaterThanOrEqual(0);
      if (host === "desktop") {
        await page.getByRole("button", { name: "Save securely" }).click();
        await expect(page.getByText(/Protected by Windows Credential Manager/)).toBeVisible();
      }
      await page.getByLabel("Thinking effort").selectOption("low");
      await expect.poll(() => page.locator("header .mark").evaluate((el: HTMLImageElement) => el.naturalWidth)).toBeGreaterThan(0);
      await page.screenshot({ path: `/tmp/cupola-settings-ai-${host}.png` });
      await page.getByRole("tab", { name: "About", exact: true }).click();
      await expect(page.getByRole("dialog")).toHaveCount(0);
      await expect(page.getByRole("tabpanel", { name: "About", exact: true })).toContainText("Version 0.5.0");
      await expect.poll(() => page.locator(".about-content .about-mark").evaluate((el: HTMLImageElement) => el.naturalWidth)).toBeGreaterThan(0);
      await page.screenshot({ path: `/tmp/cupola-settings-about-${host}.png` });
      await page.getByRole("tab", { name: "AI settings", exact: true }).click();
      await expect(page.getByLabel("Anthropic API key", { exact: true })).toHaveValue("test-only-key");
      await expect(page.getByLabel("Thinking effort")).toHaveValue("low");
      await page.getByRole("button", { name: "Back to workspace", exact: true }).click();
      await expect(page.getByRole("textbox", { name: "Ask AI", exact: true })).toHaveValue("Keep this draft");
      await expect(page.getByRole("button", { name: "AI settings", exact: true })).toHaveCount(0);
      expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
    });
  }
}
