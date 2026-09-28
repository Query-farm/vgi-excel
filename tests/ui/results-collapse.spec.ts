import { expect, test } from "@playwright/test";

for (const host of ["desktop", "office"]) {
  test(`${host} can hide and restore results without losing the splitter or SQL`, async ({ page }) => {
    await page.setViewportSize({ width: host === "office" ? 300 : 360, height: 480 });
    await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
    await page.addInitScript(() => {
      const connection = { name: "test", catalog: "sample", location: "https://example.test", isDefault: true, authentication: "anonymous" };
      localStorage.setItem("vgi.excel.connections.v1", JSON.stringify([connection]));
      localStorage.setItem("vgi.excel.default-connection.v1", "test");
      const webview = { addEventListener() {}, postMessage(request: { id: number; method: string; params?: { sql?: string } }) {
        const result = request.method === "connections.list" ? [connection] : request.method === "query.editor" ? { columns: [{ name: "value", type: "INTEGER" }], rows: request.params?.sql === "SELECT empty" ? [] : Array.from({ length: 500 }, (_, i) => [i]), rowCount: request.params?.sql === "SELECT empty" ? 0 : 500, truncated: false } : request.method === "agent.key.load" ? null : true;
        setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
      }};
      Object.defineProperty(window, "chrome", { configurable: true, value: { webview } });
    });
    await page.goto(host === "office" ? "http://127.0.0.1:4174/taskpane.html" : "http://127.0.0.1:4173/index.html");
    const editor = page.getByLabel("SQL query");
    await editor.fill("SELECT 42");
    await expect(page.locator(".query-results-pane")).toHaveCount(0);
    await expect(page.getByRole("separator")).toHaveCount(0);
    expect((await editor.boundingBox())!.height).toBeGreaterThan(220);
    await page.screenshot({ path: `/tmp/cupola-empty-${host}.png` });
    if (host === "office") return; // Loaded Office results are covered by the real WASM suite.
    if (host === "desktop") {
      await page.getByRole("button", { name: "Run", exact: true }).click();
      await page.getByRole("button", { name: "Next preview page" }).click();
      await expect(page.getByRole("cell", { name: "200", exact: true })).toBeVisible();
    }
    const splitter = page.getByRole("separator", { name: "Resize query editor and results" });
    await splitter.focus(); await page.keyboard.press("ArrowDown");
    const originalSplit = await splitter.getAttribute("aria-valuenow");
    const original = (await editor.boundingBox())!;
    await page.getByRole("button", { name: "Hide results", exact: true }).click();
    const restore = page.getByRole("button", { name: "Show results", exact: true });
    await expect(restore).toBeFocused();
    await expect(restore).toHaveAttribute("aria-expanded", "false");
    await expect(splitter).toHaveCount(0);
    await expect(page.locator("#query-results-preview")).toBeHidden();
    expect((await editor.boundingBox())!.height).toBeGreaterThan(original.height + 40);
    const bar = (await restore.boundingBox())!;
    expect(bar.y + bar.height).toBeLessThanOrEqual(480);
    await page.screenshot({ path: `/tmp/cupola-collapsed-${host}.png` });
    await restore.click();
    await expect(page.getByRole("button", { name: "Hide results", exact: true })).toBeFocused();
    await expect(splitter).toHaveAttribute("aria-valuenow", originalSplit!);
    await expect(editor).toHaveValue("SELECT 42");
    if (host === "desktop") await expect(page.getByRole("cell", { name: "200", exact: true })).toBeVisible();
    if (host === "desktop") {
      const cell = (await page.getByRole("cell", { name: "200", exact: true }).boundingBox())!;
      const grid = (await page.locator(".query-results-content .result-wrap").boundingBox())!;
      expect(cell.y + cell.height).toBeLessThanOrEqual(grid.y + grid.height);
      expect(cell.y + cell.height).toBeLessThanOrEqual(480);
    }
    await page.screenshot({ path: `/tmp/cupola-expanded-${host}.png` });
    await page.getByRole("button", { name: "Hide results", exact: true }).click();
    if (host === "desktop") {
      await page.getByRole("button", { name: "Run", exact: true }).click();
      await expect(page.getByRole("status")).toHaveText("Query completed.");
      await expect(restore).toBeVisible();
      await expect(page.locator(".results-toolbar")).toContainText("500 rows");
    }
    await editor.fill("SELECT empty");
    await page.getByRole("button", { name: "Run", exact: true }).click();
    await expect(page.getByRole("status")).toHaveText("Query completed. No rows returned.");
    await expect(page.locator(".query-results-pane")).toHaveCount(0);
    await editor.fill("SELECT 42");
    await page.reload();
    await expect(page.locator(".query-results-pane")).toHaveCount(0);
    await page.getByRole("button", { name: "Run", exact: true }).click();
    await expect(page.getByRole("button", { name: "Show results", exact: true })).toBeVisible();
    await expect(editor).toHaveValue("SELECT 42");
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  });
}
