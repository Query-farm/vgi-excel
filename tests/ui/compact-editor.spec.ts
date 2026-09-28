import { expect, test } from "@playwright/test";

for (const width of [360, 600, 720, 1060]) {
  test(`desktop first results expose rows and bottom actions at ${width}px with a saved tall editor`, async ({ page }) => {
    await page.setViewportSize({ width, height: 480 });
    await page.addInitScript(() => {
      localStorage.setItem("cupola.query.editorPercent", "72");
      const connection = { name: "test", catalog: "sample", location: "https://example.test", isDefault: true };
      const webview = { addEventListener() {}, postMessage(request: { id: number; method: string }) {
        localStorage.setItem("cupola.test.method", request.method);
        const result = request.method === "connections.list" ? [connection] : request.method === "query.editor" ? { columns: [{ name: "value", type: "INTEGER" }], rows: Array.from({ length: 500 }, (_, i) => [i]), rowCount: 500, truncated: false } : request.method === "agent.key.load" ? null : true;
        setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
      }};
      Object.defineProperty(window, "chrome", { configurable: true, value: { webview } });
    });
    await page.goto("http://127.0.0.1:4173/index.html");
    await expect(page.locator("header .workspace-tabs")).toBeVisible();
    expect((await page.locator("header").boundingBox())!.height).toBeLessThanOrEqual(width < 700 ? 74 : 46);
    await page.getByRole("button", { name: "Run", exact: true }).click();
    const cell = page.getByRole("cell", { name: "2", exact: true });
    await expect(cell).toBeVisible();
    const grid = (await page.locator(".query-results-content .result-wrap").boundingBox())!;
    const third = (await cell.boundingBox())!;
    expect(third.y + third.height).toBeLessThanOrEqual(grid.y + grid.height);
    expect((await page.getByLabel("SQL query").boundingBox())!.height).toBeGreaterThanOrEqual(60);
    const actions = page.getByRole("toolbar", { name: "Result actions" });
    const bar = (await actions.boundingBox())!;
    expect(bar.y).toBeGreaterThanOrEqual(grid.y + grid.height - 1);
    expect(bar.y + bar.height).toBeLessThanOrEqual(480);
    await expect(actions.getByRole("button").last()).toHaveText("Load into Excel");
    const primary = (await actions.getByRole("button", { name: "Load into Excel" }).boundingBox())!;
    expect(primary.x + primary.width).toBeGreaterThan(width - 12);
    await expect(page.getByRole("separator")).toHaveAttribute("aria-valuenow", "72");
    await page.screenshot({ path: `/tmp/cupola-compact-desktop-${width}.png` });
    const more = page.getByText("More", { exact: true });
    await more.click();
    await expect(page.getByRole("button", { name: "Insert static table" })).toBeVisible();
    await page.getByRole("button", { name: "Insert static table" }).focus();
    await page.keyboard.press("Escape");
    await expect(more).toBeFocused();
    await expect(page.getByRole("button", { name: "Insert static table" })).toBeHidden();
    await more.click();
    await page.getByRole("button", { name: "Copy results", exact: true }).click();
    await expect.poll(() => page.evaluate(() => localStorage.getItem("cupola.test.method"))).toBe("clipboard.write");
    await expect(page.locator(".results-more")).not.toHaveAttribute("open", "");
    await more.click();
    await page.getByLabel("SQL query").click();
    await expect(page.getByRole("button", { name: "Copy results", exact: true })).toBeHidden();
  });
}
