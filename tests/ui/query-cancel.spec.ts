import { expect, test } from "@playwright/test";

for (const completionWins of [false, true]) {
  test(`desktop query cancellation ${completionWins ? "handles completion race" : "preserves results and permits retry"}`, async ({ page }) => {
    await page.addInitScript(({ completionWins }) => {
      let pending: { id: number; queryId: string } | undefined;
      const result = { columns: [{ name: "answer", type: "INTEGER" }], rows: [[42]], rowCount: 1, truncated: false };
      const reply = (id: number, value: unknown) => window.vgiReceiveHostResponse?.({ id, result: value });
      const webview = { addEventListener() {}, postMessage(request: { id: number; method: string; params?: { queryId?: string; sql?: string } }) {
        if (request.method === "connections.list") return setTimeout(() => reply(request.id, [{ name: "test", catalog: "test", location: "https://example.com", isDefault: true }]), 0);
        if (request.method === "query.editor") {
          if (request.params?.sql === "SELECT slow") { pending = { id: request.id, queryId: request.params.queryId! }; return; }
          return setTimeout(() => reply(request.id, result), 0);
        }
        if (request.method === "query.cancel") {
          const target = pending;
          if (!target || target.queryId !== request.params?.queryId) throw new Error("Wrong cancellation target");
          setTimeout(() => { reply(request.id, true); reply(target.id, completionWins ? result : null); pending = undefined; }, 100);
          return;
        }
        setTimeout(() => reply(request.id, request.method === "agent.key.load" ? null : true), 0);
      }};
      Object.defineProperty(window, "chrome", { configurable: true, value: { webview } });
    }, { completionWins });
    await page.setViewportSize({ width: 360, height: 480 });
    await page.emulateMedia({ reducedMotion: completionWins ? "reduce" : "no-preference" });
    await page.goto("http://127.0.0.1:4173/index.html");
    await page.getByLabel("SQL query").fill("SELECT 42");
    await page.getByRole("button", { name: "Run", exact: true }).click();
    await expect(page.getByRole("table", { name: "Query results" })).toContainText("42");
    await page.getByLabel("SQL query").fill("SELECT slow");
    await page.getByRole("button", { name: "Run", exact: true }).click();
    const spinner = page.getByRole("button", { name: "Cancel query" }).locator(".busy-spinner");
    await expect(spinner).toBeVisible();
    await expect(page.locator(".query-running-status")).toHaveText("Running query…");
    const progress = await page.locator(".query-run-progress").boundingBox();
    expect(progress!.x + progress!.width).toBeLessThanOrEqual(360);
    if (completionWins) await expect(spinner).toHaveCSS("animation-name", "none");
    else {
      const transform = await spinner.evaluate(el => getComputedStyle(el).transform);
      await expect.poll(() => spinner.evaluate(el => getComputedStyle(el).transform)).not.toBe(transform);
    }
    // Ctrl+Enter during execution must not launch a second query.
    await page.getByLabel("SQL query").press("Control+Enter");
    await page.getByRole("button", { name: "Cancel query" }).click();
    await expect(page.getByRole("status")).toHaveText(completionWins ? "Query completed." : "Query cancelled.");
    await expect(page.getByLabel("SQL query")).toHaveValue("SELECT slow");
    await expect(page.getByRole("table", { name: "Query results" })).toContainText("42");
    await expect(page.locator(".notice.error")).toHaveCount(0);
    await page.getByLabel("SQL query").fill("SELECT 42");
    await page.getByRole("button", { name: "Run", exact: true }).click();
    await expect(page.getByRole("status")).toHaveText("Query completed.");
  });
}
