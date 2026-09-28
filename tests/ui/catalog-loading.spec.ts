import { expect, test } from "@playwright/test";

for (const fail of [false, true]) {
  test(`desktop catalog loading has spacing and stops on ${fail ? "failure" : "success"}`, async ({ page }) => {
    await page.setViewportSize({ width: 360, height: 480 });
    await page.addInitScript(({ fail }) => {
      const connection = { name: "test", catalog: "sample", location: "https://example.test", isDefault: true };
      const webview = { addEventListener() {}, postMessage(request: { id: number; method: string }) {
        if (request.method === "query.run") {
          (window as unknown as { finishCatalog: () => void }).finishCatalog = () => window.vgiReceiveHostResponse?.(fail
            ? { id: request.id, error: "Catalog unavailable. Try again." }
            : { id: request.id, result: { columns: [{ name: "catalog", type: "VARCHAR" }, { name: "schema", type: "VARCHAR" }, { name: "name", type: "VARCHAR" }, { name: "object_type", type: "VARCHAR" }], rows: [["sample", "main", "weather", "table"]], rowCount: 1, truncated: false } });
          return;
        }
        const result = request.method === "connections.list" ? [connection] : request.method === "agent.key.load" ? null : true;
        setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
      }};
      Object.defineProperty(window, "chrome", { configurable: true, value: { webview } });
    }, { fail });
    await page.goto("http://127.0.0.1:4173/index.html");
    await page.getByRole("tab", { name: "Catalog View" }).click();
    const loading = page.locator(".catalog-loading");
    await expect(loading).toHaveText("Loading catalog…");
    await expect(loading).toHaveAttribute("role", "status");
    await expect(loading).toHaveCSS("padding", "16px");
    await expect(loading.locator("svg")).toHaveCSS("animation-name", "cupola-spin");
    const firstTransform = await loading.locator("svg").evaluate(el => getComputedStyle(el).transform);
    await expect.poll(() => loading.locator("svg").evaluate(el => getComputedStyle(el).transform)).not.toBe(firstTransform);
    await expect(page.locator(".catalog-tree-toolbar").getByRole("button", { name: "Refresh", exact: true })).toBeDisabled();
    await expect(page.getByText("Select a catalog object", { exact: true })).toHaveCount(0);
    await expect(page.getByRole("button", { name: "Refresh", exact: true })).toHaveCSS("flex-shrink", "0");
    await page.screenshot({ path: "/tmp/cupola-catalog-loading-desktop.png" });
    await page.emulateMedia({ reducedMotion: "reduce" });
    await expect(loading.locator("svg")).toHaveCSS("animation-name", "none");
    await page.evaluate(() => (window as unknown as { finishCatalog: () => void }).finishCatalog());
    await expect(loading).toHaveCount(0);
    if (fail) await expect(page.locator(".notice.error")).toContainText("Catalog unavailable");
    else await expect(page.getByRole("tree")).toBeVisible();
  });
}

test("Office catalog loading stops on an engine startup failure", async ({ page }) => {
  await page.setViewportSize({ width: 300, height: 480 });
  await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
  let release!: () => void;
  const blocked = new Promise<void>(resolve => { release = resolve; });
  await page.route("**/haybarn/*.worker.js", async route => { await blocked; await route.fulfill({ contentType: "application/javascript", body: `self.onmessage = ({data}) => self.postMessage({messageId: 0, requestId: data.messageId, type: "ERROR", data: {name: "Error", message: "Engine startup failed."}});` }); });
  await page.addInitScript(() => {
    localStorage.setItem("vgi.excel.connections.v1", JSON.stringify([{ name: "test", catalog: "sample", location: "https://example.test", authentication: "anonymous" }]));
    localStorage.setItem("vgi.excel.default-connection.v1", "test");
  });
  await page.goto("http://127.0.0.1:4174/taskpane.html");
  await page.getByRole("tab", { name: "Catalog View" }).click();
  const loading = page.locator(".catalog-loading");
  await expect(loading).toHaveText("Loading catalog…");
  await expect(loading).toHaveCSS("padding", "16px");
  await expect(loading.locator("svg")).toHaveCSS("animation-name", "cupola-spin");
  await page.screenshot({ path: "/tmp/cupola-catalog-loading-office.png" });
  release();
  await expect(loading).toHaveCount(0);
  await expect(page.locator(".notice.error")).toBeVisible();
});
