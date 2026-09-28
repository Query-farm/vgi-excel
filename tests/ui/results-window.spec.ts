import { expect, test } from "@playwright/test";

for (const host of ["office", "desktop"]) {
  test(`${host} results window pages loaded rows and renders cell text safely`, async ({ page, context }) => {
    const origin = host === "office" ? "http://127.0.0.1:4174" : "http://127.0.0.1:4173";
    await page.goto(`${origin}/results.html`);
    const token = "11111111-1111-4111-8111-111111111111";
    await page.evaluate(token => {
      const channel = new BroadcastChannel(`cupola-results-${token}`);
      channel.onmessage = event => {
        if (event.data.type === "ready") channel.postMessage({ type: "snapshot", snapshot: {
          title: "Monthly report", result: { columns: [{ name: "row", type: "INTEGER" }, { name: "text", type: "VARCHAR" }], rows: Array.from({ length: 450 }, (_, i) => [i, i === 0 ? '<img src=x onerror="window.bad=true">' : "long cell ".repeat(100)]), rowCount: 900, truncated: true },
        }});
        if (event.data.type === "received") channel.close();
      };
    }, token);
    const viewer = await context.newPage();
    await viewer.goto(`${origin}/results.html#${token}`);
    await expect(viewer.getByRole("heading", { name: "Monthly report — Results" })).toBeVisible();
    await expect(viewer.getByText("450 of 900 rows loaded.", { exact: false })).toBeVisible();
    await expect(viewer.getByRole("table", { name: "Query results" }).locator("tbody tr")).toHaveCount(200);
    await expect(viewer.getByRole("table").locator("img")).toHaveCount(0);
    await expect(viewer.getByRole("cell", { name: '<img src=x onerror="window.bad=true">', exact: true })).toBeVisible();
    await viewer.getByRole("button", { name: "Next page" }).click();
    await expect(viewer.getByRole("cell", { name: "200", exact: true })).toBeVisible();
    await viewer.getByRole("button", { name: "Next page" }).click();
    await expect(viewer.getByRole("table").locator("tbody tr")).toHaveCount(50);
    await expect(viewer.getByRole("button", { name: "Next page" })).toBeDisabled();
    await viewer.getByLabel("Rows per page").selectOption("500");
    await expect(viewer.getByRole("table").locator("tbody tr")).toHaveCount(450);
    await viewer.setViewportSize({ width: 360, height: 480 });
    expect(await viewer.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
    expect(await viewer.locator(".viewer-grid").evaluate(el => el.scrollWidth > el.clientWidth)).toBe(true);
    await viewer.screenshot({ path: `/tmp/cupola-results-${host}-narrow.png` });
    await viewer.setViewportSize({ width: 1280, height: 850 });
    await viewer.screenshot({ path: `/tmp/cupola-results-${host}-wide.png` });
    expect(await viewer.evaluate(() => Object.values(localStorage).some(value => value.includes("Monthly report")))).toBe(false);
    await viewer.close();
  });
}

test("desktop results button sends a snapshot without rerunning SQL", async ({ page }) => {
  await page.addInitScript(() => {
    const result = { columns: [{ name: "answer", type: "INTEGER" }], rows: [[42]], rowCount: 1, truncated: false };
    const webview = { addEventListener() {}, postMessage(request: { id: number; method: string; params: unknown }) {
      const calls = JSON.parse(sessionStorage.getItem("calls") ?? "[]"); calls.push(request); sessionStorage.setItem("calls", JSON.stringify(calls));
      const value = request.method === "connections.list" ? [{ name: "test", catalog: "sample", location: "https://example.test", isDefault: true }] : request.method === "query.editor" ? result : request.method === "agent.key.load" ? null : true;
      setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result: value }), 0);
    }};
    Object.defineProperty(window, "chrome", { configurable: true, value: { webview } });
  });
  await page.goto("http://127.0.0.1:4173/index.html");
  await page.getByLabel("SQL query").fill("SELECT 42");
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await page.getByRole("button", { name: "Open results in new window" }).click();
  const calls = await page.evaluate(() => JSON.parse(sessionStorage.getItem("calls")!));
  expect(calls.filter((call: { method: string }) => call.method === "query.editor")).toHaveLength(1);
  expect(calls.find((call: { method: string }) => call.method === "results.open").params).toEqual({ title: "Query 1", result: { columns: [{ name: "answer", type: "INTEGER" }], rows: [[42]], rowCount: 1, truncated: false } });
});

for (const host of ["office", "desktop"]) {
  test(`${host} empty results retains column headings`, async ({ page }) => {
    const base = host === "office" ? "http://127.0.0.1:4174" : "http://127.0.0.1:4173";
    await page.addInitScript(() => {
      const RealChannel = BroadcastChannel;
      const source = new RealChannel("cupola-results-22222222-2222-4222-8222-222222222222");
      source.onmessage = event => { if (event.data.type === "ready") source.postMessage({ type: "snapshot", snapshot: { title: "Empty", result: { columns: [{ name: "expected_column", type: "VARCHAR" }], rows: [], rowCount: 0, truncated: false } } }); else source.close(); };
    });
    await page.goto(`${base}/results.html#22222222-2222-4222-8222-222222222222`);
    await expect(page.getByRole("columnheader", { name: /expected_column/ })).toBeVisible();
    await expect(page.getByText("This query returned no rows.")).toBeVisible();
    await expect(page.getByRole("button", { name: "Next page" })).toBeDisabled();
  });
}
