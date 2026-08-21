import { expect, test } from "@playwright/test";

const endpoint = process.env.CUPOLA_LIVE_VGI_ENDPOINT ?? "https://vgi-open-meteo.rusty-bb6.workers.dev";

test("self-hosted Haybarn WASM attaches an HTTPS VGI catalog and executes SQL", async ({ page }) => {
  const engineAssets = new Set<string>();
  page.on("response", (response) => {
    if (response.url().includes("/haybarn/")) engineAssets.add(new URL(response.url()).pathname.split("/").pop() ?? "");
  });
  await page.route("https://appsforoffice.microsoft.com/**", (route) => route.abort());
  await page.addInitScript(({ location }) => {
    localStorage.setItem("vgi.excel.connections.v1", JSON.stringify([{
      name: "open-meteo-live",
      catalog: "open_meteo",
      location,
      authentication: "anonymous",
      attachOptions: {},
    }]));
    localStorage.setItem("vgi.excel.default-connection.v1", "open-meteo-live");
  }, { location: endpoint });

  await page.goto("https://127.0.0.1:4184/taskpane.html");
  expect(await page.evaluate(() => globalThis.crossOriginIsolated)).toBe(true);
  await page.getByLabel("SQL query").fill("SELECT 42 AS wasm_answer, current_setting('TimeZone') AS local_time_zone;");
  await page.getByRole("button", { name: "Run", exact: true }).click();

  const results = page.getByRole("table", { name: "Query results" });
  await expect(results).toContainText("wasm_answer");
  await expect(results).toContainText("42");
  expect([...engineAssets].some((name) => name.endsWith(".wasm"))).toBe(true);
  expect([...engineAssets].some((name) => name.endsWith(".worker.js"))).toBe(true);
});
