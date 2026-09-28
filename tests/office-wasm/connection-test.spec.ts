import { expect, test } from "@playwright/test";

const endpoint = process.env.CUPOLA_LIVE_VGI_ENDPOINT ?? "https://vgi-open-meteo.rusty-bb6.workers.dev";

test("a stalled connection test times out, preserves the URL, and can be retried", async ({ page, context }) => {
  await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
  // Keep the real WASM network request pending to exercise worker disposal,
  // rather than mocking BrowserBackend or only racing an artificial Promise.
  let heldRequests = 0;
  await context.route("https://cupola-unreachable.invalid/**", () => { heldRequests++; });
  await page.addInitScript(() => {
    localStorage.setItem("vgi.excel.connections.v1", JSON.stringify([{
      name: "timeout-probe", catalog: "open_meteo",
      location: "https://cupola-unreachable.invalid", authentication: "anonymous", attachOptions: {},
    }]));
    localStorage.setItem("vgi.excel.default-connection.v1", "timeout-probe");
  });
  await page.goto("https://127.0.0.1:4184/taskpane.html");
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await page.getByRole("button", { name: "Test connection", exact: true }).click();
  await expect(page.locator(".connection-form .field-error")).toContainText("timed out after 20 seconds", { timeout: 25_000 });
  expect(heldRequests).toBeGreaterThan(0);
  await expect(page.locator(".notice.error")).toHaveCount(0);
  await expect(page.getByRole("button", { name: "Test connection", exact: true })).toBeEnabled();
  await expect(page.getByLabel("Server address")).toHaveValue("https://cupola-unreachable.invalid");
  await page.getByLabel("Server address").fill(endpoint);
  await page.getByRole("button", { name: "Test connection", exact: true }).click();
  await expect(page.locator(".connection-status")).toContainText("Connected · no sign-in required.", { timeout: 25_000 });
});

test("the reported misspelled hostname produces an error without changing the address", async ({ page }) => {
  await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
  const mistyped = "https://vgi-open-meteo.rusty-bb6.workers.de";
  await page.addInitScript(location => {
    localStorage.setItem("vgi.excel.connections.v1", JSON.stringify([{
      name: "hostname-probe", catalog: "open_meteo", location, authentication: "anonymous", attachOptions: {},
    }]));
    localStorage.setItem("vgi.excel.default-connection.v1", "hostname-probe");
  }, mistyped);
  await page.goto("https://127.0.0.1:4184/taskpane.html");
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await page.getByRole("button", { name: "Test connection", exact: true }).click();
  await expect(page.locator(".connection-form .field-error")).toBeVisible({ timeout: 25_000 });
  await expect(page.getByRole("button", { name: "Test connection", exact: true })).toBeEnabled();
  await expect(page.getByLabel("Server address")).toHaveValue(mistyped);
  await expect(page.locator(".notice.error")).toHaveCount(0);
});

test("URL-first discovery defaults a new name, supports overrides, and saves only explicitly", async ({ page }) => {
  await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
  await page.goto("https://127.0.0.1:4184/taskpane.html");
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await page.getByRole("button", { name: "New connection", exact: true }).click();
  const stored = await page.evaluate(() => localStorage.getItem("vgi.excel.connections.v1"));
  await page.getByLabel("Server address").fill(endpoint);
  await page.getByRole("button", { name: "Find catalogs" }).click();
  await expect(page.locator(".connection-form").getByLabel("Catalog", { exact: true })).toHaveValue("open_meteo", { timeout: 25_000 });
  await expect(page.getByLabel("Connection name", { exact: true })).toHaveValue("open_meteo");
  await page.getByLabel("Connection name", { exact: true }).fill("My weather");
  await page.getByRole("button", { name: "Test connection", exact: true }).click();
  await expect(page.locator(".connection-status")).toContainText("Connected", { timeout: 25_000 });
  expect(await page.evaluate(() => localStorage.getItem("vgi.excel.connections.v1"))).toBe(stored);
  await page.getByRole("button", { name: "Find catalogs" }).click();
  await expect(page.locator(".connection-status")).toContainText("catalog available", { timeout: 25_000 });
  await expect(page.getByLabel("Connection name", { exact: true })).toHaveValue("My weather");
  await page.getByRole("button", { name: "Enter catalog manually" }).click();
  await page.locator(".connection-form").getByLabel("Catalog", { exact: true }).fill("custom");
  await expect(page.getByLabel("Connection name", { exact: true })).toHaveValue("My weather");
  await page.getByRole("button", { name: "Save changes" }).click();
  await expect(page.locator(".connection-status")).toContainText("Connection saved");
  const saved = await page.evaluate(() => JSON.parse(localStorage.getItem("vgi.excel.connections.v1")!));
  expect(saved.find((value: any) => value.name === "My weather")).toMatchObject({ catalog: "custom", location: endpoint });
});

test("incomplete sign-in explains retry and manual entry without engine details or a duplicate notice", async ({ page, context }) => {
  await page.setViewportSize({ width: 300, height: 480 });
  await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
  await context.route("https://protected-catalogs.invalid/**", route => route.fulfill({ status: 401, headers: { "access-control-allow-origin": "*" }, body: "Authentication required" }));
  await page.goto("https://127.0.0.1:4184/taskpane.html");
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await page.getByRole("button", { name: "New connection", exact: true }).click();
  await page.getByLabel("Server address").fill("https://protected-catalogs.invalid");
  await page.getByRole("button", { name: "Find catalogs" }).click();
  const error = page.locator(".connection-form .field-error");
  await expect(error).toHaveText("Sign-in wasn’t completed. Try Find catalogs again, or enter the catalog name manually.", { timeout: 25_000 });
  await expect(page.locator(".notice.error")).toHaveCount(0);
  await expect(page.getByRole("button", { name: "Find catalogs" })).toBeFocused();
  await page.screenshot({ path: "/tmp/cupola-protected-discovery-office.png" });
  await page.locator(".connection-form").getByLabel("Catalog", { exact: true }).fill("provided_catalog");
  await expect(error).toHaveCount(0);
  await expect(page.getByRole("button", { name: "Test connection", exact: true })).toBeEnabled();
  await expect(page.getByLabel("Server address")).toHaveValue("https://protected-catalogs.invalid");
});


test("protected discovery reuses the Office session credential without saving a connection", async ({ page, context }) => {
  const location = "https://cupola-authenticated.invalid";
  let authenticated = 0;
  await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
  await context.route(`${location}/**`, async route => {
    const request = route.request();
    if (new URL(request.url()).pathname.endsWith("/catalog_catalogs")) {
      expect(request.headers().authorization).toBe("Bearer cupola-discovery-test-only");
      authenticated++;
    }
    const headers = { ...request.headers() };
    delete headers.authorization;
    const response = await route.fetch({ url: endpoint + new URL(request.url()).pathname, headers });
    await route.fulfill({ response });
  });
  await page.addInitScript(location => {
    sessionStorage.setItem(`vgi.excel.oauth.${location}`, JSON.stringify({ access_token: "cupola-discovery-test-only" }));
  }, location);
  await page.goto("https://127.0.0.1:4184/taskpane.html");
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await page.getByRole("button", { name: "New connection", exact: true }).click();
  const before = await page.evaluate(() => localStorage.getItem("vgi.excel.connections.v1"));
  await page.getByLabel("Server address").fill(location);
  await page.getByRole("button", { name: "Find catalogs" }).click();
  await expect(page.locator(".connection-form").getByLabel("Catalog", { exact: true })).toHaveValue("open_meteo", { timeout: 25_000 });
  expect(authenticated).toBeGreaterThan(0);
  expect(await page.evaluate(() => localStorage.getItem("vgi.excel.connections.v1"))).toBe(before);
  await expect(page.locator(".field-error")).toHaveCount(0);
});
