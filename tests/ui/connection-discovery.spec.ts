import { expect, test } from "@playwright/test";

test("desktop discovers catalogs without saving and preserves manual names and endpoints", async ({ page }) => {
  await page.setViewportSize({ width: 360, height: 480 });
  await page.addInitScript(() => {
    let values = [{ name: "Saved", catalog: "old", location: "https://saved.example.test", authentication: "anonymous", isDefault: true }];
    (window as any).__saved = values;
    Object.defineProperty(window, "chrome", { value: { webview: { addEventListener() {}, postMessage(request: any) {
      let result: unknown = true, error: string | undefined;
      if (request.method === "connections.list") result = values;
      if (request.method === "agent.key.load") result = null;
      if (request.method === "connections.catalogs") {
        const location = request.params.location;
        result = location.includes("multiple") ? ["weather", "quakes"] : ["weather"];
        if (location.includes("protected")) error = "IO Error: HTTP 401 no auth configured [url: https://protected.example.test/vgi.v2/catalog_catalogs]. Use bearer_token *** in ATTACH options.";
        if (location.includes("broken")) error = "Connection test timed out after 20 seconds.";
      }
      if (request.method === "connections.test") { result = { authentication: "anonymous" }; if (request.params.connection.location.includes("broken")) error = "Connection test failed."; }
      if (request.method === "connections.save" && request.params.connection.location.includes("save-error")) error = "Unable to save connection.";
      else if (request.method === "connections.save") {
        values = [...values.filter(item => item.name !== request.params.connection.name), request.params.connection];
        (window as any).__saved = values; result = values;
      }
      if (request.method === "connections.catalogs" && request.params.location.includes("protected")) {
        setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, progress: "Waiting for sign-in…" } as any), 0);
        setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result, error }), 1000);
        return;
      }
      setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result, error }), 0);
    } } }, configurable: true });
  });
  await page.goto("http://127.0.0.1:4173");
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await page.getByRole("button", { name: "New connection", exact: true }).click();
  await page.getByLabel("Server address").fill("https://single.example.de");
  await page.getByRole("button", { name: "Find catalogs" }).click();
  await expect(page.locator(".connection-form").getByLabel("Catalog", { exact: true })).toHaveValue("weather");
  await expect(page.getByLabel("Connection name", { exact: true })).toHaveValue("weather");
  await page.getByRole("button", { name: "Test connection", exact: true }).click();
  await expect(page.locator(".connection-status")).toContainText("Connection successful");
  expect(await page.evaluate(() => (window as any).__saved.length)).toBe(1);
  await page.getByLabel("Connection name", { exact: true }).fill("My data");
  await page.getByLabel("Server address").fill("https://multiple.example.de");
  await page.getByRole("button", { name: "Find catalogs" }).click();
  await expect(page.locator(".connection-form").getByLabel("Catalog", { exact: true })).toHaveValue("weather");
  await page.locator(".connection-form").getByLabel("Catalog", { exact: true }).selectOption("quakes");
  await expect(page.getByLabel("Connection name", { exact: true })).toHaveValue("My data");
  await page.getByRole("button", { name: "Enter catalog manually" }).click();
  await page.locator(".connection-form").getByLabel("Catalog", { exact: true }).fill("custom");
  await page.getByLabel("Connection name", { exact: true }).fill("saved");
  await expect(page.getByRole("button", { name: "Save changes" })).toBeDisabled();
  await expect(page.locator(".field-error")).toContainText("already exists");
  await page.getByLabel("Connection name", { exact: true }).fill("My data");
  await page.getByLabel("Server address").fill("https://protected.example.test");
  await page.getByRole("button", { name: "Find catalogs" }).click();
  await expect(page.locator(".connection-status")).toHaveText("Waiting for sign-in…");
  await expect(page.locator(".field-error")).toHaveText("Cupola couldn’t sign in to list catalogs. Try Find catalogs again, or enter the catalog name manually.");
  await expect(page.getByRole("button", { name: "Find catalogs" })).toBeFocused();
  await expect(page.locator(".notice.error")).toHaveCount(0);
  await page.getByLabel("Server address").fill("https://save-error.example.de");
  await page.getByRole("button", { name: "Save changes" }).click();
  await expect(page.locator(".field-error")).toContainText("Unable to save");
  await expect(page.locator(".notice.error")).toHaveCount(0);
  expect(await page.evaluate(() => (window as any).__saved.length)).toBe(1);
  await page.getByLabel("Server address").fill("https://broken.example.de");
  await page.getByRole("button", { name: "Find catalogs" }).click();
  await expect(page.locator(".field-error")).toContainText("enter the catalog name manually");
  await expect(page.locator(".notice.error")).toHaveCount(0);
  await expect(page.getByLabel("Server address")).toHaveValue("https://broken.example.de");
  await page.getByRole("button", { name: "Test connection", exact: true }).click();
  await expect(page.locator(".field-error")).toContainText("Connection test failed");
  await expect(page.locator(".notice.error")).toHaveCount(0);
  await page.getByRole("button", { name: "Save changes" }).click();
  await expect(page.locator(".connection-status")).toContainText("Connection saved");
  expect(await page.evaluate(() => (window as any).__saved.length)).toBe(2);
  await page.getByRole("button", { name: /^Saved / }).click();
  await page.getByRole("button", { name: "Find catalogs" }).click();
  await expect(page.getByLabel("Connection name", { exact: true })).toHaveValue("Saved");
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
});


test("Office connection save errors stay inline and preserve the draft", async ({ page }) => {
  await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
  await page.goto("http://127.0.0.1:4174/taskpane.html");
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await page.getByRole("button", { name: "New connection", exact: true }).click();
  await page.getByLabel("Server address").fill("https://example.de");
  await page.locator(".connection-form").getByLabel("Catalog", { exact: true }).fill("custom");
  await page.evaluate(() => {
    const original = Storage.prototype.setItem;
    Storage.prototype.setItem = function (key, value) {
      if (key === "vgi.excel.connections.v1") throw new Error("Storage is unavailable.");
      original.call(this, key, value);
    };
  });
  await page.getByRole("button", { name: "Save changes" }).click();
  await expect(page.locator(".field-error")).toContainText("Storage is unavailable");
  await expect(page.locator(".notice.error")).toHaveCount(0);
  await expect(page.getByLabel("Server address")).toHaveValue("https://example.de");
  await expect(page.getByRole("heading", { name: "New connection" })).toBeVisible();
  await page.locator(".connection-form").getByLabel("Catalog", { exact: true }).fill("changed");
  await expect(page.locator(".field-error")).toHaveCount(0);
});
