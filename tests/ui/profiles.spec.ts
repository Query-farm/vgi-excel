import { expect, test } from "@playwright/test";

test("desktop saves a multi-catalog profile and browses both catalogs at narrow width", async ({ page }) => {
  await page.setViewportSize({ width: 360, height: 480 });
  await page.addInitScript(() => {
    const singles = [
      { name: "Weather", catalog: "weather", location: "https://weather.example.test", authentication: "anonymous", isDefault: true },
      { name: "Earthquakes", catalog: "earthquakes", location: "https://quakes.example.test", authentication: "anonymous" },
    ];
    let values = JSON.parse(localStorage.getItem("profile-fixture") ?? "null") ?? singles;
    const webview = { addEventListener() {}, postMessage(request: any) {
      let result: unknown = true;
      if (request.method === "connections.list") result = values;
      if (request.method === "connections.save") {
        const value = request.params.connection;
        const catalogs = value.members.map((name: string) => values.find((item: any) => item.name === name).catalog);
        values = [...values.filter((item: any) => item.name !== value.name), { ...value, catalog: catalogs[0], catalogs }];
        localStorage.setItem("profile-fixture", JSON.stringify(values)); result = values;
      }
      if (request.method === "connections.use") {
        values = values.map((item: any) => ({ ...item, isDefault: item.name === request.params.name }));
        localStorage.setItem("profile-fixture", JSON.stringify(values)); result = values;
      }
      if (request.method === "agent.key.load") result = null;
      if (["query.run", "query.editor"].includes(request.method)) {
        (window as any).__lastSql = request.params.sql;
        result = { columns: ["catalog", "schema", "name", "object_type", "kind", "summary"].map(name => ({ name, type: "VARCHAR" })), rows: [["weather", "main", "forecast", "table", "table", ""], ["earthquakes", "main", "recent", "view", "view", ""]], rowCount: 2 };
      }
      setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
    } };
    Object.defineProperty(window, "chrome", { value: { webview }, configurable: true });
  });
  await page.goto("http://127.0.0.1:4173");
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await expect(page.getByLabel("Connection type")).toHaveCount(0);
  await page.getByRole("button", { name: "Combine saved connections", exact: true }).click();
  await page.getByLabel("Connection name", { exact: true }).fill("Research");
  await expect(page.getByLabel("Server address")).toHaveCount(0);
  await expect(page.locator(".field-error")).toHaveCount(0);
  await expect(page.getByRole("button", { name: "Save changes" })).toBeDisabled();
  await page.getByRole("checkbox", { name: "Weather" }).check();
  await expect(page.getByRole("button", { name: "Save changes" })).toBeDisabled();
  await page.getByRole("checkbox", { name: "Earthquakes" }).check();
  await expect(page.getByRole("button", { name: "Save changes" })).toBeEnabled();
  await page.getByRole("group", { name: "Connections to combine" }).scrollIntoViewIfNeeded();
  await page.screenshot({ path: "/tmp/cupola-ux-desktop-combined.png" });
  await page.getByRole("tabpanel", { name: "Connections", exact: true }).getByText("Advanced options", { exact: true }).click();
  await page.getByLabel("Default connection").selectOption("Earthquakes");
  await page.getByRole("button", { name: "Save changes" }).click();
  await expect(page.getByRole("status").filter({ hasText: "Connection saved" })).toBeVisible();
  await page.getByRole("button", { name: "Use as active" }).click();
  await page.reload();
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await expect(page.getByRole("group", { name: "Connections to combine" })).toBeVisible();
  await page.getByRole("tabpanel", { name: "Connections", exact: true }).getByText("Advanced options", { exact: true }).click();
  await expect(page.getByLabel("Default connection")).toHaveValue("Earthquakes");
  await page.getByRole("tab", { name: "Catalog View", exact: true }).click();
  await expect(page.getByRole("tree", { name: "weather catalog objects" })).toBeVisible();
  await expect(page.getByRole("tree", { name: "earthquakes catalog objects" })).toBeVisible();
  expect(await page.evaluate(() => (window as any).__lastSql)).toContain("IN ('earthquakes','weather')");
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
});

for (const count of [0, 1]) {
  test(`ordinary connection setup with ${count} saved connections offers no type decision`, async ({ page }) => {
    await page.setViewportSize({ width: 360, height: 480 });
    await page.addInitScript(count => {
      const values = count ? [{ name: "Weather", catalog: "weather", location: "https://weather.example.test", authentication: "anonymous", isDefault: true }] : [];
      Object.defineProperty(window, "chrome", { value: { webview: { addEventListener() {}, postMessage(request: any) {
        const result = request.method === "connections.list" ? values : request.method === "agent.key.load" ? null : true;
        setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
      } } }, configurable: true });
    }, count);
    await page.goto("http://127.0.0.1:4173");
    await page.getByRole("button", { name: "Settings", exact: true }).click();
    await page.getByRole("button", { name: "New connection", exact: true }).click();
    await expect(page.getByLabel("Connection type")).toHaveCount(0);
    await expect(page.getByRole("button", { name: "Combine saved connections" })).toHaveCount(0);
    await expect(page.getByLabel("Server address")).toBeVisible();
    await expect(page.locator(".field-error")).toHaveCount(0);
    expect(await page.locator(".connection-form input").first().getAttribute("placeholder")).toBe("https://vgi.example.com/");
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
    if (!count) await page.screenshot({ path: "/tmp/cupola-ux-desktop-single.png" });
  });
}

test("an existing one-member combination remains editable and New connection resets the form", async ({ page }) => {
  await page.addInitScript(() => {
    const values = [{ name: "Weather", catalog: "weather", location: "https://weather.example.test", authentication: "anonymous" }, { name: "Existing group", catalog: "weather", location: "", authentication: "anonymous", members: ["Weather"], isDefault: true }];
    Object.defineProperty(window, "chrome", { value: { webview: { addEventListener() {}, postMessage(request: any) {
      const result = request.method === "connections.list" ? values : request.method === "agent.key.load" ? null : true;
      setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
    } } }, configurable: true });
  });
  await page.goto("http://127.0.0.1:4173");
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await expect(page.getByRole("group", { name: "Connections to combine" })).toBeVisible();
  await expect(page.getByRole("checkbox", { name: /Weather/ })).toBeChecked();
  await expect(page.getByRole("button", { name: "Save changes" })).toBeEnabled();
  await expect(page.getByRole("button", { name: "Combine saved connections" })).toHaveCount(0);
  await page.getByRole("button", { name: "New connection", exact: true }).click();
  await expect(page.getByRole("group", { name: "Connections to combine" })).toHaveCount(0);
  await expect(page.getByLabel("Server address")).toHaveValue("");
  await expect(page.getByLabel("Connection name", { exact: true })).toHaveValue("");
});

test("Office uses the same simple connection labels at 300 pixels", async ({ page }) => {
  await page.setViewportSize({ width: 300, height: 480 });
  await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
  await page.goto("http://127.0.0.1:4174/taskpane.html");
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await page.getByRole("button", { name: "New connection", exact: true }).click();
  await expect(page.getByLabel("Server address")).toBeVisible();
  await expect(page.locator(".connection-form").getByLabel("Catalog", { exact: true })).toBeVisible();
  await expect(page.getByLabel("Connection type")).toHaveCount(0);
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  await page.screenshot({ path: "/tmp/cupola-ux-office-single.png" });
});
