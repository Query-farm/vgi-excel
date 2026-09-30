import { expect, test } from "@playwright/test";

test("workspace attaches multiple catalogs and preserves the legacy default", async ({ page }) => {
  await page.setViewportSize({ width: 360, height: 480 });
  await page.addInitScript(() => {
    let values = JSON.parse(localStorage.getItem("workspace-fixture") ?? "null") ?? [
      { name: "Weather", catalog: "weather", location: "https://weather.example.test", authentication: "anonymous", isDefault: true },
      { name: "Earthquakes", catalog: "earthquakes", location: "https://quakes.example.test", authentication: "anonymous" },
    ];
    (window as any).hostCalls = [];
    Object.defineProperty(window, "chrome", { configurable: true, value: { webview: { addEventListener() {}, postMessage(request: any) {
      (window as any).hostCalls.push(request);
      let result: unknown = true;
      if (request.method === "connections.list") result = values;
      if (request.method === "connections.workspace") {
        const members = request.params.members;
        values = values.filter((item: any) => !item.isWorkspaceProfile).map((item: any) => ({ ...item, isWorkspaceSelected: members.length === 1 && item.name === members[0] }));
        if (members.length > 1) values.push({ name: "Weather + Earthquakes", members, catalogs: members.map((name: string) => values.find((item: any) => item.name === name).catalog), catalog: values.find((item: any) => item.name === members[0]).catalog, location: "", authentication: "anonymous", isWorkspaceSelected: true, isWorkspaceProfile: true });
        localStorage.setItem("workspace-fixture", JSON.stringify(values)); result = values;
      }
      if (request.method === "agent.key.load") result = null;
      if (["query.run", "query.editor"].includes(request.method)) result = { columns: ["catalog", "schema", "name", "object_type", "kind", "summary"].map(name => ({ name, type: "VARCHAR" })), rows: [["weather", "main", "forecast", "table", "table", ""], ["earthquakes", "main", "recent", "view", "view", ""]], rowCount: 2 };
      setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
    } } } });
  });
  await page.goto("http://127.0.0.1:4173");
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  const workspace = page.getByRole("region", { name: "Workspace catalogs" });
  await expect(workspace.getByRole("checkbox", { name: /Weather/ })).toBeChecked();
  await expect(workspace.getByRole("button", { name: "Apply changes" })).toBeDisabled();
  await workspace.getByRole("checkbox", { name: /Earthquakes/ }).check();
  await expect(workspace).toContainText("Changes not applied");
  await workspace.getByLabel("Default catalog").selectOption("Earthquakes");
  await workspace.getByRole("button", { name: "Apply changes" }).click();
  await expect(workspace).toContainText("Workspace is up to date");
  await expect(page.getByRole("button", { name: "Use as active" })).toHaveCount(0);
  await page.reload();
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await expect(workspace.getByLabel("Default catalog")).toHaveValue("Earthquakes");
  await expect(workspace.getByRole("checkbox", { name: /Weather/ })).toBeChecked();
  const choice = await workspace.getByRole("checkbox", { name: /Weather/ }).boundingBox();
  expect(choice!.width).toBeLessThan(32);
  await page.screenshot({ path: "/tmp/cupola-workspace-catalogs.png" });
  await page.getByRole("tab", { name: "Catalog View", exact: true }).click();
  await expect(page.getByRole("tree", { name: "weather catalog objects" })).toBeVisible();
  await expect(page.getByRole("tree", { name: "earthquakes catalog objects" })).toBeVisible();
  const calls = await page.evaluate(() => (window as any).hostCalls);
  expect(calls.find((call: any) => call.method === "query.run").params).toMatchObject({ connection: "Weather + Earthquakes", sql: expect.stringContaining("IN ('earthquakes','weather')") });
  await page.getByRole("tab", { name: "Query Editor", exact: true }).click();
  await page.getByLabel("SQL query").fill("SELECT * FROM weather.main.forecast JOIN earthquakes.main.recent ON true");
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await expect.poll(() => page.evaluate(() => (window as any).hostCalls.find((call: any) => call.method === "query.editor")?.params.connection)).toBe("Weather + Earthquakes");
  expect(await page.evaluate(() => JSON.parse(localStorage.getItem("workspace-fixture")!).find((item: any) => item.isDefault).name)).toBe("Weather");
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
  await page.getByText("Manage profiles (advanced)", { exact: true }).click();
  await page.getByRole("button", { name: "Existing group", exact: true }).click();
  await expect(page.getByRole("group", { name: "Connections to combine" })).toBeVisible();
  await expect(page.getByRole("group", { name: "Connections to combine" }).getByRole("checkbox", { name: /Weather/ })).toBeChecked();
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
