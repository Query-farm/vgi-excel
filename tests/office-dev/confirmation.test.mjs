import { test } from "node:test";
import assert from "node:assert/strict";
import { createServer } from "vite";
import { webkit, expect } from "@playwright/test";
import { mkdtemp, rm } from "node:fs/promises";
import { resolve } from "node:path";

process.env.VITEST = "1";
test("Office AI insertion confirms in-app when browser dialogs are unavailable", async () => {
  // A warm optimizer cache can hide broken entry paths when starting from the repository root.
  const cacheDir = await mkdtemp(resolve("node_modules/.cupola-office-vite-"));
  const server = await createServer({ root: "apps/office", configFile: "apps/office/vite.config.ts", cacheDir, server: { host: "127.0.0.1", port: 0, open: false } });
  const browser = await webkit.launch();
  try {
    await server.listen();
    const base = `http://127.0.0.1:${server.httpServer.address().port}`;
    const page = await browser.newPage({ viewport: { width: 300, height: 600 } });
    const errors = [];
    page.on("pageerror", error => errors.push(error.message));
    await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
    await page.goto(`${base}/taskpane.html`);
    await page.getByRole("button", { name: "Add connection" }).waitFor();
    await page.evaluate(async cache => {
      window.confirm = () => { throw new Error("Function window.confirm is not supported."); };
      const { default: React } = await import(`/@fs${cache}/deps/react.js`);
      const { default: { createRoot } } = await import(`/@fs${cache}/deps/react-dom_client.js`);
      const { AgentResult } = await import("/src/AgentExtras.tsx");
      document.getElementById("root").hidden = true;
      const fixture = document.createElement("div"); document.body.append(fixture);
      window.testWrites = 0;
      createRoot(fixture).render(React.createElement(AgentResult, {
        value: { id: "test", connection: "Weather", sql: "SELECT 42", scope: {}, result: { columns: [{ name: "answer", type: "INTEGER" }], rows: [[42]], rowCount: 1 } },
        disabled: false, openQuery() {}, loadLabel: "Insert into Excel",
        async load() { window.testWrites++; return "Snapshot inserted."; },
      }));
    }, cacheDir);
    const insert = page.getByRole("button", { name: "Insert into Excel", exact: true });
    await insert.focus();
    await page.keyboard.press("Enter");
    const dialog = page.getByRole("dialog", { name: "Insert into Excel" });
    await expect(dialog).toBeVisible();
    await expect(dialog).toContainText("1 row at the current selection");
    await expect(dialog.getByRole("button", { name: "Cancel" })).toBeFocused();
    assert.equal(await page.evaluate(() => window.testWrites), 0);
    await page.keyboard.press("Escape");
    await expect(dialog).toHaveCount(0);
    await expect(insert).toBeFocused();
    assert.equal(await page.evaluate(() => window.testWrites), 0);
    await insert.click();
    await dialog.getByRole("button", { name: "Cancel" }).click();
    assert.equal(await page.evaluate(() => window.testWrites), 0);
    await insert.click();
    await page.screenshot({ path: "/tmp/cupola-office-confirmation.png" });
    await dialog.getByRole("button", { name: "Insert table" }).click();
    await expect(page.getByRole("status")).toHaveText("Snapshot inserted.");
    assert.equal(await page.evaluate(() => window.testWrites), 1);
    assert.deepEqual(errors, []);
  } finally { await browser.close(); await server.close(); await rm(cacheDir, { recursive: true, force: true }); }
});
