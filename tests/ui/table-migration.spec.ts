import { expect, test } from "@playwright/test";

for (const fail of [false, true]) {
  test(`legacy table migration preserves the original when creation ${fail ? "fails" : "starts"}`, async ({ page }) => {
    await page.setViewportSize({ width: 360, height: 480 });
    await page.addInitScript(({ fail }) => {
      let managed = true;
      const calls: string[] = [];
      (window as any).__tableCalls = calls;
      const webview = { addEventListener() {}, postMessage(request: any) {
        calls.push(request.method);
        let result: unknown = true;
        if (request.method === "connections.list") result = [];
        if (request.method === "agent.key.load") result = null;
        if (request.method === "excel.snapshots") result = managed ? [{ table: "LegacySales", connection: "Sales", sql: "SELECT 42", updatedAt: "2026-01-01T00:00:00Z" }] : [];
        if (request.method === "excel.forgetSnapshot") managed = false;
        if (request.method === "excel.migrateSnapshot") result = { query: "LegacySales Refreshable", loaded: true, message: "Power Query created and refresh started." };
        const error = request.method === "excel.migrateSnapshot" && fail ? "Connection unavailable" : undefined;
        setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result, error }), 0);
      } };
      Object.defineProperty(window, "chrome", { value: { webview }, configurable: true });
    }, { fail });
    await page.goto("http://127.0.0.1:4173");
    await page.getByRole("button", { name: "Workbook data", exact: true }).click();
    page.once("dialog", dialog => dialog.dismiss());
    await page.getByRole("button", { name: "Create refreshable copy" }).click();
    expect(await page.evaluate(() => (window as any).__tableCalls.includes("excel.migrateSnapshot"))).toBe(false);
    page.once("dialog", dialog => dialog.accept());
    await page.getByRole("button", { name: "Create refreshable copy" }).click();
    await expect(page.getByText(fail ? "Connection unavailable" : /Power Query created and refresh started/)).toBeVisible();
    await expect(page.getByText("LegacySales", { exact: true })).toBeVisible();
    expect(await page.evaluate(() => (window as any).__tableCalls.includes("excel.forgetSnapshot"))).toBe(false);
    page.once("dialog", dialog => dialog.accept());
    await page.getByRole("button", { name: "Keep as static table" }).click();
    await expect(page.getByText("No Cupola tables", { exact: true })).toBeVisible();
  });
}
