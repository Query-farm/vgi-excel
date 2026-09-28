import { expect, test } from "@playwright/test";

for (const host of ["desktop", "office"]) {
  test(`${host} compact composer keeps Send beside the focused input without overlap`, async ({ page }) => {
    await page.setViewportSize({ width: host === "office" ? 300 : 360, height: 480 });
    await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
    await page.addInitScript(() => {
      Object.defineProperty(window, "chrome", { configurable: true, value: { webview: { addEventListener() {}, postMessage(request: { id: number; method: string }) {
        setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result: request.method === "connections.list" ? [] : null }), 0);
      } } } });
    });
    await page.goto(host === "office" ? "http://127.0.0.1:4174/taskpane.html" : "http://127.0.0.1:4173/index.html");
    await page.getByRole("tab", { name: "Ask AI", exact: true }).click();
    const prompt = page.getByRole("textbox", { name: "Ask AI", exact: true });
    await prompt.fill("A draft question");
    await prompt.press("Shift+Enter");
    await prompt.press("Escape");
    await expect(prompt).toHaveValue("A draft question\n");
    const bounds = await prompt.evaluate(input => {
      const rect = input.getBoundingClientRect(), button = input.parentElement!.querySelector("button")!.getBoundingClientRect();
      const style = getComputedStyle(input), ring = parseFloat(style.outlineWidth) + parseFloat(style.outlineOffset);
      return { height: rect.height, gap: button.left - rect.right - ring, aligned: Math.abs(button.bottom - rect.bottom), right: button.right, width: innerWidth, inputScrolls: getComputedStyle(input).overflowY };
    });
    expect(bounds.height).toBeLessThanOrEqual(60);
    expect(bounds.gap).toBeGreaterThanOrEqual(3);
    expect(bounds.aligned).toBeLessThanOrEqual(1);
    expect(bounds.right).toBeLessThanOrEqual(bounds.width - 5);
    expect(bounds.inputScrolls).toBe("auto");
    await page.screenshot({ path: `/tmp/cupola-composer-${host}.png` });
  });
}
