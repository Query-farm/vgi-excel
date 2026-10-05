import { expect, test, type Page } from "@playwright/test";

test.beforeEach(async ({ context }) => {
  // Synthetic authentication failures must never reach a remote telemetry ingest.
  await context.route(/https:\/\/[^/]*sentry\.io\//, route => route.abort());
});

async function desktop(page: Page, profile = false): Promise<void> {
  await page.setViewportSize({ width: 360, height: 480 });
  await page.addInitScript(({ profile }) => {
    const connection = { name: "Finance", catalog: "finance", location: "https://example.test", authentication: "oauth", isSignedIn: true, isDefault: !profile };
    const group = { name: "Workspace", catalog: "finance", location: "", authentication: "anonymous", members: ["Finance", "Sales"], isDefault: true };
    const values = profile ? [connection, { ...connection, name: "Sales", catalog: "sales" }, group] : [connection];
    const state = (window as any).__authTest = { queries: 0, signIns: 0, signedIn: false, failure: "", calls: [] as string[] };
    Object.defineProperty(window, "chrome", { configurable: true, value: { webview: {
      addEventListener() {}, postMessage(request: any) {
        state.calls.push(request.method);
        let result: unknown = true, error: string | undefined;
        if (request.method === "connections.list") result = values;
        if (request.method === "agent.key.load") result = null;
        if (request.method === "query.editor") {
          state.queries++;
          if (state.failure) error = state.failure;
          else if (state.queries > 1 && !state.signedIn) error = "OAuth token refresh failed: invalid_grant AADSTS700082";
          result = { columns: [{ name: "value", type: "INTEGER" }], rows: [[42]], rowCount: 1, truncated: false };
        }
        if (request.method === "connections.signOut") { connection.isSignedIn = false; result = values; }
        if (request.method === "connections.signIn") {
          state.signIns++;
          if (state.signIns === 1) error = "Sign-in was cancelled";
          else { state.signedIn = true; connection.isSignedIn = true; }
          result = values;
        }
        setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result, error }), 0);
      },
    } } });
  }, { profile });
  await page.goto("http://127.0.0.1:4173");
  await page.getByLabel("SQL query").fill("SELECT 42");
}

test("expired desktop session offers sign-in, survives cancellation, and never replays SQL", async ({ page }) => {
  await desktop(page);
  const run = page.getByRole("button", { name: "Run", exact: true });
  await run.click();
  await expect(page.getByRole("cell", { name: "42", exact: true })).toBeVisible();
  await run.click();
  const alert = page.getByRole("alert").filter({ hasText: "needs you to sign in again" });
  await expect(alert).toBeVisible();
  await expect(alert.getByRole("button", { name: "Retry", exact: true })).toHaveCount(0);
  expect(await page.evaluate(() => (window as any).__authTest.signIns)).toBe(0);
  const signIn = alert.getByRole("button", { name: "Sign in again", exact: true });
  await signIn.focus(); await page.keyboard.press("Enter");
  await expect(alert).toContainText("Sign-in wasn’t completed");
  await expect(signIn).toBeEnabled();
  await run.click();
  expect(await page.evaluate(() => (window as any).__authTest.queries)).toBe(2);
  expect(await page.evaluate(() => (window as any).__authTest.signIns)).toBe(1);
  await signIn.click();
  await expect(alert).toHaveCount(0);
  await expect(page.getByLabel("SQL query")).toHaveValue("SELECT 42");
  expect(await page.evaluate(() => (window as any).__authTest.queries)).toBe(2);
  await run.click();
  await expect(page.getByRole("status").filter({ hasText: "Query completed" })).toBeVisible();
  expect(await page.evaluate(() => (window as any).__authTest.queries)).toBe(3);
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
});

test("permission failures keep the ordinary error flow", async ({ page }) => {
  await desktop(page);
  await page.evaluate(() => { (window as any).__authTest.failure = "HTTP 403 Forbidden: permission denied"; });
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await expect(page.getByRole("alert")).toContainText("HTTP 403");
  await expect(page.getByRole("alert").getByRole("button", { name: "Retry", exact: true })).toBeVisible();
  await expect(page.getByRole("alert").getByRole("button", { name: "Sign in again" })).toHaveCount(0);
});

test("profile failures open Connections without authenticating a guessed member", async ({ page }) => {
  await desktop(page, true);
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await expect(page.getByRole("cell", { name: "42", exact: true })).toBeVisible();
  await page.getByRole("button", { name: "Run", exact: true }).click();
  await page.getByRole("button", { name: "Open connections", exact: true }).click();
  await expect(page.getByRole("heading", { name: "Edit Finance" })).toBeVisible();
  await expect(page.getByRole("button", { name: "Sign in again", exact: true })).toBeVisible();
  expect(await page.evaluate(() => (window as any).__authTest.signIns)).toBe(0);
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
});

test("Office offers explicit recovery at 300px and keeps it available after the dialog closes", async ({ page }) => {
  await page.setViewportSize({ width: 300, height: 480 });
  await page.route("https://appsforoffice.microsoft.com/**", route => route.abort());
  await page.addInitScript(() => {
    localStorage.setItem("vgi.excel.connections.v1", JSON.stringify([{ name: "Finance", catalog: "finance", location: "https://example.test", authentication: "oauth" }]));
    let attempts = 0;
    (window as any).Office = {
      onReady: (fn: () => void) => fn(),
      AsyncResultStatus: { Succeeded: "succeeded" },
      EventType: { DialogMessageReceived: "message", DialogEventReceived: "closed" },
      context: { ui: { displayDialogAsync: (_url: string, _options: unknown, callback: (value: unknown) => void) => {
        attempts++;
        const handlers: Record<string, (event: unknown) => void> = {};
        callback({ status: "succeeded", value: { close() {}, addEventHandler: (event: string, fn: (event: unknown) => void) => { handlers[event] = fn; } } });
        setTimeout(() => {
          if (attempts === 1) handlers.closed?.({ error: 12006 });
          else handlers.message?.({ origin: location.origin, message: JSON.stringify({ ok: true, tokens: { access_token: "test-only" } }) });
        }, 50);
      } } },
    };
  });
  await page.goto("http://127.0.0.1:4174/taskpane.html");
  await page.getByLabel("SQL query").fill("SELECT 42");
  await page.getByRole("button", { name: "Settings", exact: true }).click();
  await page.locator(".connection-form").getByRole("button", { name: "Sign in again", exact: true }).click();
  const alert = page.locator(".notice.error").filter({ hasText: "needs you to sign in again" });
  await expect(alert).toBeVisible();
  await expect(page.locator(".connection-form .field-error")).toContainText("Sign-in wasn’t completed");
  await alert.getByRole("button", { name: "Sign in again", exact: true }).click();
  await expect(alert).toHaveCount(0);
  await expect(page.locator(".notice.info")).toContainText("Signed in");
  await expect(page.locator(".connection-form .field-error")).toHaveCount(0);
  await page.getByRole("tab", { name: "Query Editor", exact: true }).click();
  await expect(page.getByLabel("SQL query")).toHaveValue("SELECT 42");
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
});
