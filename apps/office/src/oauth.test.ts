import { afterEach, expect, it, vi } from "vitest";
import { signIn } from "./oauth";

afterEach(() => vi.unstubAllGlobals());
function fixture() {
  const handlers: Record<string, (event: unknown) => void> = {};
  const store = new Map<string, string>();
  const dialog = { close: vi.fn(() => handlers.closed?.({ error: 12006 })), addEventHandler: vi.fn((event: string, handler: (event: unknown) => void) => { handlers[event] = handler; }) };
  let opened!: (result: unknown) => void;
  vi.stubGlobal("window", { location: { origin: "https://cupola.example" } });
  vi.stubGlobal("sessionStorage", { setItem: (key: string, value: string) => store.set(key, value) });
  vi.stubGlobal("Office", { AsyncResultStatus: { Succeeded: "succeeded" }, EventType: { DialogMessageReceived: "message", DialogEventReceived: "closed" }, context: { ui: { displayDialogAsync: (_url: string, _options: unknown, callback: typeof opened) => { opened = callback; } } } });
  return { handlers, store, dialog, open: () => opened({ status: "succeeded", value: dialog }), fail: () => opened({ status: "failed" }) };
}
it("stores a valid session once despite the host close event, ignoring foreign messages", async () => {
  const f = fixture(); const result = signIn("https://catalog.example/vgi"); f.open();
  f.handlers.message!({ origin: "https://other.example", message: JSON.stringify({ ok: true, tokens: { access_token: "wrong" } }) });
  expect(f.store.size).toBe(0);
  f.handlers.message!({ origin: "https://cupola.example", message: JSON.stringify({ ok: true, tokens: { access_token: "test-access", refresh_token: "test-refresh", id_token: "discard" } }) });
  expect(await result).toEqual({ access_token: "test-access", refresh_token: "test-refresh" });
  expect(f.dialog.close).toHaveBeenCalledOnce();
  expect([...f.store.values()][0]).not.toContain("discard");
});
it.each([true, false])("cancels sign-in before or after opening without storing tokens (%s)", async opened => {
  const f = fixture(), abort = new AbortController();
  const result = signIn("https://catalog.example", abort.signal);
  const rejected = expect(result).rejects.toMatchObject({ name: "AbortError" });
  if (opened) f.open();
  abort.abort();
  if (!opened) f.open();
  await rejected;
  expect(f.dialog.close).toHaveBeenCalledOnce(); expect(f.store.size).toBe(0);
  f.handlers.message?.({ message: JSON.stringify({ ok: true, tokens: { access_token: "late" } }) });
  expect(f.store.size).toBe(0);
});
it.each(["blocked", "closed", "invalid"])("handles %s sign-in without persisting a session", async kind => {
  const f = fixture(); const result = signIn("https://catalog.example");
  const rejected = expect(result).rejects.toThrow();
  if (kind === "blocked") f.fail();
  else { f.open(); if (kind === "closed") f.handlers.closed!({ error: 12006 }); else f.handlers.message!({ message: '{"ok":true,"tokens":{}}' }); }
  await rejected; expect(f.store.size).toBe(0);
});
