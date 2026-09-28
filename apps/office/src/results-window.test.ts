import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { openResultsWindow } from "./results-window";

beforeEach(() => vi.stubGlobal("document", { documentElement: { dataset: { officeTheme: "light" }, style: { getPropertyValue: (key: string) => key === "--office-background" ? "#ffffff" : "#242424" } } }));
afterEach(() => { vi.unstubAllGlobals(); vi.useRealTimers(); });
describe("Office results dialog transport", () => {
  it("transfers large results in ordered chunks only to the same origin", async () => {
    vi.stubGlobal("window", globalThis);
    vi.stubGlobal("location", new URL("https://cupola.example/taskpane.html"));
    const handlers: Record<string, (event: { message: string; origin?: string }) => void> = {};
    const packets: Array<{ text: string; index: number; count: number }> = [];
    const close = vi.fn();
    const messageChild = vi.fn((text: string, options: { targetOrigin: string }) => {
      expect(options.targetOrigin).toBe("https://cupola.example");
      const packet = JSON.parse(text); packets.push(packet);
      queueMicrotask(() => handlers.message({ origin: "https://cupola.example", message: JSON.stringify(packet.index + 1 === packet.count ? { type: "received" } : { type: "chunk", index: packet.index + 1 }) }));
    });
    vi.stubGlobal("Office", { AsyncResultStatus: { Succeeded: "ok" }, EventType: { DialogMessageReceived: "message", DialogEventReceived: "close" }, context: { host: "Excel", requirements: { isSetSupported: () => true }, ui: { displayDialogAsync(url: string, options: unknown, done: (result: unknown) => void) {
      expect(url).toBe("https://cupola.example/results.html?mode=office");
      done({ status: "ok", value: { close, messageChild, addEventHandler(type: string, handler: typeof handlers[string]) { handlers[type] = handler; } } });
    } } } });
    const snapshot = { title: "Report", result: { columns: [{ name: "text", type: "VARCHAR" }], rows: [["sensitive cell ".repeat(4000)]], rowCount: 1, truncated: false } };
    const pending = openResultsWindow(snapshot);
    handlers.message({ origin: "https://untrusted.example", message: JSON.stringify({ type: "chunk", index: 0 }) });
    expect(messageChild).not.toHaveBeenCalled();
    handlers.message({ origin: "https://cupola.example", message: JSON.stringify({ type: "chunk", index: 0 }) });
    await pending;
    expect(packets.length).toBeGreaterThan(1);
    expect(packets.every(packet => packet.text.length <= 16_000)).toBe(true);
    expect(JSON.parse(packets.map(packet => packet.text).join(""))).toMatchObject(snapshot);
    expect(JSON.parse(packets.map(packet => packet.text).join(""))).toHaveProperty("theme.bodyBackgroundColor", "#ffffff");
    expect(close).not.toHaveBeenCalled();
  });
});

it("reports a blocked browser popup without putting result values in the URL", async () => {
  vi.stubGlobal("Office", undefined);
  vi.stubGlobal("window", globalThis);
  vi.stubGlobal("location", new URL("https://cupola.example/taskpane.html"));
  const close = vi.fn();
  vi.stubGlobal("BroadcastChannel", class { close = close; });
  const open = vi.fn((_url: URL | string) => null);
  vi.stubGlobal("open", open);
  await expect(openResultsWindow({ title: "Private title", result: { columns: [], rows: [], rowCount: 0, truncated: false } })).rejects.toThrow("Allow popups");
  expect(String(open.mock.calls[0]?.[0])).not.toContain("Private");
  expect(close).toHaveBeenCalledOnce();
});
