import { afterEach, expect, it, vi } from "vitest";
const service = vi.hoisted(() => ({ query: vi.fn(), value: vi.fn(), call: vi.fn() }));
vi.mock("./runtime", () => ({ formulaService: service }));
vi.mock("./telemetry", () => ({ captureError: vi.fn() }));
afterEach(() => { vi.unstubAllGlobals(); vi.clearAllMocks(); });
async function load() {
  vi.resetModules();
  const registered = new Map<string, (...args: any[]) => Promise<unknown>>();
  class FunctionError extends Error { constructor(public code: string, message: string) { super(message); } }
  vi.stubGlobal("CustomFunctions", { associate: (name: string, fn: (...args: any[]) => Promise<unknown>) => registered.set(name, fn), Error: FunctionError, ErrorCode: { notAvailable: "NA", invalidValue: "VALUE" } });
  await import("./functions");
  return registered;
}
it("registers worksheet functions and returns spill-safe values", async () => {
  const functions = await load();
  expect([...functions.keys()]).toEqual(["QUERY", "VALUE", "CALL"]);
  service.query.mockResolvedValueOnce([["answer"], [42], [null]]);
  expect(await functions.get("QUERY")!("SELECT 42", "Weather", true)).toEqual([["answer"], [42], [""]]);
  expect(service.query).toHaveBeenCalledWith("SELECT 42", "Weather", true, expect.any(AbortSignal));
  service.value.mockResolvedValueOnce(null);
  expect(await functions.get("VALUE")!("SELECT NULL", " ")).toBe("");
  service.call.mockResolvedValueOnce([["Clear sky"]]);
  expect(await functions.get("CALL")!("weather_code_text", 0)).toEqual([["Clear sky"]]);
});
it("passes Excel invocation cancellation to the running query and permits another call", async () => {
  const functions = await load();
  const invocation: { onCanceled?: () => void } = {};
  service.query.mockImplementationOnce((_sql, _connection, _headers, signal: AbortSignal) => new Promise((_resolve, reject) => signal.addEventListener("abort", () => reject(new DOMException("Cancelled", "AbortError")))));
  const pending = functions.get("QUERY")!("SELECT long_query", "Weather", false, undefined, invocation);
  const rejected = expect(pending).rejects.toThrow("Cancelled");
  invocation.onCanceled!(); await rejected;
  service.query.mockResolvedValueOnce([[42]]);
  expect(await functions.get("QUERY")!("SELECT 42", "Weather", false)).toEqual([[42]]);
});
it("returns an Excel error for a disconnected source", async () => {
  const functions = await load(); service.value.mockRejectedValueOnce(new Error("Connection unavailable"));
  await expect(functions.get("VALUE")!("SELECT 42")).rejects.toMatchObject({ code: "NA" });
});
