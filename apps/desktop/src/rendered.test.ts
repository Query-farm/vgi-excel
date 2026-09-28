import { afterEach, expect, it, vi } from "vitest";
import { invoke } from "./bridge";
import { notifyRendered } from "./rendered";
vi.mock("./bridge", () => ({ invoke: vi.fn(async () => true) }));
afterEach(() => { vi.unstubAllGlobals(); vi.clearAllMocks(); });

it.each([false, true])("reveals the native UI only after paint, unless unmounted (%s)", unmount => {
  let next = 0;
  const frames = new Map<number, FrameRequestCallback>();
  vi.stubGlobal("window", { chrome: { webview: {} } });
  vi.stubGlobal("requestAnimationFrame", (callback: FrameRequestCallback) => { frames.set(++next, callback); return next; });
  vi.stubGlobal("cancelAnimationFrame", (id: number) => frames.delete(id));
  const cleanup = notifyRendered();
  expect(invoke).not.toHaveBeenCalled();
  frames.get(1)!(0); frames.delete(1);
  expect(invoke).not.toHaveBeenCalled();
  if (unmount) cleanup();
  frames.get(2)?.(16);
  expect(invoke).toHaveBeenCalledTimes(unmount ? 0 : 1);
  if (!unmount) expect(invoke).toHaveBeenCalledWith("ui.rendered", {}, 15_000);
});
