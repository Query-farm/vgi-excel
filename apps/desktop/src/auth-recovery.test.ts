import { afterEach, expect, it, vi } from "vitest";
import { connectionSignIn, ConnectionSignInError } from "@query-farm/vgi-excel-core";
vi.mock("./telemetry", () => ({ captureError: vi.fn() }));
afterEach(() => {
  vi.unstubAllGlobals();
  for (const name of connectionSignIn.snapshot()) connectionSignIn.clear(name);
});
it("classifies the native query boundary without confusing an AI API failure or replaying SQL", async () => {
  const postMessage = vi.fn();
  vi.stubGlobal("window", { chrome: { webview: { postMessage, addEventListener: vi.fn() } } });
  const { invoke } = await import("./bridge");
  const failed = invoke("query.editor", { connection: "Finance", sql: "SELECT 42" });
  const rejected = expect(failed).rejects.toBeInstanceOf(ConnectionSignInError);
  window.vgiReceiveHostResponse!({ id: postMessage.mock.calls[0][0].id, error: "token refresh failed: invalid_grant" });
  await rejected;
  expect(connectionSignIn.snapshot()).toEqual(["Finance"]);
  expect(postMessage).toHaveBeenCalledTimes(1);
  connectionSignIn.clear("Finance");
  const other = invoke("agent.key.load");
  const otherRejected = expect(other).rejects.toThrow("HTTP 401");
  window.vgiReceiveHostResponse!({ id: postMessage.mock.calls[1][0].id, error: "HTTP 401" });
  await otherRejected;
  expect(connectionSignIn.snapshot()).toEqual([]);
});
