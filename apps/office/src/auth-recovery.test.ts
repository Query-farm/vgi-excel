import { afterEach, expect, it, vi } from "vitest";
import { tableFromArrays } from "apache-arrow";
import { connectionSignIn, ConnectionSignInError } from "@query-farm/vgi-excel-core";
import { BrowserBackend } from "./browser-backend";
import { signIn } from "./oauth";
vi.mock("./oauth", () => ({ signIn: vi.fn() }));
afterEach(() => {
  vi.unstubAllGlobals(); vi.clearAllMocks();
  for (const name of connectionSignIn.snapshot()) connectionSignIn.clear(name);
});
it("detects expiry after a working query, avoids popup loops and permits explicit recovery", async () => {
  vi.stubGlobal("sessionStorage", { getItem: () => null });
  let expired = false;
  const query = vi.fn(async (sql: string) => {
    if (sql === "SELECT 42" && expired) throw new Error("OAuth token refresh failed: invalid_grant AADSTS700082");
    return tableFromArrays({ value: [42] });
  });
  const factory = async () => ({ db: { connect: async () => ({ query }) } }) as any;
  const definition = { name: "Finance", catalog: "finance", location: "https://example.test" };
  const backend = new BrowserBackend(definition, factory);
  expect((await backend.query("SELECT 42")).rows).toEqual([[42]]);
  expired = true;
  await expect(backend.query("SELECT 42")).rejects.toBeInstanceOf(ConnectionSignInError);
  expect(connectionSignIn.snapshot()).toEqual(["Finance"]);
  const calls = query.mock.calls.length;
  await expect(backend.query("SELECT 42")).rejects.toBeInstanceOf(ConnectionSignInError);
  expect(query).toHaveBeenCalledTimes(calls);
  expect(signIn).not.toHaveBeenCalled();
  // The recovery action clears the marker and resets the runtime after sign-in.
  expired = false; connectionSignIn.clear("Finance");
  const fresh = new BrowserBackend(definition, factory);
  expect((await fresh.query("SELECT 42")).rows).toEqual([[42]]);
});
it("leaves permission failures as ordinary errors", async () => {
  vi.stubGlobal("sessionStorage", { getItem: () => null });
  const query = async (sql: string) => {
    if (sql === "SELECT 42") throw new Error("HTTP 403 Forbidden");
    return tableFromArrays({ value: [42] });
  };
  const backend = new BrowserBackend({ name: "Finance", location: "https://example.test" }, async () => ({ db: { connect: async () => ({ query }) } }) as any);
  await expect(backend.query("SELECT 42")).rejects.toThrow("HTTP 403");
  expect(connectionSignIn.snapshot()).toEqual([]);
  expect(signIn).not.toHaveBeenCalled();
});
