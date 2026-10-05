import { afterEach, expect, it, vi } from "vitest";
import { connectionSignIn } from "@query-farm/vgi-excel-core";
import { reauthenticateConnection } from "./reauthenticate";
import { host, type DesktopConnection } from "./bridge";
vi.mock("./bridge", () => ({ host: { signIn: vi.fn(), signOut: vi.fn() } }));
afterEach(() => { vi.clearAllMocks(); connectionSignIn.clear("Finance"); });
it("invalidates native sessions before signing in and keeps recovery available on cancellation", async () => {
  const connection: DesktopConnection = { name: "Finance", catalog: "finance", location: "https://example.test", authentication: "oauth" };
  const update = vi.fn();
  vi.mocked(host.signOut).mockResolvedValue([{ ...connection, isSignedIn: false }]);
  vi.mocked(host.signIn).mockRejectedValueOnce(new Error("Cancelled"));
  await expect(reauthenticateConnection(connection, update)).rejects.toThrow("Cancelled");
  expect(connectionSignIn.snapshot()).toContain("Finance");
  expect(update).toHaveBeenCalledWith([{ ...connection, isSignedIn: false }]);
  vi.mocked(host.signIn).mockResolvedValueOnce([{ ...connection, isSignedIn: true }]);
  await reauthenticateConnection(connection, update);
  expect(connectionSignIn.snapshot()).not.toContain("Finance");
  expect(vi.mocked(host.signOut).mock.invocationCallOrder[0]).toBeLessThan(vi.mocked(host.signIn).mock.invocationCallOrder[0]);
  expect(update).toHaveBeenLastCalledWith([{ ...connection, isSignedIn: true }]);
});
it("does not guess which profile member needs authentication", async () => {
  await expect(reauthenticateConnection({ name: "Group", catalog: "", location: "", authentication: "anonymous", members: ["Finance", "Sales"] }, vi.fn())).rejects.toThrow("individual");
  expect(host.signIn).not.toHaveBeenCalled();
  expect(host.signOut).not.toHaveBeenCalled();
});
