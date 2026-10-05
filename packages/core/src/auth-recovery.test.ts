import { afterEach, expect, it, vi } from "vitest";
import { connectionSignIn, ConnectionSignInError, needsConnectionSignIn } from "./auth-recovery";

afterEach(() => { for (const name of connectionSignIn.snapshot()) connectionSignIn.clear(name); });
it.each([
  "HTTP 401: Authentication required", "VGI HTTP authentication required (HTTP 401)",
  "OAuth token refresh failed: invalid_grant", "AADSTS700082: refresh token expired due to inactivity",
  "AADSTS70043", "AADSTS50173", "interaction_required", "login_required",
  "The OAuth session expired. Sign in again.", "invalid_token", "Unauthorized",
])("recognizes interactive authentication: %s", message => {
  expect(needsConnectionSignIn(new Error(message))).toBe(true);
});
it.each([
  "HTTP 403 Forbidden", "permission denied for table customers", "HTTP 500", "HTTP 429",
  "OAuth token refresh failed: network timeout", "OAuth discovery failed (503)",
  "AADSTS700016: Application not found", "AADSTS7000215: Invalid client secret",
  "Parser error near 401", "Connection timed out", "Sign-in was closed before completion. Try again.",
])("does not mislabel other failures: %s", message => {
  expect(needsConnectionSignIn(new Error(message))).toBe(false);
});
it("deduplicates failures, retains separate identities, and only emits safe recovery text", () => {
  const listener = vi.fn(), unsubscribe = connectionSignIn.subscribe(listener);
  const error = connectionSignIn.error("Finance", new Error("invalid_grant https://private.test SELECT secret bearer_token=secret"));
  expect(error).toBeInstanceOf(ConnectionSignInError);
  expect((error as Error).message).not.toMatch(/private|SELECT|bearer_token|secret/);
  connectionSignIn.error("Finance", new Error("invalid_grant"));
  expect(listener).toHaveBeenCalledTimes(1);
  connectionSignIn.require("Sales"); connectionSignIn.clear("Finance");
  expect(connectionSignIn.snapshot()).toEqual(["Sales"]);
  const ordinary = new Error("HTTP 403");
  expect(connectionSignIn.error("Sales", ordinary)).toBe(ordinary);
  unsubscribe();
});
