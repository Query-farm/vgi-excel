import { describe, expect, it } from "vitest";
import { isRecoverableAuthError } from "./auth-errors";

describe("automatic VGI authentication", () => {
  it("recognizes challenges that should open sign-in", () => {
    expect(isRecoverableAuthError(new Error("HTTP 401: Authentication required"))).toBe(true);
    expect(isRecoverableAuthError(new Error("OAuth token expired"))).toBe(true);
  });

  it("does not redirect for unrelated or unrecoverable failures", () => {
    expect(isRecoverableAuthError(new Error("Could not reach author.example.com"))).toBe(false);
    expect(isRecoverableAuthError(new Error("token exchange failed: invalid_grant"))).toBe(false);
  });
});
