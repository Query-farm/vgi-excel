import { afterEach, expect, it, vi } from "vitest";
import type { ConfigEnv, UserConfig } from "vite";
import devCerts from "office-addin-dev-certs";
import config from "../vite.config";

vi.mock("office-addin-dev-certs", () => ({
  default: {
    getHttpsServerOptions: vi.fn(() => { throw new Error("OS certificate installation must not run"); }),
    generateCertificates: vi.fn(() => { throw new Error("Certificate generation must not run"); }),
  },
}));

afterEach(() => { vi.unstubAllEnvs(); vi.clearAllMocks(); });

it.each([
  { command: "serve", mode: "test", isPreview: false, vitest: "true" },
  { command: "build", mode: "production", isPreview: false, vitest: "" },
] as const)("does not install certificates for $mode", async ({ vitest, ...env }) => {
  vi.stubEnv("VITEST", vitest);
  vi.stubEnv("SENTRY_AUTH_TOKEN", "");
  const result = await (config as (env: ConfigEnv) => Promise<UserConfig>)(env);
  expect(result.server?.https).toBeUndefined();
  expect(devCerts.getHttpsServerOptions).not.toHaveBeenCalled();
  expect(devCerts.generateCertificates).not.toHaveBeenCalled();
});
