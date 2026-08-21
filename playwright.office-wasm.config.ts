import { defineConfig, devices } from "@playwright/test";

export default defineConfig({
  testDir: "tests/office-wasm",
  timeout: 180_000,
  expect: { timeout: 120_000 },
  workers: 1,
  retries: 0,
  reporter: "line",
  use: { ...devices["Desktop Chrome"], headless: true, ignoreHTTPSErrors: true },
  webServer: {
    command: "npm exec vite -- preview --host 127.0.0.1 --port 4184",
    cwd: "apps/office",
    port: 4184,
    reuseExistingServer: false,
    timeout: 60_000,
  },
});
