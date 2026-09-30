import { defineConfig, devices } from "@playwright/test";

const hosted = !!process.env.CUPOLA_OFFICE_BASE_URL;

export default defineConfig({
  testDir: "tests/office-wasm",
  timeout: 180_000,
  expect: { timeout: 120_000 },
  workers: 1,
  retries: 0,
  reporter: "line",
  use: { ...devices["Desktop Chrome"], headless: true, ignoreHTTPSErrors: !hosted },
  webServer: hosted ? undefined : {
    command: "npm exec vite -- preview --host 127.0.0.1 --port 4184",
    cwd: "apps/office",
    port: 4184,
    reuseExistingServer: false,
    timeout: 60_000,
  },
});
