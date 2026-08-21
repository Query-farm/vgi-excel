import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "tests/ui",
  timeout: 30_000,
  use: { headless: true, ignoreHTTPSErrors: true },
  webServer: [
    { command: "npm exec vite -- preview --host 127.0.0.1 --port 4173", cwd: "apps/desktop", port: 4173, reuseExistingServer: true },
    { command: "python3 -m http.server 4174 --bind 127.0.0.1", cwd: "apps/office/dist", port: 4174, reuseExistingServer: true },
  ],
});
