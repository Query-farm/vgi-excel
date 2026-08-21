import { defineConfig, type PluginOption } from "vite";
import react from "@vitejs/plugin-react";
import { sentryVitePlugin } from "@sentry/vite-plugin";
import { readFileSync } from "node:fs";

const product = JSON.parse(readFileSync(new URL("../../package.json", import.meta.url), "utf8")) as { version: string; cupolaBuild: string };
const release = `cupola-excel@${product.version}+${product.cupolaBuild}`;
const uploadSourceMaps = !!process.env.SENTRY_AUTH_TOKEN && !!process.env.SENTRY_ORG;

export default defineConfig({
  base: "./",
  plugins: [
    react(),
    ...(uploadSourceMaps ? sentryVitePlugin({
      authToken: process.env.SENTRY_AUTH_TOKEN,
      org: process.env.SENTRY_ORG,
      project: process.env.SENTRY_DESKTOP_PROJECT ?? "cupola-excel-desktop",
      release: { name: release, dist: "desktop", setCommits: false },
      sourcemaps: { assets: "./dist/**", filesToDeleteAfterUpload: "./dist/**/*.map" },
      telemetry: false,
    }) as unknown as PluginOption[] : []),
  ] as PluginOption[],
  define: { __APP_VERSION__: JSON.stringify(product.version), __BUILD_ID__: JSON.stringify(product.cupolaBuild) },
  build: {
    outDir: "dist", emptyOutDir: true, sourcemap: uploadSourceMaps ? "hidden" : false,
    rollupOptions: { output: { entryFileNames: "assets/workbench.js", chunkFileNames: "assets/[name].js", assetFileNames: "assets/workbench[extname]" } },
  },
});
