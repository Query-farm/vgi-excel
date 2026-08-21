import { defineConfig, type PluginOption, type UserConfig } from "vite";
import react from "@vitejs/plugin-react";
import { sentryVitePlugin } from "@sentry/vite-plugin";
import devCerts from "office-addin-dev-certs";
import { cpSync, mkdirSync, readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const product = JSON.parse(readFileSync(resolve(here, "../../package.json"), "utf8")) as { version: string; cupolaBuild: string };
const release = `cupola-excel@${product.version}+${product.cupolaBuild}`;
const haybarnSource = resolve(here, "../../node_modules/@haybarn/haybarn-wasm/dist");
const haybarnFiles = [
  "duckdb-mvp.wasm", "duckdb-eh.wasm", "duckdb-coi.wasm",
  "duckdb-browser-mvp.worker.js", "duckdb-browser-eh.worker.js", "duckdb-browser-coi.worker.js",
  "duckdb-browser-coi.pthread.worker.js",
];

function copyHaybarnArtifacts(): PluginOption {
  return {
    name: "copy-haybarn-artifacts",
    writeBundle(options) {
      const target = resolve(options.dir ?? resolve(here, "dist"), "haybarn");
      mkdirSync(target, { recursive: true });
      for (const file of haybarnFiles) cpSync(resolve(haybarnSource, file), resolve(target, file));
    },
  };
}

export default defineConfig(async ({ command }) => {
  const https = command === "serve" ? await devCerts.getHttpsServerOptions() : undefined;
  const uploadSourceMaps = command === "build" && !!process.env.SENTRY_AUTH_TOKEN && !!process.env.SENTRY_ORG;
  const config: UserConfig = {
    plugins: [
      react(),
      copyHaybarnArtifacts(),
      ...(uploadSourceMaps ? sentryVitePlugin({
        authToken: process.env.SENTRY_AUTH_TOKEN,
        org: process.env.SENTRY_ORG,
        project: process.env.SENTRY_OFFICE_PROJECT ?? "cupola-excel-office",
        release: { name: release, dist: "office", setCommits: false },
        sourcemaps: { assets: "./dist/**", filesToDeleteAfterUpload: "./dist/**/*.map" },
        telemetry: false,
      }) as unknown as PluginOption[] : []),
    ] as PluginOption[],
    define: { __APP_VERSION__: JSON.stringify(product.version), __BUILD_ID__: JSON.stringify(product.cupolaBuild) },
    server: { https },
    preview: { headers: { "Cross-Origin-Opener-Policy": "same-origin", "Cross-Origin-Embedder-Policy": "require-corp" } },
    build: {
      target: "es2022",
      sourcemap: uploadSourceMaps ? "hidden" : false,
      rollupOptions: {
        input: {
          taskpane: "taskpane.html",
          oauthDialog: "oauth-dialog.html",
        },
      },
    },
  };
  return config;
});
