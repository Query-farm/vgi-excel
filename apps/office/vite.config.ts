import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import devCerts from "office-addin-dev-certs";
import { cpSync, mkdirSync, readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const product = JSON.parse(readFileSync(resolve(here, "../../package.json"), "utf8")) as { version: string; cupolaBuild: string };
const haybarnSource = resolve(here, "../../node_modules/@haybarn/haybarn-wasm/dist");
const haybarnFiles = [
  "duckdb-mvp.wasm", "duckdb-eh.wasm", "duckdb-coi.wasm",
  "duckdb-browser-mvp.worker.js", "duckdb-browser-eh.worker.js", "duckdb-browser-coi.worker.js",
  "duckdb-browser-coi.pthread.worker.js",
];

function copyHaybarnArtifacts() {
  return {
    name: "copy-haybarn-artifacts",
    writeBundle(options: { dir?: string }) {
      const target = resolve(options.dir ?? resolve(here, "dist"), "haybarn");
      mkdirSync(target, { recursive: true });
      for (const file of haybarnFiles) cpSync(resolve(haybarnSource, file), resolve(target, file));
    },
  };
}

export default defineConfig(async ({ command }) => {
  const https = command === "serve" ? await devCerts.getHttpsServerOptions() : undefined;
  return {
    plugins: [react(), copyHaybarnArtifacts()],
    define: { __APP_VERSION__: JSON.stringify(product.version), __BUILD_ID__: JSON.stringify(product.cupolaBuild) },
    server: { https },
    preview: { headers: { "Cross-Origin-Opener-Policy": "same-origin", "Cross-Origin-Embedder-Policy": "require-corp" } },
    build: {
      target: "es2022",
      rollupOptions: {
        input: {
          taskpane: "taskpane.html",
          oauthDialog: "oauth-dialog.html",
        },
      },
    },
  };
});
