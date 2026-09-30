import { defineConfig, type PluginOption, type UserConfig, type ViteDevServer } from "vite";
import react from "@vitejs/plugin-react";
import { sentryVitePlugin } from "@sentry/vite-plugin";
import devCerts from "office-addin-dev-certs";
import { cpSync, mkdirSync, readFileSync, createReadStream, statSync } from "node:fs";
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
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const path = req.url?.split("?")[0] ?? "";
        if (!path.startsWith("/haybarn/")) return next();
        const name = path.slice("/haybarn/".length);
        if (!haybarnFiles.includes(name) || !["GET", "HEAD"].includes(req.method ?? "")) {
          res.statusCode = 404;
          res.end();
          return;
        }
        const file = resolve(haybarnSource, name);
        try {
          res.setHeader("Content-Type", name.endsWith(".wasm") ? "application/wasm" : "text/javascript");
          res.setHeader("Content-Length", statSync(file).size);
          res.setHeader("Cache-Control", "no-cache");
          if (req.method === "HEAD") { res.end(); return; }
          const stream = createReadStream(file);
          stream.on("error", error => res.destroy(error));
          res.on("close", () => stream.destroy());
          stream.pipe(res);
        } catch (error) { next(error as Error); }
      });
    },
    writeBundle(options) {
      const target = resolve(options.dir ?? resolve(here, "dist"), "haybarn");
      mkdirSync(target, { recursive: true });
      for (const file of haybarnFiles) cpSync(resolve(haybarnSource, file), resolve(target, file));
    },
  };
}

// Excel fetches this public metadata from its own origin before starting the
// shared runtime. Keep cross-origin access scoped to metadata, not dev sources.
function customFunctionsMetadataCors(): PluginOption {
  const configure = (server: Pick<ViteDevServer, "middlewares">) => {
    server.middlewares.use((req, res, next) => {
      if (req.url?.split("?")[0] !== "/functions.json") return next();
      if (!["GET", "HEAD", "OPTIONS"].includes(req.method ?? "")) return next();
      res.setHeader("Access-Control-Allow-Origin", "*");
      res.setHeader("Access-Control-Allow-Methods", "GET, HEAD, OPTIONS");
      if (req.method === "OPTIONS") { res.statusCode = 204; res.end(); return; }
      next();
    });
  };
  return { name: "custom-functions-metadata-cors", configureServer: configure, configurePreviewServer: configure };
}

export default defineConfig(async ({ command, isPreview }) => {
  let https;
  if (command === "serve" && !process.env.VITEST) {
    if (isPreview) {
      // Playwright accepts this local certificate; previews need no OS trust changes.
      const certDir = resolve(here, "../../dev-certs/preview");
      const ca = resolve(certDir, "ca.crt");
      const cert = resolve(certDir, "localhost.crt");
      const key = resolve(certDir, "localhost.key");
      await devCerts.generateCertificates(ca, cert, key);
      https = { ca: readFileSync(ca), cert: readFileSync(cert), key: readFileSync(key) };
    } else {
      https = await devCerts.getHttpsServerOptions();
    }
  }
  const uploadSourceMaps = command === "build" && !!process.env.SENTRY_AUTH_TOKEN && !!process.env.SENTRY_ORG;
  const config: UserConfig = {
    plugins: [
      react(),
      copyHaybarnArtifacts(),
      customFunctionsMetadataCors(),
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
    server: { https, cors: false },
    preview: { cors: false, headers: { "Cross-Origin-Opener-Policy": "same-origin", "Cross-Origin-Embedder-Policy": "require-corp" } },
    build: {
      target: "es2022",
      sourcemap: uploadSourceMaps ? "hidden" : false,
      rollupOptions: {
        input: {
          taskpane: resolve(here, "taskpane.html"),
          results: resolve(here, "results.html"),
          oauthDialog: resolve(here, "oauth-dialog.html"),
        },
      },
    },
  };
  return config;
});
