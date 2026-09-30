import { test } from "node:test";
import assert from "node:assert/strict";
import { createServer } from "vite";
import { mkdtemp, rm } from "node:fs/promises";
import { resolve } from "node:path";

// Exercise the real dev middleware without installing development certificates.
process.env.VITEST = "1";
test("Office development serves engine assets and cross-origin metadata", async () => {
  const cacheDir = await mkdtemp(resolve("node_modules/.cupola-office-http-"));
  const server = await createServer({
    root: "apps/office", configFile: "apps/office/vite.config.ts", cacheDir,
    // The browser fixtures cover dependency optimization and module execution.
    // Keep this short-lived HTTP fixture free of background scans and warmups.
    optimizeDeps: { noDiscovery: true, include: [] },
    server: { host: "127.0.0.1", port: 0, open: false, preTransformRequests: false, watch: { ignored: () => true } },
  });
  try {
    await server.listen();
    const base = `http://127.0.0.1:${server.httpServer.address().port}`;
    for (const bundle of ["mvp", "eh", "coi"]) {
      const worker = await fetch(`${base}/haybarn/duckdb-browser-${bundle}.worker.js`);
      assert.equal(worker.status, 200);
      assert.match(worker.headers.get("content-type"), /javascript/);
      assert.ok((await worker.text()).length > 1000);
      const wasm = await fetch(`${base}/haybarn/duckdb-${bundle}.wasm`, { method: "HEAD" });
      assert.equal(wasm.status, 200);
      assert.equal(wasm.headers.get("content-type"), "application/wasm");
      assert.ok(Number(wasm.headers.get("content-length")) > 1000);
    }
    const pthread = await fetch(`${base}/haybarn/duckdb-browser-coi.pthread.worker.js`, { method: "HEAD" });
    assert.equal(pthread.status, 200);
    const missing = await fetch(`${base}/haybarn/not-an-asset.js`);
    assert.equal(missing.status, 404);
    await missing.text();
    const headers = { Origin: "https://excel.officeapps.live.com" };
    const metadata = await fetch(`${base}/functions.json`, { headers });
    assert.equal(metadata.headers.get("access-control-allow-origin"), "*");
    assert.ok((await metadata.json()).functions.some(f => f.id === "QUERY"));
    const source = await fetch(`${base}/src/main.tsx`, { headers });
    assert.equal(source.headers.get("access-control-allow-origin"), null);
    // Consume the response before closing the server that is streaming it.
    await source.text();
  } finally { await server.close(); await rm(cacheDir, { recursive: true, force: true }); }
});
