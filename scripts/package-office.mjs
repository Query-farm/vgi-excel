import { readFile } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import { resolve } from "node:path";

const baseArg = process.argv.find((arg) => arg.startsWith("--base-url="));
if (!baseArg) throw new Error("Usage: npm run package:office -- --base-url=https://cupola.example.com");
const baseUrl = baseArg.slice("--base-url=".length).replace(/\/$/, "");
if (!baseUrl.startsWith("https://")) throw new Error("The production add-in base URL must use HTTPS.");

run(process.platform === "win32" ? "npm.cmd" : "npm", ["run", "build", "-w", "@query-farm/vgi-excel-office"]);
run(process.execPath, [resolve("scripts/render-manifest.mjs"), `--base-url=${baseUrl}`]);

const manifest = await readFile(resolve("apps/office/dist/manifest.xml"), "utf8");
if (manifest.includes("localhost")) throw new Error("The packaged Office manifest still contains localhost URLs.");
for (const required of ["duckdb-coi.wasm", "duckdb-browser-coi.worker.js"]) {
  await readFile(resolve("apps/office/dist/haybarn", required));
}
console.log("Office package ready in apps/office/dist.");
console.log("The production host must send Cross-Origin-Opener-Policy: same-origin and Cross-Origin-Embedder-Policy: require-corp.");

function run(command, args) {
  const result = spawnSync(command, args, { stdio: "inherit" });
  if (result.status !== 0) process.exit(result.status ?? 1);
}
