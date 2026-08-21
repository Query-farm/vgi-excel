import { readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";

const baseArg = process.argv.find((arg) => arg.startsWith("--base-url="));
if (!baseArg) throw new Error("Usage: npm run manifest -- --base-url=https://vgi-excel.example.com");
const baseUrl = baseArg.slice("--base-url=".length).replace(/\/$/, "");
if (!baseUrl.startsWith("https://")) throw new Error("The production add-in base URL must use HTTPS.");
const path = resolve("apps/office/dist/manifest.xml");
const manifest = await readFile(path, "utf8");
await writeFile(path, manifest.replaceAll("https://localhost:3000", baseUrl));
console.log(`Rendered ${path} for ${baseUrl}`);
