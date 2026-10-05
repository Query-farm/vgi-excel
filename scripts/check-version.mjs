import { readFile } from "node:fs/promises";
import { resolve } from "node:path";

const root = JSON.parse(await read("package.json"));
const version = root.version;
const build = root.cupolaBuild;
const manifestVersion = root.officeManifestVersion;
if (!/^[1-9]\d*\.\d+\.\d+\.\d+$/.test(manifestVersion)) throw new Error("officeManifestVersion must be a four-part version of at least 1.0.");
const checks = [
  ["windows/Cupola.Updater/Cupola.Updater.csproj", text => text.includes(`<Version>${version}</Version>`), `updater version ${version}`],
  ["installer/Package.wxs", text => text.includes(`Id="CUPOLABUILD" Value="${build}"`), `MSI build ${build}`],
  ["windows/Cupola.ExcelLoader/Cupola.ExcelLoader.csproj", (text) => text.includes(`<Version>${version}</Version>`) && text.includes(`<AssemblyVersion>${version}.0</AssemblyVersion>`), `Excel loader version ${version}`],
  ["apps/desktop/package.json", (text) => JSON.parse(text).version === version, `version ${version}`],
  ["apps/office/package.json", (text) => JSON.parse(text).version === version, `version ${version}`],
  ["packages/core/package.json", (text) => JSON.parse(text).version === version, `version ${version}`],
  ["apps/desktop/vite.config.ts", (text) => text.includes("product.version"), `web version derived from package.json (${version})`],
  ["apps/office/vite.config.ts", (text) => text.includes("product.version"), `web version derived from package.json (${version})`],
  ["windows/Vgi.ExcelDna/ProductInfo.cs", (text) => text.includes(`Version = "${version}"`), `native version ${version}`],
  ["windows/Vgi.ExcelDna/Vgi.ExcelDna.csproj", (text) => text.includes(`<Version>${version}</Version>`), `assembly version ${version}`],
  ["apps/office/public/manifest.xml", (text) => text.includes(`<Version>${manifestVersion}</Version>`), `manifest version ${manifestVersion}`],
  ["installer/Package.wxs", (text) => text.includes(`Name="Cupola for Excel"`) && text.includes(`Version="${version}"`), `MSI identity Cupola for Excel ${version}`],
  ["apps/desktop/vite.config.ts", (text) => text.includes("product.cupolaBuild"), `build derived from package.json (${build})`],
  ["apps/office/vite.config.ts", (text) => text.includes("product.cupolaBuild"), `build derived from package.json (${build})`],
  ["windows/Vgi.ExcelDna/ProductInfo.cs", (text) => text.includes(`Build = "${build}"`), `native build ${build}`],
];

for (const [path, matches, expectation] of checks) {
  const text = await read(path);
  if (!matches(text)) throw new Error(`${path} does not declare ${expectation}.`);
}
console.log(`PASS: Cupola for Excel version ${version} is consistent across all deliverables.`);

async function read(path) {
  return readFile(resolve(path), "utf8");
}
