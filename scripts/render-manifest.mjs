import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';

const value = name => process.argv.find(arg => arg.startsWith(`--${name}=`))?.slice(name.length + 3);
const rawBase = value('base-url');
if (!rawBase) throw new Error('Usage: npm run manifest -- --base-url=https://cupola.example.com');
const base = new URL(rawBase);
if (base.protocol !== 'https:' || base.username || base.password || base.search || base.hash || base.pathname !== '/') {
  throw new Error('The production add-in base URL must be an HTTPS origin.');
}
const assetPath = value('asset-path') ?? '/';
if (!/^\/(?:releases\/[0-9]+\.[0-9]+\.[0-9]+-[0-9]{8}\.[0-9]+\/)?$/.test(assetPath)) throw new Error('Invalid asset path.');
const path = resolve('apps/office/dist/manifest.xml');
const manifest = await readFile(path, 'utf8');
await writeFile(path, manifest.replaceAll('https://localhost:3000/', `${base.origin}${assetPath}`));
console.log(`Rendered ${path} for ${base.origin}${assetPath}`);
