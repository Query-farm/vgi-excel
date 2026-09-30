import { createHash } from 'node:crypto';
import { readFile, readdir, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { resolve, extname } from 'node:path';

const origin = 'https://cupola.query.farm';
const assetsOrigin = 'https://cupola-assets.query.farm';
const bucket = 'cupola-excel-assets';
const product = JSON.parse(await readFile('package.json', 'utf8'));
const release = `${product.version}-${product.cupolaBuild}`;
if (!/^\d+\.\d+\.\d+-\d{8}\.\d+$/.test(release)) throw new Error('Invalid release identifier.');
const prefix = `releases/${release}`;
const dist = resolve('apps/office/dist');
const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm';
const wrangler = resolve('node_modules/wrangler/bin/wrangler.js');
const types = { '.html':'text/html; charset=utf-8', '.js':'text/javascript; charset=utf-8', '.css':'text/css; charset=utf-8', '.json':'application/json', '.xml':'application/xml', '.svg':'image/svg+xml', '.png':'image/png', '.wasm':'application/wasm' };
run(npm, ['run', 'package:office', '--', `--base-url=${origin}`, `--asset-path=/${prefix}/`]);
const files = [];
await walk(dist);
const manifest = await readFile(`${dist}/manifest.xml`, 'utf8');
if (!manifest.includes(`${origin}/${prefix}/taskpane.html`) || /localhost|query-farm\.services/.test(manifest)) throw new Error('Manifest is not production-ready.');
const inventory = { product: product.version, build: product.cupolaBuild, release, origin, files: files.sort((a,b) => a.path.localeCompare(b.path)) };
const inventoryText = JSON.stringify(inventory, null, 2) + '\n';
await writeFile(`${dist}/deployment.json`, inventoryText);
if (process.argv.includes('--prepare-only')) {
  console.log(`Prepared ${release}: ${files.length} assets. Nothing uploaded.`);
  process.exit(0);
}
const existing = await fetch(`${assetsOrigin}/${prefix}/deployment.json`, { cache: 'no-store' });
if (existing.ok) {
  if (await existing.text() !== inventoryText) throw new Error('This build already exists with different bytes. Increment cupolaBuild; releases are immutable.');
  console.log('Identical release already uploaded; verifying assets before deployment.');
} else {
  if (existing.status !== 404) throw new Error(`Could not inspect existing deployment (${existing.status}).`);
  for (const file of [...files, { path:'deployment.json', type:'application/json' }]) {
    run(process.execPath, [wrangler, 'r2', 'object', 'put', `${bucket}/${prefix}/${file.path}`, '--remote', '--file', `${dist}/${file.path}`, '--content-type', file.type, '--cache-control', 'public, max-age=31536000, immutable']);
  }
}
// Verify uploaded bytes, not just a successful upload response, before activation.
for (const file of files) {
  const response = await fetch(`${assetsOrigin}/${prefix}/${file.path}`);
  if (!response.ok || hash(Buffer.from(await response.arrayBuffer())) !== file.sha256) throw new Error(`Uploaded asset verification failed: ${file.path}`);
}
run(process.execPath, [wrangler, 'deploy', '--config', 'cloudflare/wrangler.jsonc', '--var', `RELEASE:${release}`]);
console.log(`Deployed ${release}. Manifest: ${origin}/manifest.xml`);

async function walk(dir, relative = '') {
  for (const item of await readdir(dir, { withFileTypes: true })) {
    const name = relative + item.name;
    if (item.isDirectory()) await walk(`${dir}/${item.name}`, name + '/');
    else {
      const type = types[extname(name)];
      if (!type || /(^|\/)\./.test(name) || item.isSymbolicLink()) throw new Error(`Unexpected deployment file: ${name}`);
      const bytes = await readFile(`${dir}/${item.name}`);
      files.push({ path:name, type, size:bytes.length, sha256:hash(bytes) });
    }
  }
}
function hash(bytes) { return createHash('sha256').update(bytes).digest('hex'); }
function run(command, args) {
  const result = spawnSync(command, args, { stdio:'inherit' });
  if (result.status !== 0) throw new Error(`Deployment command failed: ${command}`);
}
