import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
const product = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));
const expectedRelease = process.env.CUPOLA_OFFICE_EXPECTED_RELEASE ?? `${product.version}-${product.cupolaBuild}`;
const origin = process.env.CUPOLA_OFFICE_BASE_URL ?? 'https://cupola.query.farm';
const health = await fetch(`${origin}/health.json`);
assert.equal(health.status,200);
const { release } = await health.json();
assert.match(release,/^\d+\.\d+\.\d+-\d{8}\.\d+$/);
assert.equal(release,expectedRelease,'The public host must serve the expected release.');
const manifest = await fetch(`${origin}/manifest.xml`);
assert.equal(manifest.status,200);
const xml = await manifest.text();
assert(!/localhost|query-farm\.services/.test(xml));
assert(xml.includes(`${origin}/releases/${release}/taskpane.html`));
for (const url of new Set([...xml.matchAll(/DefaultValue="(https:[^"]+)"/g)].map(match => match[1]).filter(url => url.startsWith(origin)))) {
 const response=await fetch(url,{method:'HEAD'});
 assert.equal(response.status,200,url);
}
const wasm=await fetch(`${origin}/releases/${release}/haybarn/duckdb-coi.wasm`,{method:'HEAD',headers:{'Accept-Encoding':'identity'}});
assert.equal(wasm.status,200);assert.equal(wasm.headers.get('content-type'),'application/wasm');
assert(Number(wasm.headers.get('content-length'))>25*1024*1024);
assert.equal(wasm.headers.get('cross-origin-embedder-policy'),'require-corp');
const metadata=await fetch(`${origin}/functions.json`,{headers:{Origin:'https://excel.officeapps.live.com'}});
assert.equal(metadata.headers.get('access-control-allow-origin'),'*');
assert.equal(metadata.headers.get('access-control-allow-credentials'),null);
for (const path of ['/src/main.tsx','/.env',`/releases/${release}/assets/source.js.map`]) assert.equal((await fetch(origin+path)).status,404);
console.log(`PASS: ${release}, production manifest URLs, HTTPS assets, WASM MIME/size, metadata CORS, and private-file exclusion.`);
