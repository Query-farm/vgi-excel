import { test } from 'node:test';
import assert from 'node:assert/strict';
import worker from '../../cloudflare/office-worker.mjs';

const release = '0.5.0-20260930.2';
const origin = 'https://cupola.query.farm';
function setup() {
  const keys = [];
  const env = { RELEASE: release, OFFICE_ASSETS: { async get(key) {
    keys.push(key);
    if (key.includes('missing')) return null;
    return { body:'test', size:4, httpEtag:'"test-etag"', writeHttpMetadata(h) { h.set('Content-Type',key.endsWith('.wasm') ? 'application/wasm' : 'text/html'); } };
  } } };
  return { keys, fetch:(path, init) => worker.fetch(new Request(origin + path, init), env) };
}
test('entry redirects pin all subsequent relative engine and result requests to one release', async () => {
  const s = setup();
  const r = await s.fetch('/taskpane.html?mode=office');
  assert.equal(r.status,302);
  assert.equal(r.headers.get('Location'),`/releases/${release}/taskpane.html?mode=office`);
  assert.equal(r.headers.get('Cache-Control'),'no-store');
  assert.equal(s.keys.length,0);
});
test('streams versioned WASM with the required isolation headers and immutable caching', async () => {
  const s=setup();const r=await s.fetch(`/releases/${release}/haybarn/duckdb-coi.wasm`);
  assert.equal(r.status,200);assert.equal(await r.text(),'test');
  assert.equal(r.headers.get('Content-Type'),'application/wasm');
  assert.equal(r.headers.get('Cross-Origin-Opener-Policy'),'same-origin');
  assert.equal(r.headers.get('Cross-Origin-Embedder-Policy'),'require-corp');
  assert.match(r.headers.get('Cache-Control'),/immutable/);
  assert.deepEqual(s.keys,[`releases/${release}/haybarn/duckdb-coi.wasm`]);
});
test('Office metadata supports cross-origin GET and preflight without credentials', async () => {
  const s=setup();const r=await s.fetch('/functions.json',{method:'OPTIONS'});
  assert.equal(r.status,204);assert.equal(r.headers.get('Access-Control-Allow-Origin'),'*');
  assert.equal(r.headers.get('Access-Control-Allow-Credentials'),null);
  assert.equal(s.keys.length,0);
  assert.equal((await s.fetch('/functions.json')).headers.get('Cache-Control'),'no-cache');
});
test('stable OAuth callback serves HTML directly and never passes query strings to R2', async () => {
  const s=setup();const r=await s.fetch('/oauth-dialog.html?code=private-code&state=private-state');
  assert.equal(r.status,200);assert.equal(r.headers.get('Location'),null);
  assert.deepEqual(s.keys,[`releases/${release}/oauth-dialog.html`]);
  assert.equal(r.headers.get('Cache-Control'),'no-cache');
});
test('blocks non-public paths, source maps and writes without touching R2', async () => {
  const s=setup();
  for (const path of ['/src/App.tsx','/.env',`/releases/${release}/assets/code.js.map`,`/releases/${release}/.env.json`,`/releases/${release}/%2eprivate.json`]) {
    assert.equal((await s.fetch(path)).status,404,path);
  }
  assert.equal((await s.fetch(`/releases/${release}/taskpane.html`,{method:'PUT',body:'change'})).status,405);
  assert.equal(s.keys.length,0);
});
test('old releases remain accessible; HEAD, ETag and missing assets have correct responses', async () => {
  const s=setup();const path='/releases/0.5.0-20260930.1/taskpane.html';
  assert.equal((await s.fetch(path)).status,200);
  const head=await s.fetch(path,{method:'HEAD'});assert.equal(await head.text(),'');assert.equal(head.headers.get('Content-Length'),'4');
  const cached=await s.fetch(path,{headers:{'If-None-Match':'"test-etag"'}});assert.equal(cached.status,304);
  assert.equal((await s.fetch(`/releases/${release}/missing.js`)).status,404);
});
test('storage errors return a fixed response without reflecting credentials or paths', async () => {
 const env={RELEASE:release,OFFICE_ASSETS:{get(){throw new Error('secret credential SQL customer path');}}};
 const r=await worker.fetch(new Request(`${origin}/oauth-dialog.html?code=private`),env);
 assert.equal(r.status,503);assert.equal(await r.text(),'Asset temporarily unavailable');
});
