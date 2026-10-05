import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { gzipSync } from 'node:zlib';
import { prepareArtifact, verifyArtifact, assetPath, vgiLock } from '../../scripts/lib/vgi-artifacts.mjs';

function fixture(platform = 'wasm_threads', compression = 'none') {
  const bytes = Buffer.alloc(1024);
  bytes.write('fixture', bytes.length - 384); bytes.write('v1.5.5', bytes.length - 352); bytes.write(platform, bytes.length - 320);
  const lock = { extensionVersion: 'fixture', engineVersion: 'v1.5.5', artifacts: { [platform]: {
    url: 'https://extensions.example.test/vgi', compression, size: bytes.length,
    sha256: createHash('sha256').update(bytes).digest('hex'),
  } } };
  return { bytes, lock };
}

test('verifies fetched bytes and reuses only a matching content-addressed cache', async () => {
  const { bytes, lock } = fixture();
  const cacheDirectory = await mkdtemp(join(tmpdir(), 'cupola-vgi-'));
  let requests = 0;
  const fetchImpl = async () => { requests++; return new Response(bytes); };
  try {
    const path = await prepareArtifact('wasm_threads', { cacheDirectory, lock, fetchImpl });
    assert.deepEqual(await readFile(path), bytes);
    assert.equal(await prepareArtifact('wasm_threads', { cacheDirectory, lock, fetchImpl }), path);
    assert.equal(requests, 1);
    await writeFile(path, 'tampered');
    await assert.rejects(prepareArtifact('wasm_threads', { cacheDirectory, lock, fetchImpl }), /checksum mismatch/);
    assert.equal(requests, 1, 'a corrupt approved cache is not silently replaced');
  } finally { await rm(cacheDirectory, { recursive: true, force: true }); }
});

test('a moving upstream URL cannot silently change release bytes', async () => {
  const { lock } = fixture(); const cacheDirectory = await mkdtemp(join(tmpdir(), 'cupola-vgi-'));
  try {
    await assert.rejects(prepareArtifact('wasm_threads', { cacheDirectory, lock, fetchImpl: async () => new Response('new upstream release') }), /checksum mismatch/);
  } finally { await rm(cacheDirectory, { recursive: true, force: true }); }
});

test('checks native gzip against the decompressed binary checksum', async () => {
  const { bytes, lock } = fixture('windows_amd64', 'gzip');
  const cacheDirectory = await mkdtemp(join(tmpdir(), 'cupola-vgi-'));
  try {
    const path = await prepareArtifact('windows_amd64', { cacheDirectory, lock, fetchImpl: async () => new Response(gzipSync(bytes)) });
    assert.deepEqual(await readFile(path), bytes);
  } finally { await rm(cacheDirectory, { recursive: true, force: true }); }
});

test('rejects the wrong engine, extension revision, platform, and insecure download source', async () => {
  const { bytes, lock } = fixture();
  assert.throws(() => verifyArtifact(bytes, 'wasm_threads', { ...lock, engineVersion: 'v0.0.0' }), /metadata/);
  assert.throws(() => verifyArtifact(bytes, 'wasm_threads', { ...lock, extensionVersion: 'wrong' }), /metadata/);
  assert.throws(() => verifyArtifact(bytes, 'wasm_eh', { ...lock, artifacts: { wasm_eh: lock.artifacts.wasm_threads } }), /metadata/);
  const cacheDirectory = await mkdtemp(join(tmpdir(), 'cupola-vgi-'));
  lock.artifacts.wasm_threads.url = 'http://extensions.example.test/vgi';
  try { await assert.rejects(prepareArtifact('wasm_threads', { cacheDirectory, lock, fetchImpl: async () => { throw new Error('must not fetch'); } }), /HTTPS/); }
  finally { await rm(cacheDirectory, { recursive: true, force: true }); }
});

test('locks every shipped WASM variant and the Windows binary to one revision', () => {
  assert.deepEqual(Object.keys(vgiLock.artifacts).sort(), ['wasm_eh', 'wasm_mvp', 'wasm_threads', 'windows_amd64']);
  for (const [platform, value] of Object.entries(vgiLock.artifacts)) {
    assert.match(value.sha256, /^[a-f0-9]{64}$/);
    assert.ok(value.size > 512);
    assert.match(value.url, /^https:\/\//);
    assert.ok(assetPath(platform).includes(value.sha256));
  }
});
