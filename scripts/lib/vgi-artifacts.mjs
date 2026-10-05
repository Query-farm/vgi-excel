import { createHash } from 'node:crypto';
import { readFile, mkdir, writeFile, rename, rm } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { gunzipSync } from 'node:zlib';

export const repositoryRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
export const vgiLock = JSON.parse(readFileSync(resolve(repositoryRoot, 'vgi-extensions.lock.json'), 'utf8'));
export const wasmPlatforms = { mvp: 'wasm_mvp', eh: 'wasm_eh', coi: 'wasm_threads' };
export function artifactName(platform) { return `vgi.duckdb_extension${platform.startsWith('wasm_') ? '.wasm' : ''}`; }
export function assetPath(platform, lock = vgiLock) { return `vgi/${lock.artifacts[platform].sha256}/${platform}/${artifactName(platform)}`; }
export function verifyArtifact(bytes, platform, lock = vgiLock) {
  const expected = lock.artifacts[platform];
  if (!expected || bytes.length !== expected.size || createHash('sha256').update(bytes).digest('hex') !== expected.sha256) {
    throw new Error(`VGI ${platform} checksum mismatch. Select and approve the artifact in vgi-extensions.lock.json; builds never follow latest automatically.`);
  }
  // Haybarn validates the signature when loading. Verify footer identity here too,
  // so the lock cannot accidentally combine the wrong platform/ABI/revision.
  const field = offset => bytes.subarray(bytes.length - offset, bytes.length - offset + 32).toString().replace(/\0/g, '');
  if (field(384) !== lock.extensionVersion || field(352) !== lock.engineVersion || field(320) !== platform) {
    throw new Error(`VGI ${platform} metadata does not match the release lock.`);
  }
}
export async function prepareArtifact(platform, { sourceDirectory, cacheDirectory = resolve(repositoryRoot, 'artifacts/vgi'), fetchImpl = fetch, lock = vgiLock } = {}) {
  const expected = lock.artifacts[platform];
  if (!expected || !/^[a-f0-9]{64}$/.test(expected.sha256)) throw new Error(`Unsupported VGI platform: ${platform}`);
  const target = resolve(cacheDirectory, expected.sha256, artifactName(platform));
  try { const cached = await readFile(target); verifyArtifact(cached, platform, lock); return target; }
  catch (error) { if (error.code !== 'ENOENT') throw error; }
  let bytes;
  if (sourceDirectory) bytes = await readFile(resolve(sourceDirectory, platform, artifactName(platform)));
  else {
    if (new URL(expected.url).protocol !== 'https:') throw new Error('VGI artifact URLs must use HTTPS.');
    const response = await fetchImpl(expected.url, { signal: AbortSignal.timeout(60_000) });
    if (!response.ok || (response.url && new URL(response.url).protocol !== 'https:')) throw new Error(`Could not download pinned VGI ${platform} (${response.status}). Supply an approved local copy with --source-dir.`);
    const downloaded = Buffer.from(await response.arrayBuffer());
    bytes = expected.compression === 'gzip' ? gunzipSync(downloaded, { maxOutputLength: expected.size }) : downloaded;
  }
  verifyArtifact(bytes, platform, lock);
  await mkdir(dirname(target), { recursive: true });
  const temporary = `${target}.${process.pid}.${crypto.randomUUID()}.tmp`;
  try { await writeFile(temporary, bytes); await rename(temporary, target); }
  finally { await rm(temporary, { force: true }); }
  return target;
}
