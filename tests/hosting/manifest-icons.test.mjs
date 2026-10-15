import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

test('Office marketplace manifest icons have the required PNG dimensions', async () => {
  const manifest = await readFile('apps/office/public/manifest.xml', 'utf8');
  for (const [element, size] of [['IconUrl', 32], ['HighResolutionIconUrl', 64]]) {
    const match = manifest.match(new RegExp(`<${element} DefaultValue="([^"]+)"`));
    assert.ok(match, `${element} must be present`);
    const path = new URL(match[1]).pathname;
    assert.match(path, /^\/icons\/cupola-\d+\.png$/);
    const png = await readFile(`apps/office/public${path}`);
    assert.equal(png.subarray(0, 8).toString('hex'), '89504e470d0a1a0a');
    assert.equal(png.toString('ascii', 12, 16), 'IHDR');
    assert.equal(png.readUInt32BE(16), size, `${element} width`);
    assert.equal(png.readUInt32BE(20), size, `${element} height`);
  }
});
