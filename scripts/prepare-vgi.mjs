import { copyFile, mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';
import { artifactName, prepareArtifact, vgiLock } from './lib/vgi-artifacts.mjs';
const option = name => process.argv.find(arg => arg.startsWith(`--${name}=`))?.slice(name.length + 3);
const platforms = option('platform') ? [option('platform')] : Object.keys(vgiLock.artifacts);
const output = option('output');
for (const platform of platforms) {
  const file = await prepareArtifact(platform, { sourceDirectory: option('source-dir') });
  if (output) { await mkdir(resolve(output, platform), { recursive: true }); await copyFile(file, resolve(output, platform, artifactName(platform))); }
  console.log(`Verified VGI ${vgiLock.extensionVersion} ${platform}: ${file}`);
}
