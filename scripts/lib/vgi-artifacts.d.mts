export interface VgiLock {
  schemaVersion: number;
  extensionVersion: string;
  engineVersion: string;
  wasmPackageVersion: string;
  artifacts: Record<string, { url: string; compression: string; sha256: string; size: number }>;
}
export const repositoryRoot: string;
export const vgiLock: VgiLock;
export const wasmPlatforms: Record<'mvp' | 'eh' | 'coi', string>;
export function artifactName(platform: string): string;
export function assetPath(platform: string, lock?: VgiLock): string;
export function verifyArtifact(bytes: Buffer, platform: string, lock?: VgiLock): void;
export function prepareArtifact(platform: string, options?: { sourceDirectory?: string; cacheDirectory?: string; fetchImpl?: typeof fetch; lock?: VgiLock }): Promise<string>;
