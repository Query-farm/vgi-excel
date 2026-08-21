import type { CellScalar, ConnectionDefinition } from "./types.js";

const SENSITIVE_ATTACH_OPTIONS = new Set(["access_token", "api_key", "authorization", "bearer_token", "client_secret", "id_token", "oauth_refresh_token", "password", "refresh_token", "secret"]);
const MANAGED_ATTACH_OPTIONS = new Set(["location", "type"]);

export function assertHttpsConnection(definition: Pick<ConnectionDefinition, "name" | "location"> & Partial<Pick<ConnectionDefinition, "attachOptions">>): void {
  if (!definition.name.trim()) throw new Error("A connection name is required.");
  let url: URL;
  try {
    url = new URL(definition.location);
  } catch {
    throw new Error("VGI LOCATION must be a valid HTTPS URL.");
  }
  if (url.protocol !== "https:") throw new Error("Cupola for Excel supports HTTPS VGI endpoints only.");
  if (url.username || url.password) throw new Error("Credentials must not be embedded in the VGI URL.");
  for (const [key, value] of Object.entries(definition.attachOptions ?? {})) {
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) throw new Error(`Invalid ATTACH option name: ${key}`);
    const normalized = key.toLowerCase();
    if (SENSITIVE_ATTACH_OPTIONS.has(normalized)) throw new Error("Credentials must be supplied through VGI sign-in, not ATTACH options.");
    if (MANAGED_ATTACH_OPTIONS.has(normalized)) throw new Error(`${key.toUpperCase()} is managed by Cupola and must not be repeated in ATTACH options.`);
    if (value !== null && typeof value !== "string" && typeof value !== "number" && typeof value !== "boolean") {
      throw new Error(`ATTACH option ${key} must be a string, number, boolean, or null.`);
    }
  }
}

export function parseAttachOptionsJson(value: string): Record<string, CellScalar> {
  if (!value.trim()) return {};
  let parsed: unknown;
  try { parsed = JSON.parse(value); }
  catch { throw new Error("ATTACH options must be valid JSON."); }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) throw new Error("ATTACH options must be a JSON object.");
  const options = parsed as Record<string, CellScalar>;
  assertHttpsConnection({ name: "options", location: "https://options.invalid", attachOptions: options });
  return options;
}

export function formatAttachOptionsJson(value: Record<string, CellScalar> | undefined): string {
  return value && Object.keys(value).length ? JSON.stringify(value, null, 2) : "";
}
