import type { CatalogFunction } from "./types.js";

const MAX_NAME_LENGTH = 240;

export function wrapperName(fn: CatalogFunction, occupied: ReadonlySet<string> = new Set()): string {
  const raw = `VGI_${fn.catalog}_${fn.schema}_${fn.name}`.toUpperCase();
  let name = raw.replace(/[^A-Z0-9_.]/g, "_").replace(/_+/g, "_");
  if (!/^[A-Z_\\]/.test(name)) name = `VGI_${name}`;
  if (name.length > MAX_NAME_LENGTH) name = name.slice(0, MAX_NAME_LENGTH);
  if (!occupied.has(name)) return name;
  return `${name.slice(0, MAX_NAME_LENGTH - 9)}_${stableHash(`${fn.catalog}.${fn.schema}.${fn.name}`)}`;
}

export function wrapperFormula(fn: CatalogFunction): string {
  if (fn.kind !== "scalar") throw new Error("Only scalar VGI functions can be exposed as LAMBDA wrappers.");
  const args = fn.parameters.map((p, index) => safeParameterName(p.name, index));
  const fqName = `${fn.catalog}.${fn.schema}.${fn.name}`.replaceAll('"', '""');
  const call = [`"${fqName}"`, ...args].join(",");
  return `=LAMBDA(${[...args, `VGI.CALL(${call})`].join(",")})`;
}

function safeParameterName(value: string, index: number): string {
  let result = value.replace(/[^A-Za-z0-9_.]/g, "_");
  if (!/^[A-Za-z_\\]/.test(result) || /^[A-Z]{1,3}[0-9]+$/i.test(result)) result = `arg_${index + 1}`;
  return result || `arg_${index + 1}`;
}

function stableHash(value: string): string {
  let hash = 2166136261;
  for (let i = 0; i < value.length; i++) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, "0").toUpperCase();
}
