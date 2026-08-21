import { assertHttpsConnection, FormulaService, type QueryBackend } from "@query-farm/vgi-excel-core";
import { BrowserBackend } from "./browser-backend";
import { getDefaultConnectionName, loadConnections } from "./config";

const backends = new Map<string, QueryBackend>();

export async function resolveBackend(connectionName?: string): Promise<QueryBackend> {
  const definitions = loadConnections();
  const name = connectionName || getDefaultConnectionName();
  const definition = definitions.find((item) => item.name === name);
  if (!definition) throw new Error("No VGI connection is configured. Open the VGI task pane and add one.");
  const key = JSON.stringify(definition);
  const existing = backends.get(key);
  if (existing) return existing;

  assertHttpsConnection(definition);
  const backend = new BrowserBackend(definition);
  backends.set(key, backend);
  return backend;
}

export const formulaService = new FormulaService(resolveBackend);

export function resetRuntime(): void {
  formulaService.clear();
  backends.clear();
}

window.addEventListener("vgi-connections-changed", resetRuntime);
