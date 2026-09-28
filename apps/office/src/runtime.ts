import { assertHttpsConnection, FormulaService, type QueryBackend } from "@query-farm/vgi-excel-core";
import { BrowserBackend } from "./browser-backend";
import { getDefaultConnectionName, loadConnections } from "./config";

const backends = new Map<string, QueryBackend>();
const agentBackends = new Map<string, QueryBackend>();
const editorBackends = new Map<string, QueryBackend>();

export async function resolveBackend(connectionName?: string, editor: boolean | "agent" = false): Promise<QueryBackend> {
  const definitions = loadConnections();
  const name = connectionName || getDefaultConnectionName();
  const definition = definitions.find((item) => item.name === name);
  if (!definition) throw new Error("No VGI connection is configured. Open the VGI task pane and add one.");
  const key = JSON.stringify(definition);
  const cache = editor === "agent" ? agentBackends : editor ? editorBackends : backends;
  const existing = cache.get(key);
  if (existing) return existing;

  assertHttpsConnection(definition);
  const backend = new BrowserBackend(definition);
  cache.set(key, backend);
  return backend;
}

export const formulaService = new FormulaService(resolveBackend);

export function resetRuntime(): void {
  formulaService.clear();
  backends.clear();
  editorBackends.clear();
  agentBackends.clear();
}

window.addEventListener("vgi-connections-changed", resetRuntime);
