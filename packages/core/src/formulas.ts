import type { CellMatrix, CellScalar, QueryBackend, QueryResult } from "./types.js";

export function normalizeMatrix(value: unknown): CellMatrix {
  if (Array.isArray(value)) {
    if (value.length === 0) return [[]];
    if (Array.isArray(value[0])) {
      return (value as unknown[][]).map((row) => row.map(normalizeScalar));
    }
    return [(value as unknown[]).map(normalizeScalar)];
  }
  return [[normalizeScalar(value)]];
}

export function normalizeScalar(value: unknown): CellScalar {
  if (value == null || value === "") return value == null ? null : "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  if (typeof value === "bigint") return value.toString();
  if (value instanceof Date) return value.toISOString();
  if (value instanceof Uint8Array) return bytesToBase64(value);
  return JSON.stringify(value);
}

export function resultMatrix(result: QueryResult, includeHeaders = true): CellMatrix {
  const rows = result.rows.map((row) => row.map(normalizeScalar));
  return includeHeaders ? [result.columns.map((column) => column.name), ...rows] : rows;
}

export function scalarResult(result: QueryResult): CellScalar {
  if (result.rows.length !== 1 || result.columns.length !== 1 || result.rows[0]?.length !== 1) {
    throw new Error("VGI.VALUE requires a query that returns exactly one row and one column.");
  }
  return normalizeScalar(result.rows[0][0]);
}

export function broadcastShape(args: CellMatrix[]): { rows: number; columns: number } {
  let rows = 1;
  let columns = 1;
  for (const arg of args) {
    const argRows = arg.length || 1;
    const argColumns = arg[0]?.length || 1;
    const scalar = argRows === 1 && argColumns === 1;
    if (scalar) continue;
    if ((rows !== 1 || columns !== 1) && (rows !== argRows || columns !== argColumns)) {
      throw new Error("Range arguments to VGI.CALL must have the same shape; scalar arguments are broadcast.");
    }
    rows = argRows;
    columns = argColumns;
  }
  return { rows, columns };
}

export function broadcastArguments(values: unknown[]): CellMatrix[] {
  const matrices = values.map(normalizeMatrix);
  const shape = broadcastShape(matrices);
  return matrices.map((matrix) => {
    if (matrix.length !== 1 || matrix[0]?.length !== 1) return matrix;
    return Array.from({ length: shape.rows }, () => Array(shape.columns).fill(matrix[0][0]));
  });
}

export class FormulaService {
  private inFlight = new Map<string, Promise<QueryResult>>();

  constructor(private readonly resolveBackend: (connection?: string) => Promise<QueryBackend>) {}

  async query(sql: string, connection?: string, includeHeaders = true, signal?: AbortSignal): Promise<CellMatrix> {
    const result = await this.runDeduplicated(`query:${connection ?? ""}:${sql}`, connection, (backend) =>
      backend.query(sql, { signal }),
    );
    return resultMatrix(result, includeHeaders);
  }

  async value(sql: string, connection?: string, signal?: AbortSignal): Promise<CellScalar> {
    const result = await this.runDeduplicated(`value:${connection ?? ""}:${sql}`, connection, (backend) =>
      backend.query(sql, { signal, maxRows: 2 }),
    );
    return scalarResult(result);
  }

  async call(functionName: string, values: unknown[], signal?: AbortSignal): Promise<CellMatrix> {
    const backend = await this.resolveBackend();
    const result = await backend.call(functionName, broadcastArguments(values), { signal });
    return resultMatrix(result, false);
  }

  clear(): void {
    this.inFlight.clear();
  }

  private async runDeduplicated(
    key: string,
    connection: string | undefined,
    run: (backend: QueryBackend) => Promise<QueryResult>,
  ): Promise<QueryResult> {
    const existing = this.inFlight.get(key);
    if (existing) return existing;
    const pending = this.resolveBackend(connection).then(run);
    this.inFlight.set(key, pending);
    try {
      return await pending;
    } finally {
      this.inFlight.delete(key);
    }
  }
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return typeof btoa === "function" ? btoa(binary) : binary;
}
