export type CellScalar = string | number | boolean | null;
export type CellMatrix = CellScalar[][];

export interface QueryColumn {
  name: string;
  type: string;
}

export interface QueryResult {
  columns: QueryColumn[];
  rows: CellMatrix;
  rowCount: number;
  truncated?: boolean;
  elapsedMs?: number;
}

export interface QueryOptions {
  signal?: AbortSignal;
  maxRows?: number;
}

export interface QueryBackend {
  query(sql: string, options?: QueryOptions): Promise<QueryResult>;
  call(functionName: string, args: CellMatrix[], options?: QueryOptions): Promise<QueryResult>;
  cancel?(): Promise<void> | void;
}

export interface ConnectionDefinition {
  name: string;
  location: string;
  catalog?: string;
  attachOptions?: Record<string, CellScalar>;
}

export interface CatalogFunction {
  catalog: string;
  schema: string;
  name: string;
  description?: string;
  parameters: Array<{
    name: string;
    type: string;
    kind?: "positional" | "named" | "varargs" | "other" | "unknown";
    position?: number | null;
    description?: string;
    default?: unknown;
    choices?: unknown[];
    range?: string;
    pattern?: string;
  }>;
  returnType?: string;
  kind: "scalar" | "table" | "aggregate" | "table_in_out" | "buffering";
}
