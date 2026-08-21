export interface QueryDocument {
  id: string;
  name: string;
  sql: string;
  createdAt: number;
  updatedAt: number;
}

export interface QueryDocumentState {
  version: 1;
  documents: QueryDocument[];
  activeId: string;
}

export interface QueryDocumentStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

const STORAGE_PREFIX = "cupola.excel.query-documents.v1";
export const DEFAULT_QUERY_SQL = "SELECT current_catalog(), current_schema();";

function newId(): string {
  try {
    if (typeof crypto !== "undefined" && crypto.randomUUID) return crypto.randomUUID();
  } catch { }
  return `query-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function newDocument(name: string, sql: string): QueryDocument {
  const now = Date.now();
  return { id: newId(), name, sql, createdAt: now, updatedAt: now };
}

function freshState(sql = DEFAULT_QUERY_SQL): QueryDocumentState {
  const document = newDocument("Query 1", sql);
  return { version: 1, documents: [document], activeId: document.id };
}

function storageKey(scope: string): string {
  return `${STORAGE_PREFIX}::${encodeURIComponent(scope || "default")}`;
}

function browserStorage(): QueryDocumentStorage | undefined {
  try { return typeof localStorage === "undefined" ? undefined : localStorage; }
  catch { return undefined; }
}

export function loadQueryDocumentState(scope: string, storage: QueryDocumentStorage | undefined = browserStorage(), initialSql = DEFAULT_QUERY_SQL): QueryDocumentState {
  if (!storage) return freshState(initialSql);
  try {
    const raw = storage.getItem(storageKey(scope));
    if (!raw) return freshState(initialSql);
    const parsed = JSON.parse(raw) as Partial<QueryDocumentState>;
    const ids = new Set<string>();
    const documents = Array.isArray(parsed.documents) ? parsed.documents.filter((value): value is QueryDocument => {
      if (!value || typeof value.id !== "string" || ids.has(value.id) || typeof value.sql !== "string") return false;
      ids.add(value.id);
      return true;
    }).map((value, index) => ({
      id: value.id,
      name: typeof value.name === "string" && value.name.trim() ? value.name.trim() : `Query ${index + 1}`,
      sql: value.sql,
      createdAt: Number.isFinite(value.createdAt) ? value.createdAt : Date.now(),
      updatedAt: Number.isFinite(value.updatedAt) ? value.updatedAt : Date.now(),
    })) : [];
    if (!documents.length) return freshState(initialSql);
    const activeId = documents.some((value) => value.id === parsed.activeId) ? parsed.activeId! : documents[0].id;
    return { version: 1, documents, activeId };
  } catch { return freshState(initialSql); }
}

export function saveQueryDocumentState(scope: string, state: QueryDocumentState, storage: QueryDocumentStorage | undefined = browserStorage()): void {
  if (!storage) return;
  try { storage.setItem(storageKey(scope), JSON.stringify(state)); } catch { }
}

export function addQueryDocument(state: QueryDocumentState, sql = "", name?: string): QueryDocumentState {
  const document = newDocument(name?.trim() || nextName(state.documents), sql);
  return { ...state, documents: [...state.documents, document], activeId: document.id };
}

export function removeQueryDocument(state: QueryDocumentState, id: string): QueryDocumentState {
  const removedIndex = state.documents.findIndex((value) => value.id === id);
  if (removedIndex < 0) return state;
  const documents = state.documents.filter((value) => value.id !== id);
  if (!documents.length) return freshState("");
  const activeId = state.activeId === id ? documents[Math.min(removedIndex, documents.length - 1)].id : state.activeId;
  return { ...state, documents, activeId };
}

export function renameQueryDocument(state: QueryDocumentState, id: string, name: string): QueryDocumentState {
  const trimmed = name.trim();
  if (!trimmed) return state;
  return { ...state, documents: state.documents.map((value) => value.id === id ? { ...value, name: trimmed, updatedAt: Date.now() } : value) };
}

export function updateQueryDocumentSql(state: QueryDocumentState, id: string, sql: string): QueryDocumentState {
  return { ...state, documents: state.documents.map((value) => value.id === id ? { ...value, sql, updatedAt: Date.now() } : value) };
}

export function activateQueryDocument(state: QueryDocumentState, id: string): QueryDocumentState {
  return state.documents.some((value) => value.id === id) ? { ...state, activeId: id } : state;
}

function nextName(documents: QueryDocument[]): string {
  const used = new Set(documents.flatMap((value) => {
    const match = /^Query (\d+)$/.exec(value.name);
    return match ? [Number(match[1])] : [];
  }));
  let index = 1;
  while (used.has(index)) index++;
  return `Query ${index}`;
}
