const READ_PREFIXES = new Set(["SELECT", "WITH", "DESCRIBE", "DESC", "SHOW", "EXPLAIN"]);

/** Strip comments and leading whitespace without attempting to rewrite SQL. */
export function firstSqlKeyword(sql: string): string {
  let value = sql.trimStart();
  while (value.startsWith("--") || value.startsWith("/*")) {
    if (value.startsWith("--")) {
      const end = value.indexOf("\n");
      value = end < 0 ? "" : value.slice(end + 1).trimStart();
    } else {
      const end = value.indexOf("*/", 2);
      value = end < 0 ? "" : value.slice(end + 2).trimStart();
    }
  }
  return value.match(/^([A-Za-z]+)/)?.[1]?.toUpperCase() ?? "";
}

/**
 * Conservative agent gate. The engine remains authoritative; this prevents
 * an AI tool call from intentionally submitting DDL/DML or stacked SQL.
 */
export function assertAgentReadOnlySql(sql: string): void {
  if (!READ_PREFIXES.has(firstSqlKeyword(sql))) {
    throw new Error("The AI agent may only run read-only SQL statements.");
  }
  if (hasStatementSeparator(sql)) {
    throw new Error("The AI agent may run only one SQL statement at a time.");
  }
  if (/\b(ATTACH|DETACH|INSTALL|LOAD|COPY|EXPORT|IMPORT|CALL|PRAGMA|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP|TRUNCATE|REPLACE|GRANT|REVOKE|VACUUM)\b/i.test(maskSqlLiterals(sql))) {
    throw new Error("The AI agent may not change data, connections, extensions, or external state.");
  }
}

function hasStatementSeparator(sql: string): boolean {
  const masked = maskSqlLiterals(sql).trim();
  const withoutTrailing = masked.endsWith(";") ? masked.slice(0, -1) : masked;
  return withoutTrailing.includes(";");
}

/** Replace quoted strings/identifiers and comments before keyword scanning. */
export function maskSqlLiterals(sql: string): string {
  let out = "";
  let i = 0;
  while (i < sql.length) {
    const c = sql[i];
    const next = sql[i + 1];
    if (c === "-" && next === "-") {
      const end = sql.indexOf("\n", i + 2);
      i = end < 0 ? sql.length : end;
      out += " ";
      continue;
    }
    if (c === "/" && next === "*") {
      const end = sql.indexOf("*/", i + 2);
      i = end < 0 ? sql.length : end + 2;
      out += " ";
      continue;
    }
    if (c === "'" || c === '"') {
      const quote = c;
      out += " ";
      i++;
      while (i < sql.length) {
        if (sql[i] === quote && sql[i + 1] === quote) {
          i += 2;
        } else if (sql[i] === quote) {
          i++;
          break;
        } else {
          i++;
        }
      }
      continue;
    }
    out += c;
    i++;
  }
  return out;
}

export function quoteIdentifier(value: string): string {
  return `"${value.replaceAll('"', '""')}"`;
}

export function quoteLiteral(value: string): string {
  return `'${value.replaceAll("'", "''")}'`;
}

export function qualifiedFunctionName(value: string): string {
  const parts = value.split(".");
  if (parts.length < 2 || parts.length > 3 || parts.some((p) => !p)) {
    throw new Error("Use a schema-qualified or catalog.schema-qualified function name.");
  }
  return parts.map(quoteIdentifier).join(".");
}
