export type CatalogObject = { catalog: string; schema: string; name: string; kind: string; description?: string; row?: Record<string, unknown> };

export type OfficeTreeFolder = { id: string; label: string; order: number; values: CatalogObject[] };
export type OfficeTreeSchema = { id: string; name: string; folders: OfficeTreeFolder[] };

export function officeTreeModel(objects: CatalogObject[], includeHidden = false): OfficeTreeSchema[] {
  const schemas = new Map<string, Map<string, OfficeTreeFolder>>();
  for (const value of objects) {
    if (!includeHidden && value.name.includes("$")) continue;
    const definition = value.kind === "table" ? { id: "tables", label: "Tables", order: 0 }
      : value.kind === "view" ? { id: "views", label: "Views", order: 1 }
      : value.kind.includes("function") ? { id: "functions", label: "Functions", order: 2 }
      : value.kind.includes("macro") ? { id: "macros", label: "Macros", order: 3 }
      : { id: "objects", label: "Other objects", order: 4 };
    const folders = schemas.get(value.schema) ?? new Map<string, OfficeTreeFolder>();
    const current = folders.get(definition.id) ?? { ...definition, values: [] };
    current.values.push(value);
    folders.set(definition.id, current);
    schemas.set(value.schema, folders);
  }
  return [...schemas]
    .sort(([left], [right]) => left.localeCompare(right, undefined, { sensitivity: "base" }))
    .map(([name, folders]) => {
      const values = [...folders.values()].sort((left, right) => left.order - right.order || left.label.localeCompare(right.label));
      for (const value of values) value.values.sort((left, right) => left.name.localeCompare(right.name, undefined, { sensitivity: "base" }));
      return { id: `schema:${name}`, name, folders: values };
    });
}
