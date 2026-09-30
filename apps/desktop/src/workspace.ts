import type { DesktopConnection } from "./bridge";

export function workspaceConnection(values: DesktopConnection[]): DesktopConnection | undefined {
  return values.find(value => value.isWorkspaceSelected) ?? values.find(value => value.isDefault) ?? values.find(value => !value.isWorkspaceProfile);
}

export function workspaceMembers(values: DesktopConnection[]): string[] {
  const current = workspaceConnection(values);
  return current ? current.members?.length ? current.members : [current.name] : [];
}

export function workspaceSelectionError(members: string[], values: DesktopConnection[]): string {
  if (!members.length) return "Select at least one catalog.";
  if (members.length > 16) return "Choose at most 16 catalogs.";
  const catalogs = members.map(name => values.find(value => value.name === name && !value.members?.length)?.catalog.toLowerCase());
  if (catalogs.some(catalog => !catalog)) return "A selected connection is no longer available.";
  if (new Set(catalogs).size !== catalogs.length) return "These connections use the same catalog name. Choose connections with different catalog names.";
  return "";
}

export function sameWorkspace(left: string[], right: string[]): boolean {
  return left[0] === right[0] && left.length === right.length && left.every(name => right.includes(name));
}
