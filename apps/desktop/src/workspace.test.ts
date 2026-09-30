import { describe, expect, it } from "vitest";
import { sameWorkspace, workspaceConnection, workspaceMembers, workspaceSelectionError } from "./workspace";
import type { DesktopConnection } from "./bridge";
const sales: DesktopConnection = { name: "Sales", catalog: "sales", location: "https://sales.test", authentication: "anonymous", isDefault: true };
const inventory: DesktopConnection = { name: "Inventory", catalog: "inventory", location: "https://inventory.test", authentication: "anonymous" };
describe("workspace selection", () => {
  it("uses the workspace across hosts without switching legacy defaults", () => {
    const combined = { ...sales, name: "Sales + Inventory", members: ["Sales", "Inventory"], isDefault: false, isWorkspaceSelected: true, isWorkspaceProfile: true };
    expect(workspaceConnection([sales, inventory, combined])).toBe(combined);
    expect(workspaceMembers([sales, inventory, combined])).toEqual(["Sales", "Inventory"]);
    expect(workspaceMembers([sales, inventory])).toEqual(["Sales"]);
    expect(workspaceMembers([])).toEqual([]);
  });
  it("validates catalog selection before applying", () => {
    expect(workspaceSelectionError([], [sales])).toContain("at least one");
    expect(workspaceSelectionError(["Missing"], [sales])).toContain("no longer available");
    expect(workspaceSelectionError(["Sales", "Inventory"], [sales, inventory])).toBe("");
    expect(workspaceSelectionError(["Sales", "Inventory"], [sales, { ...inventory, catalog: "SALES" }])).toContain("same catalog name");
  });
  it("treats only the first catalog as the default", () => {
    expect(sameWorkspace(["A", "B", "C"], ["A", "C", "B"])).toBe(true);
    expect(sameWorkspace(["A", "B"], ["B", "A"])).toBe(false);
    expect(sameWorkspace(["A"], ["A", "B"])).toBe(false);
  });
});
