import { describe, expect, it } from "vitest";
import { officeTreeModel, type CatalogObject } from "./catalog-tree";

describe("officeTreeModel", () => {
  it("puts tables first, sorts names, and hides dollar-sign objects by default", () => {
    const objects: CatalogObject[] = [
      { catalog: "weather", schema: "main", name: "forecast_current", kind: "function" },
      { catalog: "weather", schema: "main", name: "zeta_stations", kind: "table" },
      { catalog: "weather", schema: "main", name: "Alpha_stations", kind: "table" },
      { catalog: "weather", schema: "main", name: "sys$statistics", kind: "table" },
      { catalog: "weather", schema: "main", name: "city_view", kind: "view" },
    ];

    const model = officeTreeModel(objects);
    expect(model[0].folders.map((folder) => folder.label)).toEqual(["Tables", "Views", "Functions"]);
    expect(model[0].folders[0].values.map((value) => value.name)).toEqual(["Alpha_stations", "zeta_stations"]);
    expect(officeTreeModel(objects, true)[0].folders[0].values.map((value) => value.name)).toEqual(["Alpha_stations", "sys$statistics", "zeta_stations"]);
  });
});
