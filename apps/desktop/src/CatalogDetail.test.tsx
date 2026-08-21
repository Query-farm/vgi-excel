import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { CatalogDetail, CatalogTree, catalogTreeModel, type CatalogDetails } from "./App";

describe("CatalogDetail", () => {
  it("builds a catalog, schema, and object-type tree with per-branch expansion", () => {
    const result = {
      columns: ["catalog", "schema", "name", "object_type"].map((name) => ({ name, type: "VARCHAR" })),
      rows: [
        ["weather", "main", "forecast_current", "function"],
        ["weather", "main", "cities", "table"],
        ["weather", "main", "zeta_stations", "table"],
        ["weather", "main", "Alpha_stations", "table"],
        ["weather", "main", "sys$statistics", "table"],
        ["weather", "reference", "country_codes", "view"],
      ],
      rowCount: 6,
    };
    const model = catalogTreeModel(result);
    expect(model.map((schema) => schema.name)).toEqual(["main", "reference"]);
    expect(model[0].folders.map((folder) => folder.label)).toEqual(["Tables", "Functions"]);
    expect(model[0].folders[0].rows.map((row) => row[2])).toEqual(["Alpha_stations", "cities", "zeta_stations"]);
    expect(catalogTreeModel(result, true)[0].folders[0].rows.map((row) => row[2])).toEqual(["Alpha_stations", "cities", "sys$statistics", "zeta_stations"]);
    const html = renderToStaticMarkup(<CatalogTree result={result} catalog="weather" onSelect={() => undefined}/>);
    expect(html).toContain('role="tree"');
    expect(html).toContain('aria-expanded="false"');
    expect(html).toContain("Show hidden (1)");
    expect(html).not.toContain("Expand all");
    expect(html).not.toContain("Collapse all");
    const expandedHtml = renderToStaticMarkup(<CatalogTree result={result} catalog="weather" forceExpanded onSelect={() => undefined}/>);
    expect(expandedHtml).toContain('aria-label="cities, table"');
    expect(expandedHtml).not.toContain("<small>");
  });

  it("renders rich VGI signatures, docs, result columns, and examples", () => {
    const tags = {
      "vgi.category": "forecast",
      "vgi.doc_md": "## Current weather\n\nReturns **current conditions**.",
      "vgi.result_columns_schema": JSON.stringify([{ name: "temperature_2m", type: "DOUBLE", description: "Temperature" }]),
      "vgi.example_queries": JSON.stringify([{ description: "Boston", sql: "SELECT * FROM open_meteo.main.forecast_current(42.3, -71.0)" }]),
    };
    const value: CatalogDetails = {
      title: "open_meteo.main.forecast_current",
      metadata: {
        columns: ["function_type", "description", "return_type", "parameters", "parameter_types", "examples", "tags"].map((name) => ({ name, type: "VARCHAR" })),
        rows: [["table", "Current weather", null, '["latitude","temperature_unit"]', '["DOUBLE","VARCHAR"]', "[]", JSON.stringify(tags)]], rowCount: 1,
      },
      fields: {
        columns: ["arg_position", "arg_name", "arg_type", "kind", "arg_choices"].map((name) => ({ name, type: "VARCHAR" })),
        rows: [[0, "latitude", "DOUBLE", "positional", null], [null, "temperature_unit", "VARCHAR", "named", '["celsius","fahrenheit"]']], rowCount: 2,
      },
    };
    const html = renderToStaticMarkup(<CatalogDetail value={value}/>);
    expect(html).toContain("temperature_unit := VARCHAR");
    expect(html).toContain("<strong>current conditions</strong>");
    expect(html).toContain("temperature_2m");
    expect(html).toContain("SELECT * FROM open_meteo.main.forecast_current");
  });
});
