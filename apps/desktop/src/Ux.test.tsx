import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { Notice, Onboarding, WorkspaceTabs, formatSql, resultTsv } from "./Ux";

describe("Cupola UX primitives", () => {
  it("renders accessible tabs and onboarding", () => {
    const tabs = renderToStaticMarkup(<WorkspaceTabs value="query" tabs={[{ id: "query", label: "Query" }, { id: "catalog", label: "Catalog" }]} onChange={() => undefined}/>);
    expect(tabs).toContain('role="tablist"');
    expect(tabs).toContain('aria-selected="true"');
    expect(renderToStaticMarkup(<Onboarding onConnect={() => undefined}/>)).toContain("Connect a VGI data source to begin");
  });

  it("uses assertive error notices", () => {
    expect(renderToStaticMarkup(<Notice value={{ kind: "error", message: "Connection failed" }}/>)).toContain('role="alert"');
  });

  it("formats SQL and copies tabular results", () => {
    expect(formatSql("select * from data where x=1")).toContain("\nFROM data\nWHERE");
    expect(resultTsv({ columns: [{ name: "x" }], rows: [[1], [null]] })).toBe("x\n1\n");
  });
});
