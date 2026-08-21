import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { ChatMarkdown } from "./ChatMarkdown";

describe("Office ChatMarkdown", () => {
  it("renders headings, tables, and SQL blocks instead of raw Markdown", () => {
    const html = renderToStaticMarkup(<ChatMarkdown content={"## Forecast\n\n| city | °C |\n| --- | ---: |\n| Paris | 21 |\n\n```sql\nSELECT 21;\n```"}/>);
    expect(html).toContain("<h3>Forecast</h3>");
    expect(html).toContain("<table>");
    expect(html).toContain("language-sql");
    expect(html).toContain("Copy SQL");
    expect(html).not.toContain("```sql");
  });
});
