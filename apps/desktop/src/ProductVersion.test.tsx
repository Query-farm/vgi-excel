import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { ProductVersion } from "./ProductVersion";

describe("Cupola product chrome", () => {
  it("shows the shipped release version", () => {
    const html = renderToStaticMarkup(<ProductVersion/>);
    expect(html).toContain("v0.4.0");
    expect(html).toContain("product-footer");
  });
});
