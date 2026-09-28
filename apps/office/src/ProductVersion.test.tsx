import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { AboutContent } from "./ProductVersion";
import product from "../../../package.json";

describe("Cupola product chrome", () => {
  it("shows the shipped release version", () => {
    const html = renderToStaticMarkup(<AboutContent diagnostics="test diagnostics" onCopy={() => {}}/>);
    expect(html).toContain(`Version ${product.version}`);
    expect(html).toContain(product.cupolaBuild);
    expect(html).toContain('href="https://query.farm"');
  });
});
