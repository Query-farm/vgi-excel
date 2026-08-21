import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { Notice, Onboarding, WorkspaceTabs } from "./Ux";

describe("Microsoft 365 Cupola shell", () => {
  it("has keyboard-addressable tab semantics and first-run guidance", () => {
    const tabs = renderToStaticMarkup(<WorkspaceTabs value="agent" tabs={[{ id: "query", label: "Query" }, { id: "agent", label: "Ask Cupola" }]} onChange={() => undefined}/>);
    expect(tabs).toContain('role="tablist"');
    expect(tabs).toContain('aria-controls="panel-agent"');
    expect(renderToStaticMarkup(<Onboarding onConnect={() => undefined}/>)).toContain("Add connection");
    expect(renderToStaticMarkup(<Notice value={{ kind: "progress", message: "Loading" }}/>)).toContain('aria-live="polite"');
  });
});
