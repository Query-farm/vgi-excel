import { describe, expect, it } from "vitest";
import { buildExcelAgentSystemPrompt } from "./prompt";

describe("multi-catalog prompt", () => {
  it("includes every attached catalog and excludes unrelated objects", () => {
    const prompt = buildExcelAgentSystemPrompt({
      connection: { name: "Research", catalog: "weather", catalogs: ["weather", "earthquakes"] },
      objects: ["weather", "earthquakes", "unrelated"].map(catalog => ({ catalog, schema: "main", name: "recent", kind: "table" })),
    });
    expect(prompt).toContain("Default catalog: weather");
    expect(prompt).toContain("weather.main.recent");
    expect(prompt).toContain("earthquakes.main.recent");
    expect(prompt).not.toContain("unrelated.main.recent");
  });
});
