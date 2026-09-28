import { describe, expect, it } from "vitest";
import { catalogDraft, catalogNames, connectionNameError, catalogDiscoveryError } from "./connection-draft.js";

describe("connection drafts", () => {
  it("defaults an untouched name and preserves an explicit or saved identity", () => {
    const draft = { name: "", catalog: "", location: "https://example.de" };
    expect(catalogDraft(draft, "weather", false)).toEqual({ ...draft, name: "weather", catalog: "weather" });
    expect(catalogDraft({ ...draft, name: "My data" }, "weather", true).name).toBe("My data");
    expect(catalogDraft(draft, "weather", false).location).toBe(draft.location);
  });
  it("rejects duplicate names without rejecting an edit of the same identity", () => {
    expect(connectionNameError(" WEATHER ", "", [{ name: "Weather" }])).toContain("already exists");
    expect(connectionNameError("Weather", "Weather", [{ name: "Weather" }])).toBe("");
    expect(connectionNameError("Weather", "Other", [{ name: "Weather" }])).toContain("already exists");
  });
  it("preserves catalog spelling and removes empty, invalid, and duplicate results", () => {
    expect(catalogNames([["Weather"], [null], [42], [""], ["Weather"], ["weather"]])).toEqual(["Weather", "weather"]);
  });
});


describe("catalog discovery errors", () => {
  it("explains protected discovery without suggesting an ineffective sign-in retry", () => {
    for (const status of [401, 403]) {
      const error = new Error(`IO Error: VGI HTTP authentication required (HTTP ${status}) but no auth configured for this catalog [url: https://private.example.test/vgi.v2/catalog_catalogs]. Use bearer_token secret or oauth_refresh_token secret in ATTACH options.`);
      expect(catalogDiscoveryError(error)).toBe("Cupola couldn’t sign in to list catalogs. Try Find catalogs again, or enter the catalog name manually.");
    }
  });
  it("shows a useful timeout message", () => {
    expect(catalogDiscoveryError(new Error("Connection test timed out after 20 seconds"))).toContain("Finding catalogs timed out");
  });
  it("keeps raw server, credential, and SQL details out of discovery messages", () => {
    const result = catalogDiscoveryError(new Error("SELECT catalog FROM vgi_catalogs('https://private.example.test'); password=secret"));
    expect(result).toBe("Cupola couldn’t list the catalogs. Check the server address and try again, or enter the catalog name manually.");
  });
});
