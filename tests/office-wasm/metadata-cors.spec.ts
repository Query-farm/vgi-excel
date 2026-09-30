import { expect, test, officeOrigin } from "./test";

test("Excel can fetch public function metadata across origins without exposing dev sources", async ({ page, request, context }) => {
  const base = officeOrigin;
  const origin = "https://excel.officeapps.live.com";
  // Explicitly grant loopback access so this test isolates CORS from browser prompts.
  await context.grantPermissions(["local-network-access"], { origin });
  // Supply an isolated host document, without contacting Excel or requiring login.
  await page.route(`${origin}/cupola-metadata-test`, route => route.fulfill({
    contentType: "text/html", body: "<!doctype html><title>Metadata host</title>",
  }));
  await page.goto(`${origin}/cupola-metadata-test`);
  const ids = await page.evaluate(async url => {
    const response = await fetch(`${url}/functions.json`, { credentials: "omit" });
    if (!response.ok) throw new Error(`Metadata status ${response.status}`);
    const metadata = await response.json();
    return metadata.functions.map((entry: { id: string }) => entry.id);
  }, base);
  expect(ids).toEqual(expect.arrayContaining(["QUERY", "VALUE", "CALL"]));
  const preflight = await request.fetch(`${base}/functions.json`, {
    method: "OPTIONS", headers: { Origin: origin, "Access-Control-Request-Method": "GET" },
  });
  expect(preflight.status()).toBe(204);
  expect(preflight.headers()["access-control-allow-origin"]).toBe("*");
  expect(preflight.headers()["access-control-allow-credentials"]).toBeUndefined();
  const source = await request.get(`${base}/src/main.tsx`, { headers: { Origin: origin } });
  expect(source.headers()["access-control-allow-origin"]).toBeUndefined();
});
