import { expect, it } from "vitest";
import { bundledVgiLoadSql } from "./browser-backend";
import { vgiLock, wasmPlatforms, assetPath } from "../../../scripts/lib/vgi-artifacts.mjs";

it.each(["mvp", "eh", "coi"] as const)("loads the packaged %s binary under the deployed release base", bundle => {
  const platform = wasmPlatforms[bundle];
  const sql = bundledVgiLoadSql(bundle, "https://cupola.example/releases/test/");
  expect(sql).toBe(`LOAD 'https://cupola.example/releases/test/${assetPath(platform)}'`);
  expect(sql).toContain(vgiLock.artifacts[platform].sha256);
  expect(sql).not.toContain("INSTALL");
  expect(sql).not.toContain("community");
});
