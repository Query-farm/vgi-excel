import { expect, it } from "vitest";
import { resolveTheme } from "./theme";
it("defaults to light when Office cannot supply a theme", () => {
  expect(resolveTheme()).toEqual({ mode: "light", background: "#ffffff", foreground: "#242424" });
});
it("uses Office body colors rather than the OS preference", () => {
  expect(resolveTheme({ bodyBackgroundColor: "#ffffff", bodyForegroundColor: "#222222" }).mode).toBe("light");
  expect(resolveTheme({ bodyBackgroundColor: "#292929", bodyForegroundColor: "#eeeeee" })).toEqual({ mode: "dark", background: "#292929", foreground: "#eeeeee" });
});
it("rejects invalid colors and handles older host theme flags", () => {
  expect(resolveTheme({ isDarkTheme: true, bodyBackgroundColor: "url(https://invalid.example)" }).background).toBe("#292929");
  expect(resolveTheme({ bodyBackgroundColor: "bad" }).mode).toBe("light");
});
