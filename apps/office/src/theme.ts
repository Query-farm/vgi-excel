export interface HostTheme { bodyBackgroundColor?: string; bodyForegroundColor?: string; isDarkTheme?: boolean }

export function resolveTheme(theme?: HostTheme): { mode: "light" | "dark"; background: string; foreground: string } {
  const color = (value?: string) => /^#[a-f\d]{6}$/i.test(value ?? "") ? value! : undefined;
  const background = color(theme?.bodyBackgroundColor);
  const rgb = background ? [1, 3, 5].map(index => parseInt(background.slice(index, index + 2), 16)) : undefined;
  const dark = rgb ? rgb[0]! * .299 + rgb[1]! * .587 + rgb[2]! * .114 < 128 : theme?.isDarkTheme === true;
  return { mode: dark ? "dark" : "light", background: background ?? (dark ? "#292929" : "#ffffff"), foreground: color(theme?.bodyForegroundColor) ?? (dark ? "#f5f5f5" : "#242424") };
}

export function applyTheme(theme?: HostTheme): void {
  const resolved = resolveTheme(theme);
  const root = document.documentElement;
  root.dataset.officeTheme = resolved.mode;
  root.style.colorScheme = resolved.mode;
  root.style.setProperty("--office-background", resolved.background);
  root.style.setProperty("--office-foreground", resolved.foreground);
}

export function followOfficeTheme(): () => void {
  const update = () => {
    try { applyTheme(typeof Office === "undefined" ? undefined : Office.context?.officeTheme); }
    catch { applyTheme(); }
  };
  update();
  // Theme-change events aren't exposed consistently by Excel hosts. Re-read the
  // inexpensive host property without re-rendering or disturbing active work.
  const timer = window.setInterval(update, 2000);
  window.addEventListener("focus", update);
  return () => { clearInterval(timer); window.removeEventListener("focus", update); };
}
