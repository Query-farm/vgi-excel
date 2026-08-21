const INVALID_WORKSHEET_CHARACTERS = /[\\/?*\[\]:\u0000-\u001f]/g;

/** Normalize an untrusted/user/AI supplied label to an Excel worksheet name. */
export function excelWorksheetName(value: unknown, fallback = "VGI Result"): string {
  let name = String(value ?? "").trim()
    .replace(INVALID_WORKSHEET_CHARACTERS, " - ")
    .replace(/\s+/g, " ")
    .replace(/^'+|'+$/g, "")
    .trim();
  if (!name) name = String(fallback || "VGI Result").replace(INVALID_WORKSHEET_CHARACTERS, " ").trim() || "VGI Result";
  if (/^history$/i.test(name)) name = "History Data";
  name = Array.from(name).slice(0, 31).join("").trim().replace(/'+$/g, "").trim();
  return name || "VGI Result";
}
