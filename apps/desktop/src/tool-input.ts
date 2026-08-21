export function parseStreamedToolInput(raw: string): { input: Record<string, unknown>; parseError: boolean } {
  if (raw.trim() === "") return { input: {}, parseError: false };
  try { return { input: JSON.parse(raw) as Record<string, unknown>, parseError: false }; }
  catch { return { input: { __parseError: raw }, parseError: true }; }
}
