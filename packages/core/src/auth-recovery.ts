/** These indicate interactive sign-in, not a generic permission or network failure. */
export function needsConnectionSignIn(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error ?? "");
  return /\b(?:invalid_grant|invalid_token|interaction_required|login_required)\b|\bAADSTS(?:700082|700084|70043|50173|50076|50079)\b|\bHTTP(?:\s+status)?\s*[:=(]?\s*401\b|\bunauthenticated\b|\bunauthori[sz]ed\b|authentication required|sign.in is required|(?:token|session)[^.\n]{0,40}expired|expired[^.\n]{0,40}(?:token|session)/i.test(message);
}

export class ConnectionSignInError extends Error {
  constructor(readonly connection: string) {
    super("Your connection needs you to sign in again. Open Connections in Settings to continue.");
    this.name = "ConnectionSignInError";
  }
}

// Memory only: no credentials, engine messages, or workbook metadata are retained.
let required: readonly string[] = [];
const listeners = new Set<() => void>();
export const connectionSignIn = {
  snapshot: () => required,
  subscribe(listener: () => void): () => void { listeners.add(listener); return () => { listeners.delete(listener); }; },
  require(name: string): void {
    if (required.includes(name)) return;
    required = [...required, name]; listeners.forEach(listener => listener());
  },
  clear(name: string): void {
    if (!required.includes(name)) return;
    required = required.filter(value => value !== name); listeners.forEach(listener => listener());
  },
  error(name: string, error: unknown): unknown {
    if (!needsConnectionSignIn(error)) return error;
    this.require(name);
    return new ConnectionSignInError(name);
  },
};
