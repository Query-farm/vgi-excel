const UNRECOVERABLE = /token exchange failed|token refresh failed|invalid_grant|AADSTS\d+/i;
const RECOVERABLE = [
  /\b401\b/,
  /\b403\b/,
  /unauthori[sz]/i,
  /unauthenticated/i,
  /authenticat/i,
  /\boauth\b/i,
  /invalid[_ ]token/i,
  /(token[^.]{0,20}expired|expired[^.]{0,20}token)/i,
];

export function isRecoverableAuthError(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error ?? "");
  return !!message && !UNRECOVERABLE.test(message) && RECOVERABLE.some((pattern) => pattern.test(message));
}
