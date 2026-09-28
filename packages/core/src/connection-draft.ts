/** Keep friendly identities independent once the user has chosen a name. */
export function catalogDraft<T extends { name: string; catalog?: string }>(draft: T, catalog: string, nameEdited: boolean): T {
  return { ...draft, catalog, name: nameEdited ? draft.name : catalog };
}

export function connectionNameError(name: string, originalName: string, saved: readonly { name: string }[]): string {
  const key = name.trim().toLowerCase();
  return saved.some(item => item.name.toLowerCase() === key && item.name.toLowerCase() !== originalName.toLowerCase())
    ? "A connection with this name already exists. Choose a different name." : "";
}

export function catalogNames(rows: readonly (readonly unknown[])[]): string[] {
  return [...new Set(rows.map(row => row[0]).filter((value): value is string => typeof value === "string" && value.length > 0))];
}

/** Keep discovery errors actionable without exposing engine SQL or credentials. */
export function catalogDiscoveryError(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error ?? "");
  if (message.startsWith("Sign-in wasn’t completed.")) return "Sign-in wasn’t completed. Try Find catalogs again, or enter the catalog name manually.";
  if (/sign.in|oauth|invalid_grant|token refresh|\b40[13]\b|no auth configured|authentication required|unauthori[sz]ed/i.test(message)) {
    return "Cupola couldn’t sign in to list catalogs. Try Find catalogs again, or enter the catalog name manually.";
  }
  if (/timed out|timeout/i.test(message)) {
    return "Finding catalogs timed out. Check the server address and try again, or enter the catalog name manually.";
  }
  return "Cupola couldn’t list the catalogs. Check the server address and try again, or enter the catalog name manually.";
}
