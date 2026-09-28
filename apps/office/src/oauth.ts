import { sessionTokenKey, type OAuthTokens } from "./config";

export async function signIn(serviceUrl: string, signal?: AbortSignal): Promise<OAuthTokens> {
  signal?.throwIfAborted();
  const dialogUrl = new URL("/oauth-dialog.html", window.location.origin);
  dialogUrl.searchParams.set("service", serviceUrl);
  return new Promise((resolve, reject) => {
    let dialog: Office.Dialog | undefined;
    let settled = false;
    const finish = (error?: Error, tokens?: OAuthTokens) => {
      if (settled) return;
      settled = true;
      signal?.removeEventListener("abort", abort);
      try { dialog?.close(); } catch { /* The host may have closed it already. */ }
      if (error) reject(error); else resolve(tokens!);
    };
    const abort = () => finish(new DOMException("Sign-in cancelled.", "AbortError"));
    signal?.addEventListener("abort", abort, { once: true });
    if (signal?.aborted) { abort(); return; }
    try {
      Office.context.ui.displayDialogAsync(dialogUrl.toString(),
        { height: 65, width: 40, displayInIframe: false }, result => {
          if (settled) { try { result.value?.close(); } catch {} return; }
          if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
            finish(new Error("Unable to open sign-in. Allow the Cupola sign-in popup and try again.")); return;
          }
          dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, event => {
            if (settled) return;
            const message = event as { message: string; origin?: string };
            if (message.origin && message.origin !== window.location.origin) return;
            try {
              const payload = JSON.parse(message.message);
              if (payload.ok !== true) { finish(new Error("Sign-in was not completed. Try again.")); return; }
              const value = payload.tokens;
              if (!value || typeof value.access_token !== "string" || !value.access_token.trim() ||
                  (value.refresh_token !== undefined && typeof value.refresh_token !== "string")) {
                throw new Error("Sign-in returned an invalid session. Try again.");
              }
              const tokens: OAuthTokens = { access_token: value.access_token };
              if (value.refresh_token) tokens.refresh_token = value.refresh_token;
              if (typeof value.expires_in === "number" && Number.isFinite(value.expires_in)) tokens.expires_in = value.expires_in;
              sessionStorage.setItem(sessionTokenKey(serviceUrl), JSON.stringify(tokens));
              finish(undefined, tokens);
            } catch { finish(new Error("Could not save the sign-in session. Try again.")); }
          });
          dialog.addEventHandler(Office.EventType.DialogEventReceived, () => finish(new Error("Sign-in was closed before completion. Try again.")));
        });
    } catch { finish(new Error("Unable to open sign-in. Try again.")); }
  });
}
