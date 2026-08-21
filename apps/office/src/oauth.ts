import { sessionTokenKey, type OAuthTokens } from "./config";

export async function signIn(serviceUrl: string): Promise<OAuthTokens> {
  const dialogUrl = new URL("/oauth-dialog.html", window.location.origin);
  dialogUrl.searchParams.set("service", serviceUrl);
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      dialogUrl.toString(),
      { height: 65, width: 40, displayInIframe: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
          reject(new Error(result.error?.message || "Unable to open the OAuth dialog."));
          return;
        }
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (event) => {
          try {
            const payload = JSON.parse((event as { message: string }).message) as
              | { ok: true; tokens: OAuthTokens }
              | { ok: false; error: string };
            dialog.close();
            if (!payload.ok) reject(new Error(payload.error));
            else {
              sessionStorage.setItem(sessionTokenKey(serviceUrl), JSON.stringify(payload.tokens));
              resolve(payload.tokens);
            }
          } catch (error) {
            dialog.close();
            reject(error);
          }
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (event) => {
          reject(new Error(`OAuth dialog closed (${(event as { error: number }).error}).`));
        });
      },
    );
  });
}
