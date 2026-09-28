import type { ResultsSnapshot } from "./ResultsViewer";
import { openBrowserResults, receiveBrowserResults } from "./results-channel";

let closeOfficeResults: (() => void) | undefined;
export function openResultsWindow(snapshot: ResultsSnapshot): Promise<void> {
  const root = document.documentElement;
  snapshot = { ...snapshot, theme: { bodyBackgroundColor: root.style.getPropertyValue("--office-background"), bodyForegroundColor: root.style.getPropertyValue("--office-foreground"), isDarkTheme: root.dataset.officeTheme === "dark" } };
  if (typeof Office === "undefined" || !Office.context?.host) return openBrowserResults(snapshot);
  if (!Office.context.requirements.isSetSupported("DialogApi", "1.2")) return Promise.reject(new Error("This Excel version cannot open a results window. Update Microsoft 365 and try again."));
  const url = new URL("./results.html?mode=office", location.href);
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(url.href, { width: 90, height: 90, displayInIframe: false }, response => {
      if (response.status !== Office.AsyncResultStatus.Succeeded) { reject(new Error("Could not open the results window. Close any other Cupola dialog, allow popups if prompted, and try again.")); return; }
      const dialog = response.value;
      let loaded = false;
      const timeout = window.setTimeout(() => { dialog.close(); reject(new Error("The results window did not respond. Try opening it again.")); }, 20_000);
      // Pull fixed-size chunks to avoid a large result exceeding a host's message limit.
      const data = JSON.stringify(snapshot), size = 16_000, count = Math.ceil(data.length / size);
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, event => {
        const message = event as { message: string; origin?: string };
        if (message.origin && message.origin !== location.origin) return;
        let request: { type?: string; index?: number };
        try { request = JSON.parse(message.message); } catch { return; }
        if (request.type === "chunk" && Number.isInteger(request.index) && request.index! >= 0 && request.index! < count) {
          dialog.messageChild(JSON.stringify({ type: "chunk", index: request.index, count, text: data.slice(request.index! * size, (request.index! + 1) * size) }), { targetOrigin: location.origin });
        }
        if (request.type === "received") { loaded = true; clearTimeout(timeout); resolve(); }
        if (request.type === "close") { clearTimeout(timeout); dialog.close(); if (!loaded) reject(new Error("Results window closed before loading. Try opening it again.")); }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => { clearTimeout(timeout); if (!loaded) reject(new Error("Results window closed before loading. Try opening it again.")); });
    });
  });
}
export function receiveResults(onResult: (value: ResultsSnapshot) => void, onError: () => void): () => void {
  if (new URLSearchParams(location.search).get("mode") !== "office") return receiveBrowserResults(onResult, onError);
  let active = true;
  const script = document.createElement("script");
  script.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
  const timeout = window.setTimeout(onError, 20_000);
  script.onerror = () => { clearTimeout(timeout); if (active) onError(); };
  script.onload = () => { void Office.onReady(() => {
    if (!active) return;
    const send = (value: unknown) => Office.context.ui.messageParent(JSON.stringify(value), { targetOrigin: location.origin });
    closeOfficeResults = () => send({ type: "close" });
    const chunks: string[] = [];
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, event => {
      if (event.origin && event.origin !== location.origin) return;
      try {
        const message = JSON.parse(event.message);
        if (message.type !== "chunk" || message.index !== chunks.length || typeof message.text !== "string") return;
        chunks.push(message.text);
        if (chunks.length === message.count) { clearTimeout(timeout); onResult(JSON.parse(chunks.join(""))); chunks.length = 0; send({ type: "received" }); }
        else send({ type: "chunk", index: chunks.length });
      } catch { clearTimeout(timeout); onError(); }
    }, result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) { clearTimeout(timeout); onError(); return; }
      send({ type: "chunk", index: 0 });
    });
  }); };
  document.head.append(script);
  return () => { active = false; clearTimeout(timeout); script.remove(); closeOfficeResults = undefined; };
}
export function closeResults(): void { if (closeOfficeResults) closeOfficeResults(); else window.close(); }
export async function maximizeResults(): Promise<void> {
  if (document.fullscreenElement) await document.exitFullscreen();
  else await document.documentElement.requestFullscreen();
}

export const nativeResultsWindow = false;
