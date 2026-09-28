import type { ResultsSnapshot } from "./ResultsViewer";

// Result values stay in memory. Only a random channel identifier enters the URL.
export function openBrowserResults(snapshot: ResultsSnapshot): Promise<void> {
  const token = crypto.randomUUID();
  const channel = new BroadcastChannel(`cupola-results-${token}`);
  const url = new URL("./results.html", location.href);
  url.hash = token;
  const popup = window.open(url, "_blank", "popup,width=1100,height=800,resizable=yes,scrollbars=yes");
  if (!popup) { channel.close(); return Promise.reject(new Error("Allow popups for Cupola, then try opening the results window again.")); }
  return new Promise((resolve, reject) => {
    const cleanup = () => { clearTimeout(timeout); channel.close(); };
    const timeout = window.setTimeout(() => { cleanup(); reject(new Error("The results window did not respond. Close it and try again.")); }, 20_000);
    channel.onmessage = event => {
      if (event.data?.type === "ready") channel.postMessage({ type: "snapshot", snapshot });
      if (event.data?.type === "received") { cleanup(); resolve(); }
    };
  });
}
export function receiveBrowserResults(onResult: (value: ResultsSnapshot) => void, onError: () => void): () => void {
  const token = location.hash.slice(1);
  if (!/^[a-f0-9-]{36}$/.test(token)) { onError(); return () => {}; }
  window.opener = null;
  const channel = new BroadcastChannel(`cupola-results-${token}`);
  const cleanup = () => { clearInterval(retry); clearTimeout(timeout); channel.close(); };
  const timeout = window.setTimeout(() => { cleanup(); onError(); }, 20_000);
  const retry = window.setInterval(() => channel.postMessage({ type: "ready" }), 250);
  channel.onmessage = event => {
    if (event.data?.type !== "snapshot") return;
    onResult(event.data.snapshot);
    channel.postMessage({ type: "received" });
    cleanup();
  };
  channel.postMessage({ type: "ready" });
  return cleanup;
}
