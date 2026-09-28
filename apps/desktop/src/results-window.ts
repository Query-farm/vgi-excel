import { invoke } from "./bridge";
import type { ResultsSnapshot } from "./ResultsViewer";
import { receiveBrowserResults } from "./results-channel";

export async function openResultsWindow(snapshot: ResultsSnapshot): Promise<void> {
  await invoke("results.open", snapshot as unknown as Record<string, unknown>, 30_000);
}
export function receiveResults(onResult: (value: ResultsSnapshot) => void, onError: () => void): () => void {
  if (!window.chrome?.webview) return receiveBrowserResults(onResult, onError);
  let active = true;
  void invoke<ResultsSnapshot>("results.ready", {}, 20_000).then(value => { if (active) onResult(value); }, () => { if (active) onError(); });
  return () => { active = false; };
}
export function closeResults(): void { if (window.chrome?.webview) void invoke("results.close"); else window.close(); }
export async function maximizeResults(): Promise<void> {
  if (window.chrome?.webview) await invoke("results.maximize");
  else if (document.fullscreenElement) await document.exitFullscreen();
  else await document.documentElement.requestFullscreen();
}

export const nativeResultsWindow = typeof window !== "undefined" && !!window.chrome?.webview;
