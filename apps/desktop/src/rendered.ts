import { invoke } from "./bridge";

// Notify the native cover only after React has committed and the browser has painted.
export function notifyRendered(): () => void {
  let frame = requestAnimationFrame(() => {
    frame = requestAnimationFrame(() => {
      if (window.chrome?.webview) void invoke("ui.rendered", {}, 15_000).catch(() => {});
    });
  });
  return () => cancelAnimationFrame(frame);
}
