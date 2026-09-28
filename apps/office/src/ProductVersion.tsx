import { useEffect, useRef } from "react";
import { browserRuntimeDiagnostics } from "./browser-backend";

export function AboutDialog({ onClose, onCopy }: { onClose(): void; onCopy(): void }): React.JSX.Element {
  const runtime = browserRuntimeDiagnostics();
  const diagnostics = `Cupola for Excel ${__APP_VERSION__}\nBuild ${__BUILD_ID__}\nTransport: HTTPS VGI only\nHost: Microsoft 365 Office Add-in\nEngine: Haybarn WebAssembly (${runtime.selectedBundle ?? "not started"})\nCross-origin isolated: ${runtime.crossOriginIsolated}\nSharedArrayBuffer: ${runtime.sharedArrayBuffer}\nEngine assets: ${runtime.assetBase}`;
  const dialog = useRef<HTMLElement>(null);
  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    dialog.current?.querySelector<HTMLButtonElement>("button")?.focus();
    return () => { previous?.focus(); };
  }, []);
  return <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}><section ref={dialog} onKeyDown={event => {
    if (event.key === "Escape") { event.preventDefault(); onClose(); }
    if (event.key !== "Tab") return;
    const controls = Array.from(dialog.current?.querySelectorAll<HTMLElement>('button:not(:disabled), a[href], summary') ?? []).filter(el => el.getClientRects().length > 0);
    const first = controls[0], last = controls[controls.length - 1];
    if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus(); }
    else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus(); }
  }} className="about-dialog" role="dialog" aria-modal="true" aria-labelledby="about-title"><button className="modal-close" aria-label="Close About" onClick={onClose}>×</button><img className="about-mark" src="./cupola-mark.svg" alt=""/><h2 id="about-title">Cupola for Excel</h2><p><a href="https://query.farm" target="_blank" rel="noopener noreferrer">Query.Farm</a></p><p className="about-version">Version {__APP_VERSION__}<br/>Build {__BUILD_ID__}</p><dl><div><dt>Transport</dt><dd>HTTPS VGI only</dd></div><div><dt>Data delivery</dt><dd>Excel table snapshots</dd></div><div><dt>Engine</dt><dd>Haybarn WebAssembly · {runtime.selectedBundle ?? "not started"}</dd></div><div><dt>OAuth worker bridge</dt><dd>{runtime.crossOriginIsolated && runtime.sharedArrayBuffer ? "Available" : "Unavailable in this host"}</dd></div><div><dt>Host</dt><dd>Microsoft 365</dd></div></dl><div className="actions"><button onClick={onCopy}>Copy diagnostics</button><button className="primary" onClick={onClose}>Done</button></div><textarea className="sr-only" readOnly value={diagnostics} aria-hidden="true"/></section></div>;
}

export function AboutContent({ diagnostics, onCopy }: { diagnostics: string; onCopy(): void }): React.JSX.Element {
  return <div className="about-content"><img className="about-mark" src="./cupola-mark.svg" alt=""/><h3>Cupola for Excel</h3><p className="about-version">Version {__APP_VERSION__}<br/>Build {__BUILD_ID__}</p><p><a href="https://query.farm" target="_blank" rel="noopener noreferrer">Query.Farm</a></p><details><summary>Diagnostics</summary><pre>{diagnostics}</pre></details><button onClick={onCopy}>Copy diagnostics</button></div>;
}
