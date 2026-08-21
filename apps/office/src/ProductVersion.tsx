import { browserRuntimeDiagnostics } from "./browser-backend";

export function ProductVersion({ onAbout }: { onAbout?(): void }): React.JSX.Element {
  return <footer className="product-footer"><span>Query.Farm</span>{onAbout ? <button className="version-link" onClick={onAbout} aria-label={`About Cupola for Excel version ${__APP_VERSION__}`}>v{__APP_VERSION__}</button> : <span>v{__APP_VERSION__}</span>}</footer>;
}

export function AboutDialog({ onClose, onCopy }: { onClose(): void; onCopy(): void }): React.JSX.Element {
  const runtime = browserRuntimeDiagnostics();
  const diagnostics = `Cupola for Excel ${__APP_VERSION__}\nBuild ${__BUILD_ID__}\nTransport: HTTPS VGI only\nHost: Microsoft 365 Office Add-in\nEngine: Haybarn WebAssembly (${runtime.selectedBundle ?? "not started"})\nCross-origin isolated: ${runtime.crossOriginIsolated}\nSharedArrayBuffer: ${runtime.sharedArrayBuffer}\nEngine assets: ${runtime.assetBase}`;
  return <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}><section className="about-dialog" role="dialog" aria-modal="true" aria-labelledby="about-title"><button className="modal-close" aria-label="Close About" onClick={onClose}>×</button><img className="about-mark" src="./cupola-mark.svg" alt=""/><h2 id="about-title">Cupola for Excel</h2><p className="about-version">Version {__APP_VERSION__}<br/>Build {__BUILD_ID__}</p><dl><div><dt>Transport</dt><dd>HTTPS VGI only</dd></div><div><dt>Data delivery</dt><dd>Excel table snapshots</dd></div><div><dt>Engine</dt><dd>Haybarn WebAssembly · {runtime.selectedBundle ?? "not started"}</dd></div><div><dt>OAuth worker bridge</dt><dd>{runtime.crossOriginIsolated && runtime.sharedArrayBuffer ? "Available" : "Unavailable in this host"}</dd></div><div><dt>Host</dt><dd>Microsoft 365</dd></div></dl><div className="actions"><button onClick={onCopy}>Copy diagnostics</button><button className="primary" onClick={onClose}>Done</button></div><textarea className="sr-only" readOnly value={diagnostics} aria-hidden="true"/></section></div>;
}
