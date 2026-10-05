import { useEffect, useState } from "react";
import { invoke } from "./bridge";

type UpdateStatus = { supported: boolean; daily: boolean; phase: string; message: string; lastChecked?: string; version?: string; build?: string; notes?: string };
export function Updates(): React.JSX.Element {
  const [state, setState] = useState<UpdateStatus | null>(null);
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState("");
  const [error, setError] = useState("");
  useEffect(() => {
    let active = true;
    void invoke<UpdateStatus>("updates.status", { autoCheck: true }).then(value => { if (active && typeof value?.phase === "string") setState(value); }).catch(() => { if (active) setError("Could not read update status. Reopen Cupola and try again."); });
    return () => { active = false; };
  }, []);
  async function act(method: string, params: Record<string, unknown> = {}): Promise<void> {
    setBusy(true); setError(""); setProgress(method === "updates.check" ? "Checking for updates…" : method === "updates.download" ? "Downloading and verifying the installer…" : "");
    try { setState(await invoke<UpdateStatus>(method, params, 660_000, setProgress)); }
    catch { setError("The update action could not finish. Reopen Cupola and try again."); }
    finally { setBusy(false); setProgress(""); }
  }
  const notes = state?.notes?.startsWith("https://github.com/Query-farm/vgi-excel/releases/tag/v") ? state.notes : undefined;
  return <section className="cupola-updates" aria-labelledby="updates-title"><h4 id="updates-title">Updates</h4>
    <p role="status" aria-live="polite">{error || progress || state?.message || "Reading update status…"}</p>
    {state?.version && <p>Available: {state.version}<br/>Build {state.build}</p>}
    {state?.lastChecked && <p className="muted">Last checked {new Date(state.lastChecked).toLocaleString()}</p>}
    {state?.supported && <>
      <label className="update-preference"><input type="checkbox" checked={state.daily} disabled={busy} onChange={event => void act("updates.preference", { daily: event.target.checked })}/> Check for updates daily</label>
      <p className="muted">Checks run when you open Cupola. Updates install only when you choose.</p>
      {state.phase === "ready" && <p>Save your work, then close all Excel windows. The updater will wait until Excel is closed. Windows may ask for administrator approval.</p>}
      <div className="actions">
        {state.phase === "available" ? <button className="primary" disabled={busy} onClick={() => void act("updates.download")}>Download update</button>
          : state.phase === "ready" ? <button className="primary" disabled={busy} onClick={() => void act("updates.install")}>Install update</button>
          : <button disabled={busy} onClick={() => void act("updates.check")}>{state.phase === "error" ? "Retry update check" : "Check for updates"}</button>}
        {notes && <a href={notes} target="_blank" rel="noopener noreferrer">Release notes</a>}
      </div>
    </>}
  </section>;
}
