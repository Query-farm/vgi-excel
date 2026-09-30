import { useEffect, useState } from "react";
import type { DesktopConnection } from "./bridge";
import { sameWorkspace, workspaceMembers, workspaceSelectionError } from "./workspace";

export function WorkspaceConnections({ values, busy, apply }: { values: DesktopConnection[]; busy: boolean; apply(members: string[]): Promise<void> }): React.JSX.Element | null {
  const applied = workspaceMembers(values);
  const appliedKey = JSON.stringify(applied);
  const [members, setMembers] = useState(applied);
  const [error, setError] = useState("");
  useEffect(() => { setMembers(JSON.parse(appliedKey)); setError(""); }, [appliedKey]);
  const singles = values.filter(value => !value.members?.length);
  const profiles = values.filter(value => value.members?.length && !value.isWorkspaceProfile);
  const dirty = !sameWorkspace(members, applied);
  const validation = workspaceSelectionError(members, values);
  if (!singles.length) return null;
  return <section className="workspace-catalogs" aria-label="Workspace catalogs">
    <h3>Workspace catalogs</h3>
    <p>Use these catalogs together in Query Editor, Ask AI, and Catalog View.</p>
    <fieldset className="combined-connections" disabled={busy}><legend className="sr-only">Included connections</legend>
      {singles.map(value => <label className="combined-connection-choice" key={value.name}><input type="checkbox" checked={members.includes(value.name)} onChange={event => { setMembers(event.target.checked ? [...members, value.name] : members.filter(name => name !== value.name)); setError(""); }}/><span><strong>{value.name}</strong>{value.catalog !== value.name && <small>{value.catalog}</small>}</span></label>)}
      {members.length > 0 && <label>Default catalog<select value={members[0]} onChange={event => { setMembers([event.target.value, ...members.filter(name => name !== event.target.value)]); setError(""); }}>{members.map(name => <option key={name} value={name}>{values.find(value => value.name === name)?.catalog ?? name}</option>)}</select><small>Used when SQL doesn’t include a catalog name.</small></label>}
      {validation && <p className="field-error" role="alert">{validation}</p>}
      <button className="primary" disabled={!dirty || !!validation} onClick={() => { setError(""); void apply(members).catch(error => setError(error instanceof Error ? error.message : "Could not update the workspace.")); }}>Apply changes</button>
    </fieldset>
    {error ? <p className="field-error" role="alert">{error}</p> : <p role="status">{dirty ? "Changes not applied" : "Workspace is up to date"}</p>}
    {!!profiles.length && <details><summary>Saved profiles (advanced)</summary>{profiles.map(value => <div className="workspace-profile" key={value.name}><strong>{value.name}</strong><button disabled={busy} onClick={() => { setMembers([...value.members!]); setError(""); }}>Load into workspace</button></div>)}</details>}
  </section>;
}
