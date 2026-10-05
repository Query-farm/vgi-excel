import { useState, useSyncExternalStore } from "react";
import { connectionSignIn } from "@query-farm/vgi-excel-core";

interface Connection { name: string; members?: string[] }

/** Recovery is explicit: never replay a query or a workbook write here. */
export function AuthRecovery({ connections, busy, onSignIn, onOpenConnections }: {
  connections: readonly Connection[]; busy: boolean;
  onSignIn(name: string): Promise<void>; onOpenConnections(): void;
}): React.JSX.Element | null {
  const required = useSyncExternalStore(connectionSignIn.subscribe, connectionSignIn.snapshot);
  const [pending, setPending] = useState(false);
  const [failed, setFailed] = useState<string | null>(null);
  const name = required.find(name => connections.some(value => value.name === name));
  const connection = connections.find(value => value.name === name);
  if (!connection) return null;
  const profile = !!connection.members?.length;
  async function signIn(): Promise<void> {
    setPending(true); setFailed(null);
    try { await onSignIn(connection!.name); }
    catch { setFailed(connection!.name); }
    finally { setPending(false); }
  }
  return <div className="notice error" role="alert">
    <span className="notice-icon" aria-hidden="true">!</span>
    <span>{profile ? "A connection needs you to sign in again." : `Your connection “${connection.name}” needs you to sign in again.`}
      {profile && " Open Connections, choose the connection that needs attention, then select Sign in again."}
      {failed === connection.name && " Sign-in wasn’t completed. Try again when you’re ready."}
      {pending && " Waiting for sign-in…"}
    </span>
    <div className="notice-actions">
      <button disabled={busy || pending} onClick={profile ? onOpenConnections : () => void signIn()}>{profile ? "Open connections" : pending ? "Signing in…" : "Sign in again"}</button>
      <button disabled={pending} aria-label="Dismiss sign-in message" onClick={() => connectionSignIn.clear(connection.name)}>×</button>
    </div>
  </div>;
}
