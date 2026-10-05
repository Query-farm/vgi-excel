import { connectionSignIn } from "@query-farm/vgi-excel-core";
import { host, type DesktopConnection } from "./bridge";

export async function reauthenticateConnection(connection: DesktopConnection, update: (values: DesktopConnection[]) => void): Promise<void> {
  if (connection.members?.length) throw new Error("Choose an individual connection to sign in.");
  // The existing native sign-out invalidates this session and dependent profiles.
  // Keep the recovery marker if the user cancels or the provider rejects sign-in.
  connectionSignIn.require(connection.name);
  update(await host.signOut(connection));
  update(await host.signIn(connection));
  connectionSignIn.clear(connection.name);
}
