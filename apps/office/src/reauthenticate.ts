import { connectionSignIn, type ConnectionDefinition } from "@query-farm/vgi-excel-core";
import { signIn } from "./oauth";
import { resetRuntime } from "./runtime";

export async function reauthenticateConnection(connection: ConnectionDefinition): Promise<void> {
  // Preserve the old session on cancellation. Only a completed dialog replaces it.
  connectionSignIn.require(connection.name);
  await signIn(connection.location);
  resetRuntime();
  connectionSignIn.clear(connection.name);
}
