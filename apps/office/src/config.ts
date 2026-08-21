import type { ConnectionDefinition } from "@query-farm/vgi-excel-core";

const CONNECTIONS_KEY = "vgi.excel.connections.v1";
const DEFAULT_KEY = "vgi.excel.default-connection.v1";

export function loadConnections(): ConnectionDefinition[] {
  try {
    return JSON.parse(localStorage.getItem(CONNECTIONS_KEY) ?? "[]") as ConnectionDefinition[];
  } catch {
    return [];
  }
}

export function saveConnections(connections: ConnectionDefinition[]): void {
  localStorage.setItem(CONNECTIONS_KEY, JSON.stringify(connections));
  window.dispatchEvent(new Event("vgi-connections-changed"));
}

export function getDefaultConnectionName(): string | undefined {
  return localStorage.getItem(DEFAULT_KEY) ?? loadConnections()[0]?.name;
}

export function setDefaultConnectionName(name: string): void {
  localStorage.setItem(DEFAULT_KEY, name);
  window.dispatchEvent(new Event("vgi-connections-changed"));
}

export function sessionTokenKey(serviceOrigin: string): string {
  return `vgi.excel.oauth.${new URL(serviceOrigin).origin}`;
}

export function getServiceToken(location: string): OAuthTokens | null {
  try {
    const value = sessionStorage.getItem(sessionTokenKey(location));
    return value ? (JSON.parse(value) as OAuthTokens) : null;
  } catch {
    return null;
  }
}

export interface OAuthTokens {
  access_token: string;
  refresh_token?: string;
  expires_in?: number;
  id_token?: string;
}
