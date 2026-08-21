import { authorizationUrl, oidcMetadataUrl, protectedResourceMetadataUrl } from "./oauth-core";

interface ResourceMetadata {
  authorization_servers: string[];
  scopes_supported?: string[];
  client_id?: string;
  token_endpoint?: string;
}

interface OidcMetadata {
  authorization_endpoint: string;
  token_endpoint: string;
}

interface Pending {
  service: string;
  verifier: string;
  tokenEndpoint: string;
  clientId: string;
  redirectUri: string;
  scope: string;
  state: string;
}

const PENDING_KEY = "vgi.excel.oauth.pending";
const statusElement = document.querySelector<HTMLParagraphElement>("#status")!;

Office.onReady(() => void run().catch((error) => finish({ ok: false, error: message(error) })));

async function run(): Promise<void> {
  const url = new URL(window.location.href);
  const code = url.searchParams.get("code");
  const returnedState = url.searchParams.get("state");
  if (code) {
    const pending = JSON.parse(sessionStorage.getItem(PENDING_KEY) ?? "null") as Pending | null;
    if (!pending || pending.state !== returnedState) throw new Error("OAuth state validation failed.");
    statusElement.textContent = "Completing sign-in…";
    const body = new URLSearchParams({
      grant_type: "authorization_code",
      code,
      redirect_uri: pending.redirectUri,
      client_id: pending.clientId,
      code_verifier: pending.verifier,
    });
    const response = await fetch(pending.tokenEndpoint, {
      method: "POST",
      headers: { "content-type": "application/x-www-form-urlencoded" },
      body,
    });
    if (!response.ok) throw new Error(`Token exchange failed (${response.status}): ${await response.text()}`);
    sessionStorage.removeItem(PENDING_KEY);
    finish({ ok: true, tokens: await response.json() });
    return;
  }

  const service = url.searchParams.get("service");
  if (!service) throw new Error("No VGI service was supplied.");
  statusElement.textContent = "Discovering the identity provider…";
  const resourceUrl = protectedResourceMetadataUrl(service);
  const resource = await fetchJson<ResourceMetadata>(resourceUrl);
  const issuer = resource.authorization_servers?.[0];
  if (!issuer) throw new Error("The VGI service did not advertise an OAuth authorization server.");
  const oidc = await fetchJson<OidcMetadata>(oidcMetadataUrl(issuer));
  const clientId = resource.client_id || "vgi-excel";
  const scope = resource.scopes_supported?.join(" ") || "openid profile offline_access";
  const redirectUri = `${window.location.origin}/oauth-dialog.html`;
  const verifier = randomBase64Url(64);
  const challenge = base64Url(new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier))));
  const state = randomBase64Url(24);
  const pending: Pending = {
    service,
    verifier,
    tokenEndpoint: resource.token_endpoint || oidc.token_endpoint,
    clientId,
    redirectUri,
    scope,
    state,
  };
  sessionStorage.setItem(PENDING_KEY, JSON.stringify(pending));
  window.location.replace(authorizationUrl(oidc.authorization_endpoint, { clientId, redirectUri, scope, state, challenge }));
}

async function fetchJson<T>(url: string): Promise<T> {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`OAuth discovery failed (${response.status}).`);
  return response.json() as Promise<T>;
}

function finish(payload: unknown): void {
  Office.context.ui.messageParent(JSON.stringify(payload));
}

function randomBase64Url(bytes: number): string {
  const value = crypto.getRandomValues(new Uint8Array(bytes));
  return base64Url(value);
}

function base64Url(value: Uint8Array): string {
  let binary = "";
  for (const byte of value) binary += String.fromCharCode(byte);
  return btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replace(/=+$/, "");
}

function message(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
