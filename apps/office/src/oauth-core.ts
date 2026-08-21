export function protectedResourceMetadataUrl(serviceUrl: string): string {
  const service = httpsUrl(serviceUrl, "VGI service");
  return new URL("/.well-known/oauth-protected-resource", service.origin).toString();
}

export function oidcMetadataUrl(issuerUrl: string): string {
  const issuer = httpsUrl(issuerUrl, "OAuth issuer");
  const base = issuer.toString().endsWith("/") ? issuer : new URL(`${issuer.toString()}/`);
  return new URL(".well-known/openid-configuration", base).toString();
}

export function authorizationUrl(
  endpoint: string,
  values: { clientId: string; redirectUri: string; scope: string; state: string; challenge: string },
): string {
  const authorize = httpsUrl(endpoint, "OAuth authorization endpoint");
  authorize.search = new URLSearchParams({
    response_type: "code",
    client_id: values.clientId,
    redirect_uri: values.redirectUri,
    scope: values.scope,
    state: values.state,
    code_challenge: values.challenge,
    code_challenge_method: "S256",
  }).toString();
  return authorize.toString();
}

function httpsUrl(value: string, label: string): URL {
  let url: URL;
  try { url = new URL(value); } catch { throw new Error(`${label} must be a valid HTTPS URL.`); }
  if (url.protocol !== "https:") throw new Error(`${label} must use HTTPS.`);
  return url;
}
