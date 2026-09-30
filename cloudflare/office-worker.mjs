const releasePath = /^\/releases\/[0-9]+\.[0-9]+\.[0-9]+-[0-9]{8}\.[0-9]+\//;
const aliases = new Set(['taskpane.html', 'results.html', 'functions.json', 'manifest.xml', 'cupola-mark.svg']);

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const metadata = url.pathname.endsWith('/functions.json');
    const headers = new Headers({
      'Cross-Origin-Opener-Policy': 'same-origin',
      'Cross-Origin-Embedder-Policy': 'require-corp',
      'Cross-Origin-Resource-Policy': 'cross-origin',
      'X-Content-Type-Options': 'nosniff',
      'Referrer-Policy': 'strict-origin-when-cross-origin',
    });
    if (metadata) {
      headers.set('Access-Control-Allow-Origin', '*');
      headers.set('Access-Control-Allow-Methods', 'GET, HEAD, OPTIONS');
    }
    const respond = (body, status = 200) => new Response(body, { status, headers });
    if (request.method === 'OPTIONS' && metadata) return respond(null, 204);
    if (!['GET', 'HEAD'].includes(request.method)) {
      headers.set('Allow', metadata ? 'GET, HEAD, OPTIONS' : 'GET, HEAD');
      return respond('Method not allowed', 405);
    }
    headers.set('Cache-Control', 'no-store');
    if (url.pathname === '/health.json') {
      headers.set('Content-Type', 'application/json');
      return respond(request.method === 'HEAD' ? null : JSON.stringify({ product: 'Cupola for Excel', release: env.RELEASE }));
    }
    if (url.pathname === '/' || aliases.has(url.pathname.slice(1)) && url.pathname !== '/manifest.xml' && url.pathname !== '/functions.json') {
      headers.set('Location', `/releases/${env.RELEASE}/${url.pathname === '/' ? 'taskpane.html' : url.pathname.slice(1)}${url.search}`);
      return respond(null, 302);
    }
    // OAuth has a stable redirect URI registered with the data service's provider.
    const stable = ['/manifest.xml', '/functions.json', '/oauth-dialog.html'].includes(url.pathname);
    const key = stable ? `releases/${env.RELEASE}${url.pathname}` : url.pathname.slice(1);
    if ((!stable && !releasePath.test(url.pathname)) || /[%\\]|(?:^|\/)\./.test(key) || key.endsWith('.map') || !/\.(html|js|css|json|xml|svg|png|wasm)$/.test(key)) {
      return respond('Not found', 404);
    }
    try {
      const object = await env.OFFICE_ASSETS.get(key);
      if (!object) return respond('Not found', 404);
      object.writeHttpMetadata(headers);
      headers.set('ETag', object.httpEtag);
      headers.set('Cache-Control', stable ? 'no-cache' : 'public, max-age=31536000, immutable');
      if (request.headers.get('If-None-Match') === object.httpEtag) return respond(null, 304);
      headers.set('Content-Length', String(object.size));
      return respond(request.method === 'HEAD' ? null : object.body);
    } catch {
      // Requests may contain OAuth callback parameters. Never log request URLs.
      headers.set('Cache-Control', 'no-store');
      return respond('Asset temporarily unavailable', 503);
    }
  },
};
