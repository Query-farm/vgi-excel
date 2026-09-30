# Cloudflare hosting for the Office add-in

Cupola's Microsoft 365 add-in runs in Excel's browser environment. npm builds its
static files; hosting does not require a Node.js application server.

## R2 asset bucket

The dedicated bucket created for Cupola for Excel is:

| Setting | Value |
| --- | --- |
| Bucket | `cupola-excel-assets` |
| Custom domain | `cupola-assets.query.farm` |
| Location hint | Eastern North America (`enam`) |
| Storage class | Standard |
| Minimum TLS | 1.2 |
| Public `r2.dev` URL | Disabled |

The custom domain provides public read access to uploaded objects. This bucket
is for distributable application assets, not workbook data or credentials. The
existing `cupola-assets` bucket and its domain are separate and remain unchanged.

`https://cupola-assets.query.farm/health.json` is a small public readiness probe.
Application files and Haybarn assets use immutable `releases/{version}-{build}/` paths.

## Browser access

The bucket has this Wrangler CORS configuration. It permits reads of public
assets from any origin and does not grant browser upload or deletion access.

```json
{
  "rules": [
    {
      "allowed": {
        "origins": ["*"],
        "methods": ["GET", "HEAD"],
        "headers": ["Range"]
      },
      "exposeHeaders": ["ETag", "Content-Length", "Content-Range"],
      "maxAgeSeconds": 3600
    }
  ]
}
```

Save that JSON to a file and apply it with:

```sh
wrangler r2 bucket cors set cupola-excel-assets --file cors.json
```

Inspect the live settings with:

```sh
wrangler r2 bucket domain list cupola-excel-assets
wrangler r2 bucket dev-url get cupola-excel-assets
wrangler r2 bucket cors list cupola-excel-assets
curl -i -H 'Origin: https://query.farm' https://cupola-assets.query.farm/health.json
```

Cloudflare returns CORS response headers when the request includes `Origin`.
Bucket configuration is managed through authenticated Cloudflare tooling;
no Cloudflare credentials belong in the browser bundle or repository.

## Application host

`https://cupola.query.farm` serves the Office application through the
`cupola-excel-office` Worker in `cloudflare/office-worker.mjs`. The Worker reads
static objects from R2 and adds COOP/COEP headers. There is no Node.js application
server or customer-data API. The large Haybarn WASM files remain in R2 rather
than Cloudflare Pages, whose per-file limit is 25 MiB.

- Installation manifest: <https://cupola.query.farm/manifest.xml>
- Deployment status: <https://cupola.query.farm/health.json>
- OAuth SPA redirect URI: `https://cupola.query.farm/oauth-dialog.html`

The manifest pins HTML, icons, and function metadata to one release. HTML in turn
pins its JavaScript, workers, and WASM. The stable manifest and OAuth callback
are revalidated; versioned files are cached immutably. Old release objects stay
available for existing manifests and open panes. Public function metadata allows
cross-origin GET/HEAD/OPTIONS without credentials. Office SDK script tags use
anonymous CORS to load under the isolation headers.

## Deploying

Use Node.js 22.12 or newer and the locked Wrangler dependency:

```sh
npm ci
npx wrangler login
npm run check
npm test
npm run test:hosting
npm run build
npm run test:ui
npm run test:office-wasm
npm run test:office-dev
npm run deploy:office
CUPOLA_OFFICE_BASE_URL=https://cupola.query.farm npm run test:office-hosted
```

`deploy:office` packages the production manifest, uploads files with their MIME
types, verifies every uploaded SHA-256 digest, and then activates the Worker.
`-- --prepare-only` builds the package and inventory without publishing.
A build identifier cannot be reused with different file bytes. Increment
`cupolaBuild` and the matching native declaration for each installable update.
Increment `officeManifestVersion` when updating the Office manifest. Microsoft
requires a manifest version of at least 1.0, so this four-part package version is
tracked separately from Cupola’s product version and checked against `package.json`.

The GitHub Actions workflow **Cupola for Excel Office deployment** is manually
triggered on `main` and uses the `cupola-office` environment. Configure
`CLOUDFLARE_API_TOKEN` as a repository or environment secret. Scope it to the
Cloudflare account and `query.farm` zone with Workers deployment, R2 Storage Edit,
and Workers Routes Edit permissions. The initial local deployment can use
Wrangler OAuth; CI requires its own API token. Never put that token into the
browser bundle, repository, or a `VITE_` variable.

The workflow runs validation before publishing, tests the hosted application
against a live HTTPS VGI catalog afterward, and saves the manifest and deployment
inventory as an artifact. Hosted tests suppress Sentry event delivery. These
browser checks do not replace real Excel qualification.

For rollback, select a previously verified release inventory and activate its
prefix without rebuilding or overwriting its objects:

```sh
npx wrangler deploy --config cloudflare/wrangler.jsonc --var RELEASE:VERSION-BUILD
```

Then verify `/health.json` and rerun the hosted tests with
`CUPOLA_OFFICE_EXPECTED_RELEASE=VERSION-BUILD`. Existing installed manifests
remain pinned to their release; distribute the selected release's manifest again
when rolling back an installed Office add-in.

## Installing in Microsoft 365

Download the production manifest and sideload it into Excel for a pilot, or have
a Microsoft 365 administrator upload it through **Settings → Integrated apps →
Upload custom apps** and select the intended users. Hosting alone does not install
the add-in in anyone's Excel and does not create a Microsoft Marketplace listing.

VGI services must permit the application origin through CORS. For protected
services, register the stable OAuth redirect URI above with the provider used by
that service. Cupola cannot register it for an arbitrary customer's provider.
Complete [Office release qualification](office-release-qualification.md) in a
signed-in Excel session, including real formulas, workbook insertion, and a
protected-service login, before broad distribution.

References: [R2 custom domains](https://developers.cloudflare.com/r2/buckets/public-buckets/),
[R2 CORS](https://developers.cloudflare.com/r2/buckets/cors/),
[Pages limits](https://developers.cloudflare.com/pages/platform/limits/),
[Office add-in distribution](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish).
