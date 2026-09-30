# Office release qualification

This covers the Microsoft 365 Office add-in, separately from the Windows XLL.
Passing browser tests does not establish that every Excel host is qualified.

## Hosted deployment evidence (2026-09-30)

Cupola 0.5.0 build 20260930.3 is hosted at `https://cupola.query.farm`,
with installation manifest version 1.0.0.1 at `/manifest.xml`.

- Every uploaded application asset was verified against its local SHA-256 digest
  before the Cloudflare Worker activated the release.
- Microsoft's manifest validator accepted the production manifest, including its
  XML schema, HTTPS source locations, icons, and version.
- All eight hosted browser checks passed: real Microsoft Office SDK loading,
  cross-origin isolation, metadata CORS, connection discovery/error recovery,
  session-credential reuse with a mocked protected service, and live Haybarn VGI
  queries with cancellation and result windows. No real protected-service login
  or Excel workbook writes are established by these browser tests.
- Automated deployment is defined in `.github/workflows/office-deploy.yml`.
  Its first GitHub run still requires the `CLOUDFLARE_API_TOKEN` secret.

The live Excel checks below remain required for this production origin before
broad distribution. Hosting does not install the add-in into a Microsoft 365
organization or publish it to Microsoft Marketplace.

## Current evidence (2026-09-28)

- Real Safari / Excel for the web: the user has loaded Cupola and exercised Ask AI.
  The development certificate, metadata CORS, missing engine assets, and unsupported
  browser confirmation calls found during that session have been corrected.
- WebKit: Office light/dark themes override the OS preference; confirmation dialogs
  remain readable; result windows inherit the parent theme. Missing host theme
  information defaults to light. Theme refresh does not reload the UI.
- WebKit: insertion requires in-app confirmation; Cancel and Escape make no writes.
  The workbook boundary also requires separate approval before overwriting cells.
- WebKit: a fresh pane retains connection definitions and query drafts, while OAuth
  tokens remain scoped to the original browser session.
- Live WebKit/Haybarn: the worksheet-function implementation returned a QUERY spill
  containing 42, VALUE returned 7, and CALL returned Clear sky from Open Meteo.
  This invokes the function code, not Excel's formula calculation engine.
- Deterministic tests: Office dialog opening failure, user closure, malformed and
  foreign messages, success, late messages, and cancellation before/after opening.
  Successful login stores only the session tokens needed by Cupola. No real login
  or production telemetry is involved in these tests.
- Deterministic tests: custom-function registration, spill conversion, cancellation,
  subsequent queries, and Excel error conversion. Existing agent cancellation tests
  cover inventory, SQL, and catalog work and a follow-up interaction.

## Remaining live Excel checks

Use a disposable workbook. These checks require a signed-in Excel session and,
for protected data, an authorized user completing the identity provider's login.
Do not paste tokens, passwords, or API keys into reports.

1. Reload Excel and verify that Cupola matches Excel's light/dark appearance.
2. In Settings → VGI connections, enter an authorized protected endpoint, choose
   Find catalogs, and complete sign-in. Confirm discovery and a read-only query.
   Close an unfinished sign-in once and confirm that a retry succeeds.
3. Save an anonymous Open Meteo connection named Weather, using
   `https://vgi-open-meteo.rusty-bb6.workers.dev` and catalog `open_meteo`.
4. Enter `=VGI.QUERY("SELECT 42 AS answer","Weather")` and verify the answer header
   and 42 spill. Enter `=VGI.VALUE("SELECT 7","Weather")` and verify 7.
   With Weather selected as default, enter
   `=VGI.CALL("open_meteo.main.weather_code_text",0)` and verify Clear sky.
5. In Ask AI, request a small result. Cancel insertion once, then approve it.
   Verify one static table is added. Test an occupied destination and decline
   replacement; existing data must remain unchanged.
6. Stop an AI response with the Stop button and another with Escape. Verify that
   no additional table is created and a follow-up interaction succeeds.
7. Save the workbook, close its browser tab, reopen it, and open Cupola. Verify
   inserted data remains, formulas calculate, and saved connections/drafts appear.
   A fresh Office session may require signing in again; credentials are not stored
   in the workbook. Office snapshots do not become refreshable Power Query tables.

Record host/browser versions and the outcome of each step. Mark protected sign-in
and actual workbook save/reopen as pending until these checks are observed.

## Repeatable checks

```sh
npm run check
npm test
npm run build
npm run test:office-dev
npm run test:ui
npm run test:office-wasm
```

The last command includes live HTTPS VGI integration. It uses a standalone browser
preview; it does not replace the live Excel checks above. A production-hosted
candidate must repeat the host checks with the production OAuth redirect URI.
