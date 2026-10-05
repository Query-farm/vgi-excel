# Cupola for Excel tests

The repository has three complementary suites.

## Cross-platform tests

Run TypeScript checks, pure connection/SQL policy tests, Arrow-to-Excel temporal
and numeric boundary conversion, mocked OAuth URL policy, and a mocked
multi-round agent tool call:

```sh
npm run check
npm test
npm run build
npm run test:ui
```

These tests require Node.js 22.12 or newer and no Excel installation, VGI
credentials, model API key, or trusted development certificate. Unit tests never
install certificates. The live WASM preview generates its own untrusted HTTPS
certificate in `dev-certs/preview`; Playwright accepts it for that test only.
The Playwright suite renders both task panes at 300, 320, 340, 400, 720, and
1060 pixels and checks onboarding, connection settings, footer visibility, and
horizontal containment. It also creates multiple Ask AI conversation tabs,
reloads both the desktop Workbench and Microsoft 365 task pane, and verifies
that the selected tab, visible transcript, model history, and connection
scoping survive without persisting API keys or process-local query result IDs.
AI regression tests cover signed thinking streams, workspace headers, stable
prompt-cache prefixes, metadata filters and examples, narrow settings panels,
and Markdown-table clipboard output. Anthropic responses are mocked.
Accounting numeric tests keep complete DECIMAL columns numeric when every value
fits Excel's 15-significant-digit limit, retain unsafe precision as text, apply
scale-aware number formats, and preserve leading-zero account codes.

## Cloudflare hosting checks

`npm run test:hosting` checks the Worker’s routing, isolation headers, metadata
CORS, OAuth callback handling, and exclusion of private files. Run
`CUPOLA_OFFICE_BASE_URL=https://cupola.query.farm npm run test:office-hosted`
after deployment to verify the public manifest and assets, load the real Office
SDK, and exercise the WASM integration against the hosted build. Hosted tests
block Sentry delivery. `npm run test:office-dev` runs the WebKit development tests
sequentially with isolated, empty Vite caches. The HTTP-only fixture disables
background dependency scanning, preloading, and file watching so it shuts down
cleanly on Linux. Browser fixtures exercise normal dependency optimization.

## Live Microsoft 365 / Haybarn-WASM test

```sh
npm run test:office-wasm
```

Unlike the fast UI tests, this uses the production Office bundle and a real
network connection. It verifies that:

1. Haybarn WASM and its worker are served from the add-in origin, not a CDN.
2. The page is cross-origin isolated, as required by the threaded runtime and
   VGI OAuth worker bridge.
3. The browser backend loads VGI, attaches the HTTPS Open Meteo catalog, sets a
   local timezone, and executes a SQL result through the visible Query Editor.

Set `CUPOLA_LIVE_VGI_ENDPOINT=https://...` to use another anonymous test VGI
catalog. Keep this separate from the deterministic unit/UI job when network
access is unavailable.

## Windows integration tests

Run the complete Windows build and test pipeline from PowerShell on a machine
with Excel installed:

```powershell
.\tests\run-windows.ps1 `
  -HaybarnPath C:\path\to\haybarn.exe `
  -VgiExtensionPath C:\path\to\vgi.duckdb_extension `
  -OdbcDriverPath C:\path\to\haybarn_odbc.dll
```

If Node is unavailable on the Windows test host, build `apps/desktop/dist` on
the development machine first, copy the repository, and add `-SkipWebBuild`.

The runner performs:

1. Desktop connection, agent SQL policy, and ribbon contract tests.
2. Worker-free local Haybarn compatibility queries covering DuckDB temporal
   resolutions, DST boundaries, Excel date boundaries, every fixed-width
   integer family, decimal width transitions, floating-point bounds, BIGNUM,
   NaN, and infinities through the exact values handed to Excel.
3. Packed 32/64-bit XLL and MSI builds.
4. Real Haybarn stdin integration against the HTTPS Open Meteo catalog.
5. Real Excel COM tests for XLL registration, `VGI_VALUE`, `VGI_CALL`, dynamic
   `VGI_QUERY` spills, diagnostics, and rejection of plain HTTP.
6. MSI database inspection for the x64 XLL, machine-wide Excel loader, Haybarn, and the VGI extension,
   the full WebView2 managed/native payload, and an
   assertion that no companion executable or source maps are present. The
   publisher also rejects source maps in the desktop bundle and staged updater.

The runner passes `-HaybarnPath` to the policy suite as `VGI_HAYBARN_PATH`, so
numeric and temporal checks use the real engine. When running the policy
executable separately, set that variable to avoid its captured-payload fallback.

The native policy suite also validates the Power Query M handoff: read-only SQL
enforcement, M/ODBC escaping, the `CupolaConnection` identity, and the absence
of endpoint or credential material in workbook formulas. Once the ODBC fork is
registered, the real Excel UI can create the query through the Query Editor's
**Power Query** button and verify Refresh All end to end.

The Excel test adds uniquely named connections, backs up the user's connection
registry/default and Excel XLL registration, restores them in `finally`, closes
only the Excel instance it created, and verifies that process exits. This keeps
a smoke-test package path from replacing the user's installed VGI Ribbon.
Pass `-SkipExcel` on build agents without Microsoft Excel.

After installing, close Excel and run `tests\excel\active-install-smoke.ps1`.
It verifies the expected version/build and active installation directory. For
the machine MSI it also verifies automatic COM loader activation without an
explicit `RegisterXLL` call. Optional `-ExpectedVersion` and `-ExpectedBuild`
arguments allow testing a release other than the current checkout.

OAuth network exchanges are deliberately mocked in the fast suite. A live
provider test requires a registered test client and redirect URI, so it should
run in a separate credentialed environment rather than the default test job.

`excel\credential-manager-smoke.cmd` runs the desktop bridge tests with a
required Windows Credential Manager round trip. Run it in an interactive user
session; service logons such as OpenSSH do not have a usable credential vault.

The interactive WebView smoke test also verifies that the native bottom-right resize grip selects both axes and that the window can grow vertically and horizontally.

Power Query policy tests reject missing identities, unrelated drivers, and mismatched connection metadata. Pass `-OdbcDriverPath` (or `CUPOLA_ODBC_DRIVER_PATH`) to the Windows suite. `tests\odbc\cupola-connection.ps1` uses a temporary driver registration and tests braced connection names, conflicting parameters, HTTPS/credential rules, DPAPI failures, and a live VGI query. It requires an elevated Windows session for temporary driver registration.

The native runner's `--power-query` mode tests the production workbook creation path and Excel Refresh All in a throwaway workbook. Set `CUPOLA_ODBC_DRIVER_PATH`; set `CUPOLA_TEST_INSTALLED_POWERQUERY=1` to test the installed Cupola driver without a temporary driver registration. The test also saves its temporary workbook, quits Excel, opens the workbook in a new Excel instance, and verifies a fresh result after another refresh. Use an interactive session with `CUPOLA_TEST_INTERACTIVE_POWERQUERY=1` for Excel's first-use **Default or Custom** authentication prompt. It adds and removes only a uniquely named connection in the user registry so Excel's separate Mashup process can resolve the fixture. No real OAuth credentials are needed.

Build 20260922.3 validation on Europa: the Windows suite, direct live ODBC queries, and active-install smoke pass. The interactive installed-driver Power Query test also passes initial load, changed-query Refresh All, and another changed-query refresh after saving the workbook and reopening it in a new Excel instance. Excel first-use ODBC authentication may require interactive confirmation. This live test uses an anonymous HTTPS VGI catalog; it does not establish successful refresh against an OAuth-protected catalog.

The WebView regression opens Connections with an empty isolated store, saves a connection, and checks that it survives reload. `VGI_EXCEL_CONFIG_HOME` also isolates the WebView2 browser profile. Power Query integration saves a local recovery copy of an existing registry under `artifacts/local-test-backups` before adding its fixture; these files contain local connection metadata and must never be committed or uploaded. Cleanup removes only the exact fixture name using the JSON array API. Do not pipe a Windows PowerShell `ConvertFrom-Json` array directly into cleanup filters: its array can be emitted as one pipeline item.

The enterprise deployment contract checks machine-wide loader registration, prerequisite guards, installer transaction ordering, and complete web-asset packaging. Production MSI installation no longer includes the developer updater or 32-bit XLL; the standalone developer updater still carries both XLLs. See `docs/windows-production-readiness.md` for the qualification gates.

`tests\packaging\migration-smoke.ps1` tests developer-to-MSI startup cleanup in
an isolated temporary registry subtree. It verifies that other add-ins retain
their exact values and registry types, similar filenames are not removed, stale
Add-in Manager entries are cleaned without OPEN entries, and reruns are safe.
It does not modify the user's real Office registration.

`tests\packaging\azure-signing-smoke.ps1` tests the Azure signing adapter offline.
It rejects credential-bearing configuration, non-Azure endpoints, unsigned or
untimestamped outputs, unexpected publishers, and attempts to Authenticode-sign
the VGI extension. Signing and signature inspection are test doubles; no Azure
request or real signature is generated. See `docs/azure-signing.md` for setup
and approval-dependent verification.

The Windows suite also runs `--native-sessions`: deterministic session-cache
lifecycle/concurrency tests plus the real pinned native C API, chunk/value
conversion (including NULs and nanoseconds), HTTPS attachment reuse, and cold/warm
timings. For this mode set `VGI_HAYBARN_NATIVE_PATH` to the pinned
`haybarn_odbc.dll` and `VGI_EXTENSION_PATH` to its matching VGI extension. Tests
use an isolated connection/OAuth configuration directory.

### Windows multi-catalog coverage

The policy suite tests profile resolution, missing/nested/duplicate members,
catalog collisions, referenced-member deletion, and every ODBC identity field.
`--native-sessions` queries the public Weather and Earthquakes catalogs together
and checks that changed member settings invalidate the persistent profile session.
`tests/odbc/cupola-connection.ps1` checks the same live cross-catalog query,
fail-closed resolution, and reopened profile ordering using an isolated registry.

For real Excel Power Query initial load, Refresh All, and workbook reopen against
both catalogs, set `CUPOLA_ODBC_DRIVER_PATH` to the built driver (with the reviewed
VGI extension alongside it), then run:

```powershell
dotnet run --project windows/Vgi.ExcelDna.Tests -c Release -- --power-query-profile
```

This creates uniquely named temporary public connections, removes only those
fixtures afterward, and leaves existing user connections intact. Set
`VGI_EXCEL_TELEMETRY=0` for all test runs. The Playwright profile test covers
creation, default-member selection, persistence, and both catalog trees at 360 px.

Set `CUPOLA_TEST_EXCEL_PID_LOG` to a local JSONL file to record the process ID and
UTC start time of each Excel instance created by the Power Query tests. An outer
timeout supervisor can use both values, together with the process name, to clean
up only test-owned Excel instances. Never terminate all Excel processes to clean
up a failed test.

The runner logs elapsed time at direct ODBC validation, initial worksheet load,
Refresh All, workbook reopen, and cleanup. These phase records contain no SQL,
connection names, endpoints, or credentials. A COM call can block on Excel's
authentication dialog, so an external timeout is still needed; a polling-loop
timeout alone cannot guarantee cleanup.

If a first refresh stays at **Getting Data**, bring Excel's **ODBC driver** dialog
to the foreground. Select **Default or Custom**, leave credential connection
string properties blank, and select **Connect**. Cupola's saved connection store
and Power Query's per-source authentication choice are separate. The public test
fixtures need no credentials. Excel's embedded dialog controls may be absent from
UI Automation inspection; an empty accessibility tree does not prove that no
dialog exists. Window captures must account for the Windows display scale.

After publishing Windows artifacts, an administrator terminal in the interactive
Windows desktop can run `tests\excel\power-query-profile-smoke.cmd`. It opens the
temporary test workbook visibly so first-use ODBC prompts can be answered. For
these public fixtures choose Default or Custom with no credentials. It tests
initial load, Refresh All, and reopening the workbook, then removes its fixtures.

### Static tables and legacy compatibility

The Windows suite now runs `--workbook-tables` with isolated test configuration
and a separate Excel instance. It checks static insertion creates no refresh
metadata or external connection, failed migration preserves the original table,
legacy refresh grows and shrinks the table, metadata survives save/reopen, and
retiring management preserves values and formula references. The Office unit test
asserts static insertion does not alter workbook settings. UI tests cover migration
confirmation, creation failure, preservation of the original, and explicit metadata
retirement. Interactive Power Query loading/Refresh All qualification remains a
separate test; these checks do not substitute for it.

The updater regression test distinguishes live Excel processes from terminated
process entries still listed by Windows. Only live Excel processes block updates
and fresh-start installation checks.

Europa qualification on 2026-09-26, installed build 20260925.4: real Excel Power
Query passed initial load, changed-query Refresh All, and another changed-query
refresh after reopening the workbook in a new Excel process, for both a single
Weather catalog and a Weather/Earthquakes profile. The initial stall was Excel's
first-use ODBC authentication dialog: a method-only thread trace showed
`CredentialsChallengeResolver.ShowResolveDialog`, and a DPI-aware window capture
revealed the dialog that UI Automation had missed. Selecting Default or Custom
with no credentials resolved the stall. Later refreshes and reopening needed no
further authentication prompt. These anonymous public-catalog tests do not
qualify OAuth-protected Power Query refresh.

Europa qualification on 2026-09-27, packaged build 20260927.17: the AI
Power Query destination path passed initial load, changed-query Refresh All,
and refresh after save/reopen in a new Excel process. The test also verified
worksheet normalization, unique sheet/table names, and preservation of an
existing table. Excel required Default or Custom with blank credentials on
first use of the temporary connection; later refreshes needed no prompt.

The interactive resize test starts at the supported minimum window size before
checking growth in both axes, so Windows' maximum window height on a small remote
desktop does not cause a false failure.

The Office WASM connection tests hold a real worker network request pending,
assert the 20-second inline timeout and unchanged endpoint, then retry against
public Open-Meteo. They also exercise the reported `.workers.de` typo. Run the
same checks with `npx playwright test --config playwright.office-wasm.config.ts
connection-test.spec.ts --browser webkit` after building the Office bundle to cover
WebKit as well as Chromium. The live catalog test checks `vgi_catalogs(location)`
on the bundled WASM engine. Unit tests cover deadline cleanup, early failure,
retry, and ignored late completion.

Connection-dialog coverage includes URL-first discovery, manual overrides,
untouched-name defaults, duplicate-name rejection, and explicit-only saving.
`tests/ui/connection-discovery.spec.ts` exercises the desktop bridge UI at 360×480;
the live Office suite exercises discovery through the bundled WASM engine.
The Windows policy suite checks deadline cancellation and bridge duplicate-name
rejection. `--native-sessions` also verifies pre-ATTACH catalog discovery, draft
tests that preserve the saved registry/default, the reported misspelled hostname,
and successful retry against the public endpoint.

Europa qualification for build **20260926.1**: the full Windows suite and
interactive Credential Manager/WebView checks passed. The developer updater
installed the build, and `active-install-smoke.ps1` verified the active version
and versioned installation directory. Saved connection files were not modified
by installation. Browser validation passed 130 unit tests, 32 UI tests, four
live Chromium tests, and three WebKit connection tests, including assertions
that connection errors appear inline without duplicate global banners.


The business-user connection flow is covered by `tests/ui/profiles.spec.ts`: no
type selector in ordinary setup, persistent workspace membership and default
catalog selection, shared catalog browsing and query execution, compatibility with
existing one-member profiles, and reset back to a new connection.
Narrow-layout checks cover 360×480 desktop and 300-pixel Office panes.

Build **20260926.2** received a UI/UX review of the business-user flow and final
360×480 desktop / 300-pixel Office screenshots. All 36 UI tests and four live
Office WASM tests passed. Europa passed packaging/native checks, interactive
Credential Manager/WebView checks, real Excel workbook and packed-XLL checks,
and active-install verification. Pre/post hashes of the saved connection and
default-connection files matched after installation.


Protected discovery regressions exercise a synthetic session bearer token through
real Office WASM and assert it reaches catalog discovery without saving a draft.
Incomplete sign-in stays inline and focuses Find catalogs. Native policy tests
cover safe credential quoting, nonpersistent engine OAuth, and refresh enablement;
`--native-sessions` exercises the signed extension's discovery API over HTTPS.
These automated tests do not complete a real provider's interactive login.

Build 20260926.4 qualification on Europa passed the Windows/native/MSI suite,
interactive Credential Manager and WebView checks, real workbook and packed-XLL
checks, and active-install verification. A separate interactive-session probe
successfully discovered catalogs using an existing saved OAuth session without
opening sign-in or writing the connection registry. Saved connection/default
file hashes were unchanged across installation and verification. The real Office
WASM suite passed six cases, including cached-bearer discovery and bounded retry.

Query cancellation is exercised with real CPU-bound SQL in Office WASM and the
Windows native session suite. Windows also checks queued cancellation, request
isolation, and preserved temporary tables; UI tests cover retained results and
completion races.

Results-window tests cover paging, literal cell rendering, narrow scrolling,
Office dialog chunk transfer/origin checks, and snapshot isolation after rerunning
a query. The interactive Windows WebView suite checks a separate native results
window, maximization, and refusal of connection/workbook bridge operations.

The compact editor UI regression covers the first loaded preview at 360–1060 px
with a saved tall SQL split: three complete rows must fit, the bottom action bar
must stay inside the window, and More must support keyboard dismissal and outside
clicks. Native ribbon policy checks exclude the retired bulk-refresh commands.

Agent cancellation tests cover interrupted streaming, retry waits, initial catalog
inventory, SQL tools, catalog tools, and continuing the same conversation. The
desktop UI checks the Stopping state, duplicate-send prevention, and draft
preservation. Office WASM runs and stops a real CPU-bound agent query, resumes
the conversation, and verifies the editor session remains usable. Native session
tests exercise both editor and agent request IDs, including queued cancellation.

Model discovery tests cover pagination, capability-driven requests, token limits,
credential/workspace cache isolation, timeouts, safe error messages, daily cache
reuse, manual overrides, and preserving selections on refresh failures in both
hosts. Browser tests mock provider responses and use synthetic credentials.

Ask AI experience regressions check explicit clarification/continuation, cancelling
a waiting question, retained per-answer results, SQL handoff without execution,
Excel confirmation, source/sampling descriptions, and compact layouts. The Office
WASM test exercises these with real SQL; provider messages are deterministic mocks.
Unit tests cover repeated-tool and query-failure limits, continuation after stopping,
and numerical-only local token/cache diagnostics.

Ask AI transcript coverage verifies per-reply model attribution after switching
models, ordered query/result rendering, provider thinking summaries in JSON and
SSE, exclusion of signatures/redacted blocks from the displayed timeline, and
cancellation with earlier results intact. The live Office suite exercises the same
query/result grouping and summaries with a real WASM query.

The staged Ask AI new-worksheet tests assert that no workbook change occurs
before confirmation, and that confirmation passes the original SQL and saved
connection to Power Query without preview rows. They cover successful load,
query-only partial success, and driver failure without a static fallback. The
real `--power-query` suite also supplies AI-style worksheet/table names, verifies
normalization and collision handling without changing existing data, then checks
Refresh All and save/reopen refresh.

Office development checks run with `npm run test:office-dev`. They verify the
worker/WASM routes and metadata CORS, plus the Ask AI insertion confirmation
in WebKit with browser confirmation dialogs disabled. Cancel and Escape must
leave the workbook untouched; approving inserts once. Native HTML dialogs
provide in-app confirmation because Office.js disables `window.confirm`.

The Office development suite also checks host-controlled light/dark themes,
result-window appearance, and connection/draft restoration without carrying
OAuth tokens into a new pane session. Track the remaining real Excel sign-in,
worksheet-function, and save/reopen checks in
[Office release qualification](../docs/office-release-qualification.md).

Workspace regression coverage lives in `tests/ui/profiles.spec.ts`,
`apps/desktop/src/workspace.test.ts`, and the native `WorkspaceTests`. It checks
multiple visible catalogs, the shared query connection, default catalog changes,
legacy formula defaults, stable Power Query attachment identities, invalid
selection, and narrow-width connection controls. Run `tests\run-windows.ps1`
with `-SkipWebBuild` after copying a desktop bundle built on macOS.

## GitHub-hosted release checks

The release workflow runs all web checks and the Windows suite with `-SkipExcel`,
then validates signed MSI contents. Its `signing-test` mode performs a real Azure
OIDC signing request on a disposable file; the default Windows suite's Azure
tests remain offline. Signed GitHub releases remain drafts until real Excel and
installation qualification is recorded. See [GitHub releases](../docs/github-releases.md).

## MSI updater validation

The desktop policy suite includes updater release parsing, version ordering,
trusted download origins, checksums, and rejection of unsigned files.
`tests/ui/updates.spec.ts` covers the update flow at 360×480, daily preferences,
and the explicit installation handoff. The Windows suite also checks the new
MSI's identity through the updater's own Windows Installer reader.

To verify Authenticode handling against a production-signed MSI, run:

```powershell
dotnet run --project windows/Vgi.ExcelDna.Tests -c Release -- --verify-update-publisher C:\path\CupolaForExcel.msi
```

See [Windows updates](../docs/windows-updates.md) for the signed upgrade release
gate and enterprise policy. Offline tests do not replace this lifecycle check.

## OAuth session recovery

`auth-recovery.test.ts` in core, Office, and desktop covers interactive-error
classification, the WASM/native bridge boundaries, and repeated failures without
query replay. `reauthenticate.test.ts` covers cancellation and successful session
reset in both hosts. `tests/ui/auth-recovery.spec.ts` covers the persistent action,
profile navigation, permission errors, retained SQL drafts, and the 300px Office
and 360px desktop layouts. These use synthetic errors and mocked sign-in; they
neither contact Entra nor send Sentry events. A real Entra expiry/revocation test
remains part of manual OAuth qualification.

## Packaged VGI coverage

`npm run test:extensions` checks artifact hashes, footer ABI/platform/revision,
cache validation, and gzip decoding with deterministic fixtures. `npm test`
includes it and the Office variant/path selection tests.
`npm run test:office-wasm` also loads the packaged extension with upstream VGI
requests blocked in threaded and non-isolated WASM runtimes, and verifies that
a missing deployment asset fails without a community fallback. These remain
live HTTPS catalog tests, with `CUPOLA_LIVE_VGI_ENDPOINT` supported.

The Windows policy suite checks explicit-path loading and missing-file failure.
The ODBC integration temporarily removes the extension from its isolated test
package and verifies a repair message, then restores it for the live query.
Use the approved binary from [the VGI lock](../vgi-extensions.lock.json) when
running `tests\run-windows.ps1`; packaging rejects other revisions.
