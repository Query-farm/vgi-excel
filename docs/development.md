# Development and sideloading

## Office add-in

Install dependencies and start the HTTPS development server:

```sh
npm install
npm run dev
```

The first run creates and trusts a Microsoft Office development certificate.
Sideload `apps/office/public/manifest.xml` into Excel. The manifest points to
`https://localhost:3000` and configures one long-lived shared runtime for the
task pane, ribbon command, and custom functions.

For a production package:

```sh
npm run package:office -- --base-url=https://vgi-excel.example.com
```

This builds the add-in, copies the exact Haybarn WASM and worker artifacts into
`apps/office/dist/haybarn`, renders the production manifest, and rejects a
package that still contains localhost URLs. Host everything under
`apps/office/dist` at that HTTPS origin. The host must serve `.wasm` as
`application/wasm` and return `Cross-Origin-Opener-Policy: same-origin` plus
`Cross-Origin-Embedder-Policy: require-corp`; VGI's worker OAuth bridge requires
that cross-origin-isolated context. Remote VGI services must allow that origin
through CORS and OAuth providers must register
`https://vgi-excel.example.com/oauth-dialog.html` as a SPA redirect URI.

Run the real browser/engine integration test with:

```sh
npm run test:office-wasm
```

It starts the production preview over trusted HTTPS, verifies cross-origin
isolation, loads the self-hosted worker/WASM, attaches Open Meteo over VGI, and
executes SQL. Override the endpoint with `CUPOLA_LIVE_VGI_ENDPOINT`.

### Sentry releases and source maps

The Microsoft 365 app, desktop WebView, and native XLL report to separate Sentry
projects. Their public ingest DSNs are compiled into production packages and may
be overridden with `VITE_SENTRY_OFFICE_DSN`,
`VITE_SENTRY_DESKTOP_DSN`, and `VGI_EXCEL_SENTRY_DSN`. Use
`VITE_SENTRY_ENVIRONMENT` for the two browser hosts and
`VGI_EXCEL_SENTRY_ENVIRONMENT` for the XLL. The release name is
`cupola-excel@<version>+<build>` and the distributions are `office`, `desktop`,
and `xll`.

Browser source maps are generated only when both `SENTRY_AUTH_TOKEN` and
`SENTRY_ORG` are present for a production Vite build. The build uploads maps to
`cupola-excel-office` or `cupola-excel-desktop` (override with
`SENTRY_OFFICE_PROJECT` or `SENTRY_DESKTOP_PROJECT`) and then deletes the local
maps, so source maps are not deployed or included in the Windows installer.
Keep the Sentry organization auth token in CI/repository secrets; never prefix
it with `VITE_` or put it in an `.env` file committed to this repository.

Telemetry is intentionally error-only: tracing, replay, logs, metrics, automatic
sessions, and breadcrumbs are disabled. Before-send filters remove request/user
contexts, SQL, results, prompts/responses, credentials, URLs, workbook metadata,
customer identifiers, source context, and local user paths. Set
`VITE_SENTRY_ENABLED=0` at browser build time or
`VGI_EXCEL_TELEMETRY=0` in the Excel process environment for the native kill
switch. Local Vite servers report nothing unless
`VITE_SENTRY_ENABLE_LOCAL=1` is explicitly set.
The XLL also disables Sentry's process-global unhandled and unobserved-task
hooks, so it reports only failures explicitly captured by Cupola and never
exceptions raised by Excel or another add-in in the shared Excel process.

## Windows Excel-DNA package

The XLL drives the released `haybarn.exe` CLI as a private child process. It
sends SQL over stdin, keeping connection details out of process arguments, and
accepts only HTTPS VGI locations. `haybarn.exe` and `vgi.duckdb_extension` are
installed beside the packed XLL. Set `VGI_HAYBARN_PATH` or
`VGI_EXTENSION_PATH` only for development overrides.

To stage both packed XLLs, Haybarn, and the VGI extension—and optionally build
the MSI—run:

```powershell
.\windows\publish.ps1 -HaybarnPath C:\path\to\haybarn.exe `
  -VgiExtensionPath C:\path\to\vgi.duckdb_extension -BuildMsi
```

Release signing is optional for developer builds. Production builds should add
`-CertificateThumbprint <thumbprint>`; the publisher timestamps and signs the
XLL/native payload and MSI. `artifacts\xll\release-manifest.json` records the
product version, build, file sizes, and SHA-256 hashes.

`HaybarnPath` must be the native binary under `haybarn_cli\_bin`, not the small
uv launcher under a virtual environment's `Scripts` directory; the launcher is
not relocatable.

Install or update the staged XLL from one permanent, per-user registration:

```powershell
.\artifacts\xll\install-xll.ps1 -PackagePath .\artifacts\xll
```

For a local developer install, double-click **Update Cupola for Excel.cmd** in
the staged package to run the same command. It writes diagnostics to
`%TEMP%\Cupola-for-Excel-update.log`.

The updater copies each build to a versioned directory under
`%LOCALAPPDATA%\QueryFarm\VgiExcel\AddIn`, removes stale VGI XLL registrations,
and registers the correct 32-bit or 64-bit XLL. It closes Excel before switching
versions and reopens it when the update is ready. Save workbook changes first;
if an unsaved workbook is detected, the update stops without closing Excel.
Windowless orphaned Excel processes are terminated so they cannot retain the
previous XLL. Excel-DNA assemblies cannot be unloaded safely from a running
Excel process, so a process restart is required to activate new code. No
companion service runs in the background.

### Power Query and the ODBC fork

The desktop Query Editor's **Power Query** button creates an M query using this
DSN-less contract:

```text
Driver={Cupola for Excel};CupolaConnection={friendly connection name};
```

The ODBC driver resolves that name through
`%LOCALAPPDATA%\QueryFarm\VgiExcel\desktop-connections.json`, including the VGI
HTTPS location, catalog alias, and ATTACH options, and uses Cupola's Windows
OAuth session rather than exposing tokens in M. During fork development, set
`CUPOLA_ODBC_DRIVER_NAME` to the registered driver display name. The created
query remains in Excel's Queries & Connections even when the driver is absent,
which lets UI tests validate the handoff before driver integration is present.

## XLL

Build both bitnesses through Excel-DNA:

```powershell
dotnet build windows\Vgi.ExcelDna\Vgi.ExcelDna.csproj -c Release
```

The XLL adds a **VGI** ribbon and modeless WebView2 Workbench for HTTPS
connections, SQL testing, catalog exploration, a streaming agent, confirmed
result insertion, refresh, and diagnostics. The Microsoft Edge WebView2
Evergreen Runtime is required; current Microsoft 365 installations normally
provide it. If WebView2 cannot initialize, the XLL offers its native WinForms
Workbench as a fallback. Set `VGI_EXCEL_NATIVE_WORKBENCH=1` to force that
fallback for diagnostics. It
registers `VGI.QUERY`,
`VGI.VALUE`, and `VGI.CALL` for equivalent-add-in conversion, plus direct
legacy aliases `VGI_QUERY`, `VGI_VALUE`, and `VGI_CALL`. Excel 2016/2019 users
must use the underscore aliases, select an output range, and confirm
array-returning formulas with Ctrl+Shift+Enter; newer Excel spills them
automatically.
Detailed XLL failures are written to
`%LOCALAPPDATA%\QueryFarm\VgiExcel\xll.log`; worksheet cells return `#N/A` and
`VGI_LAST_ERROR()` returns the latest diagnostic in the current Excel process.
Structured agent lifecycle and tool diagnostics are written as redacted NDJSON
to `%LOCALAPPDATA%\QueryFarm\VgiExcel\agent.log` (rotated at 5 MB). User prompt
text, model response text, and the Anthropic key are not recorded; tool SQL and
errors are retained after credential-pattern redaction so query failures can be
diagnosed.
Structured OAuth discovery, callback, token-exchange, persistence, and ATTACH
events are written as redacted NDJSON to
`%LOCALAPPDATA%\QueryFarm\VgiExcel\oauth.log` (rotated at 2 MB). A flow ID ties
the stages of one browser sign-in together. The log includes HTTP status codes,
elapsed time, token presence and lengths, but never tokens, authorization
codes, PKCE verifiers, client secrets, or bearer headers.
Credential-free connection definitions are stored at
`%LOCALAPPDATA%\QueryFarm\VgiExcel\desktop-connections.json`.
OAuth connections use RFC 9728 discovery and a system-browser PKCE flow. The
refresh session is encrypted with Windows DPAPI for the current user; access
and identity tokens remain in memory. The Workbench can also save or forget the
Anthropic key as a per-user Windows generic credential. It is loaded into the
password field only while the Workbench is running and is never written to the
workbook or connection registry.

The embedded agent adapts the production loop used by `vgi-web-frontend`:
Anthropic SSE streaming, cancellation, bounded retries, multi-turn history,
repeated-tool protection, result paging, and visible tool progress. The web UI
can request only the narrow operations exposed by the native bridge. C# applies
the final SQL policy and blocks mutation, local file readers, arbitrary URL
scans, environment access, secret inspection, and external-database scanners.
Catalog and agent function discovery combine the complete attached inventory
from `duckdb_functions()` with per-argument VGI metadata from
`vgi_function_arguments()`. The latter supplies named-vs-positional semantics,
constraints, defaults, choices, patterns, and descriptions; DuckDB metadata
retains zero-argument and non-VGI callables, examples, tags, and return types.

## Automated tests

Run the cross-platform suite with
`npm run check && npm test && npm run build && npm run test:ui`. Run the live
Microsoft 365 engine test separately with `npm run test:office-wasm`.
On a Windows machine with Excel, run:

```powershell
.\tests\run-windows.ps1 `
  -HaybarnPath C:\path\to\haybarn.exe `
  -VgiExtensionPath C:\path\to\vgi.duckdb_extension
```

The Windows runner builds the release artifacts, executes native HTTPS queries,
loads the packed XLL in a private Excel instance, validates formulas and spill
results, checks HTTP rejection and diagnostics, and inspects the MSI contents.
It backs up and restores the user's desktop connection files even when a test
fails. See `tests/README.md` for individual commands and coverage.
