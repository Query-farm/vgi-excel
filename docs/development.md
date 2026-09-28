# Development and sideloading

## Office add-in

Use Node.js 22.12 or newer. Install the locked dependencies and start the HTTPS development server:

```sh
npm ci
npm run dev
```

The first development-server run creates and trusts a Microsoft Office development certificate.
Unit tests and production builds do not generate or install certificates. The
live Playwright preview uses a separate, untrusted certificate under the ignored
`dev-certs/preview` directory; Playwright accepts it without changing OS trust.
Sideload `apps/office/public/manifest.xml` into Excel. The manifest points to
`https://localhost:3000` and configures one long-lived shared runtime for the
task pane, ribbon command, and custom functions. The development server serves
Haybarn workers and WASM directly from the installed package at `/haybarn/`;
a production build is not required first. Run `npm run test:office-dev` to
check these asset routes and custom-function metadata access.

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
that cross-origin-isolated context. Serve the public `/functions.json` metadata
with `Access-Control-Allow-Origin: *` for GET/HEAD (and OPTIONS if requested),
without credentials: Excel for the web fetches it from its own origin before
starting the shared runtime. The Vite development and preview servers apply
this rule only to that metadata path. Remote VGI services must allow that origin
through CORS and OAuth providers must register
`https://vgi-excel.example.com/oauth-dialog.html` as a SPA redirect URI.

Run the real browser/engine integration test with:

```sh
npm run test:office-wasm
```

It starts the production preview over HTTPS, verifies cross-origin
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

The x64 XLL calls the native C API in the pinned `haybarn_odbc.dll` directly.
It keeps one in-memory database/session per friendly connection name, loads the
bundled VGI extension, and attaches through HTTPS once. It does not require an
ODBC registration for worksheet or AI queries; Power Query still uses its
separate ODBC contract. Queries on one session serialize; distinct connections
have independent sessions. Settings/credential changes rebuild the session on
its next query, sign-out/removal closes it, and add-in shutdown disposes all
sessions. Failed queries evict their session without automatically replaying SQL.
OAuth refresh-token attachments retain the VGI extension's refresh behavior;
bearer-only expired sessions require signing in again.

Result reading uses native materialized chunks, preserving empty-result schemas,
UTF-8/NUL strings, exact large numeric values, nested JSON, and nanosecond timestamps.
Session time zones are applied to timestamp-with-time-zone values. The five-minute
query deadline interrupts native execution instead of killing a child process.
Use `VGI_HAYBARN_NATIVE_PATH` and `VGI_EXTENSION_PATH` only for development overrides.
`haybarn.exe` remains in the package for compatibility/integration tests, but is
not invoked by the XLL query path. `VGI_HAYBARN_PATH` now applies only to CLI tests.

To stage both packed XLLs, Haybarn, and the VGI extension—and optionally build
the MSI—run:

```powershell
.\windows\publish.ps1 -HaybarnPath C:\path\to\haybarn.exe `
  -VgiExtensionPath C:\path\to\vgi.duckdb_extension `
  -OdbcDriverPath C:\path\to\haybarn_odbc.dll -BuildMsi
```

Release signing is optional for developer builds. Production builds should add
`-CertificateThumbprint <thumbprint>`; the publisher timestamps and signs the
XLL/native payload and MSI. `artifacts\xll\release-manifest.json` records the
product version, build, file sizes, and SHA-256 hashes.

`HaybarnPath` must be the native binary under `haybarn_cli\_bin`, not the small
uv launcher under a virtual environment's `Scripts` directory; the launcher is
not relocatable.

Refresh the VGI extension when updating an older native package. In the matching
Haybarn CLI, `FORCE INSTALL vgi FROM community` downloads the current signed
extension; `SELECT install_path FROM duckdb_extensions() WHERE extension_name =
'vgi'` locates the file to pass as `-VgiExtensionPath`. An old extension can fail
against current VGI workers even when the Haybarn executable version matches.
Run the native HTTPS suite before distributing the package.

When copying a macOS-built desktop bundle to Windows for `-SkipWebBuild`, replace
the destination `apps/desktop/dist` directory instead of merging files into an
old build. The publisher rejects source maps in the web bundle or staged updater,
including files left behind by earlier builds.

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

The bundled Haybarn ODBC fork resolves the name through
`%LOCALAPPDATA%\QueryFarm\VgiExcel\desktop-connections.json`. It decrypts the
same Windows-user DPAPI OAuth session as the XLL, loads the signed VGI extension,
and connects using `ATTACH ... TYPE vgi LOCATION ...`. Credentials never enter
M, ODBC parameters, process arguments, or driver diagnostics. Missing names,
ambiguous names, HTTP endpoints, embedded credentials, reserved ATTACH options,
and conflicting ODBC parameters fail closed. Generic Haybarn DSNs are unchanged.

Before creating a workbook query, the XLL connects and reads the driver's
connection-scoped `cupola_connection_info()` contract, compares all identity and
configuration fields, and confirms `duckdb_databases()` contains the VGI
catalog. A driver display name alone is insufficient. `CUPOLA_ODBC_DRIVER_NAME`
can override the registered name, but the selected driver must pass this check.

The driver source is `~/Development/haybarn/haybarn-odbc` (Cupola changes on
base `82989a8db00573bbeb2da32e5b2fcc3baa7b6cd6`). Build from an x64 Visual Studio
Developer Command Prompt with CMake/Ninja:

```bat
cmake -S path\to\haybarn-odbc -B path\to\haybarn-odbc\build\cupola -G Ninja -DCMAKE_BUILD_TYPE=Release
cmake --build path\to\haybarn-odbc\build\cupola --target haybarn_odbc --parallel 8
```

Pass the resulting `build\cupola\bin\haybarn_odbc.dll` to
`windows\publish.ps1 -OdbcDriverPath ...` or set `CUPOLA_ODBC_DRIVER_PATH`.
The DLL is packaged next to `vgi.duckdb_extension`. The MSI registers it for
64-bit Windows. The developer updater uses `register-odbc.ps1`, elevating only
machine-wide registration when needed. Power Query currently requires x64 Excel.

On first use, Excel's ODBC credential dialog requires **Default or Custom**
with no extra credentials; Cupola supplies its own session. This is Excel's
per-source permission, separate from VGI OAuth. See the
[Microsoft ODBC connector documentation](https://learn.microsoft.com/en-us/power-query/connectors/odbc).

## XLL

Build both bitnesses through Excel-DNA:

```powershell
dotnet build windows\Vgi.ExcelDna\Vgi.ExcelDna.csproj -c Release
```

The XLL adds a **Cupola** ribbon and modeless WebView2 Workbench for HTTPS
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
  -VgiExtensionPath C:\path\to\vgi.duckdb_extension `
  -OdbcDriverPath C:\path\to\haybarn_odbc.dll
```

The Windows runner builds the release artifacts, executes native HTTPS queries,
loads the packed XLL in a private Excel instance, validates formulas and spill
results, checks HTTP rejection and diagnostics, and inspects the MSI contents.
It backs up and restores the user's desktop connection files even when a test
fails. See `tests/README.md` for individual commands and coverage.

## Managed Windows release builds

The managed MSI uses a machine-wide COM loader to activate the packed Excel-DNA
XLL for each Excel user. It installs the Cupola ODBC driver in Program Files.
The developer updater remains a separate per-user workflow; migrate its startup
registration before testing the MSI as described in
[enterprise deployment](enterprise-deployment.md).

A release build requires an x64 Visual Studio Developer shell with C++ build
tools, a Windows SDK, CMake, Ninja, Python 3, Git, and the .NET SDK. Build pinned
native inputs from a clean checkout with:

```powershell
python windows/build-native-inputs.py --vgi-extension C:\approved\vgi.duckdb_extension
.\windows\publish.ps1 `
  -HaybarnPath .\artifacts\native-inputs\haybarn.exe `
  -VgiExtensionPath .\artifacts\native-inputs\vgi.duckdb_extension `
  -OdbcDriverPath .\artifacts\native-inputs\haybarn_odbc.dll `
  -BuildMsi -Production -CertificateThumbprint YOUR_CERTIFICATE_THUMBPRINT
```

`windows/native-inputs.lock.json` pins upstream revisions and input checksums;
`windows/odbc/cupola.patch` contains the reviewed Cupola integration and Windows
build fixes. The builder emits provenance alongside the native files. Production
publishing verifies that provenance against the source lock and patch, then
requires valid timestamped signatures from the selected certificate. Its private
key must be accessible to the build identity in the Windows certificate store;
Azure Artifact Signing uses the alternative `-AzureSigningConfigPath` option.
See [Azure signing setup](azure-signing.md) for preparation, OIDC, and the
approval-dependent qualification steps.

The engine-signed VGI extension must remain byte-for-byte unchanged: adding an
Authenticode signature would invalidate its Haybarn signature. Do not sign that
file with SignTool. MSI payloads and deployment scripts are signed separately.

The manually triggered Windows release workflow requires a dedicated Excel
runner with the above tools. See [remaining release gates](windows-production-readiness.md)
for clean-machine, multiple-user, OAuth, and enterprise rollout qualification.
After an installation, run `tests\excel\active-install-smoke.ps1` with Excel
closed. It checks the repository's expected version/build and active directory;
for a managed installation it verifies automatic COM loader activation without
manually loading the XLL.

## Windows multi-catalog profiles

A saved desktop connection may instead contain `Members`, an ordered array of
existing friendly connection names. The first member supplies the default catalog;
all members are attached to the same native engine connection. Profiles have an
empty `Location`, anonymous root `Authentication`, and empty `AttachOptions`;
member definitions supply those settings. Empty or absent `Members` preserves the
single-catalog format. Nested profiles, duplicate members, ambiguous names,
missing members, and case-insensitive catalog collisions are rejected.

Native session fingerprints include the complete ordered ATTACH setup for every
member. Editing a member causes its profiles to reconnect on their next query.
Signing out invalidates dependent native profile sessions. Already-open ODBC
sessions retain their attachment state until disconnected; a new ODBC connection
resolves the current profile and encrypted per-user OAuth stores again.

The DSN-less Power Query string remains unchanged. `cupola_connection_info()`
returns contract version 1 for a single catalog and version 2 with one row per
profile member. Profile rows include `connection_name` (profile), `member_name`,
`catalog_alias`, `location`, `authentication`, and `attach_options`. The native
bridge validates every identity row and attached VGI catalog before creating M.
Neither endpoints nor credentials are embedded in workbook connection definitions.

## Static snapshots and Cupola table refresh (20260925.3)

New snapshot insertion never writes `_CupolaSnapshot_` defined names (Windows) or
`cupola.snapshot.*` settings (Office). Both insert the query values into ordinary
Excel tables. Passing source metadata fields to the native insert bridge does not
restore the old managed-table behavior. Windows defaults to **Load refreshable
table**, using the existing Power Query/ODBC path and Excel Refresh All.

Cupola table metadata remains readable and its manual refresh methods remain available
for compatibility. Windows **Create refreshable copy** resolves the source from
that metadata at the native boundary and creates a separate Power Query table.
It never deletes or modifies the original table or metadata, including when driver
validation, query creation, or initial loading fails. An initial background refresh
being started is not proof that loading succeeded. Users verify the new table and
explicitly choose **Keep as static table** when ready to remove old metadata.
Existing formulas keep referencing their original table until the user changes them.
Office keeps manual Cupola table refresh but does not offer the Windows ODBC migration.

## Connection test deadlines

The Office connection form tests a draft using a dedicated Haybarn WASM worker.
A 20-second deadline covers engine startup and attachment as well as the probe
query. Success, failure, and timeout dispose that worker; a stalled endpoint does
not leave the shared query runtime blocked. OAuth sign-in following an explicit
authentication challenge is outside the network-probe deadline, with a fresh
bounded probe after sign-in. The endpoint is never corrected or suggested by the
connection test. Errors remain inline and the form can be retried.

Both connection dialogs start with an HTTPS address and **Find catalogs**, using
`SELECT catalog FROM vgi_catalogs(location)` through Haybarn before ATTACH. A single
result is selected automatically; multiple results require a choice. Manual catalog
entry remains available when discovery is unsupported, restricted, or fails. A new
friendly name follows the catalog until edited; saved identities never change
automatically. Duplicate names are rejected. Hostnames are never corrected or
suggested. Advanced ATTACH options remain separate.

The XLL probes use disposable native sessions with a 20-second caller deadline,
five-second HTTP request timeouts, and no HTTP retries. Cancellation interrupts the
native query; its owning worker disposes handles only after the query exits. At most
one native probe can remain active, preventing repeated timeouts from accumulating
sessions. Human OAuth sign-in is outside the probe deadline. Discovery does not
attempt interactive sign-in; restricted discovery can use manual catalog entry.
Connection discovery, test, and save errors remain inline beside their actions;
they are not repeated in the global notice banner. Editing or retrying clears the
previous error.
Testing drafts (including profiles) never saves connection definitions or changes
the default connection. Explicit Save is required. OAuth sessions may still be
stored securely after a user completes sign-in.


## Simple connection setup (20260926.2)

Both hosts use **New connection → Server address → Find catalogs**, followed by
**Catalog** and **Connection name**. Advanced settings stay collapsed. Windows
shows a separate **Combine saved connections** action when two individual
connections are available; it never asks users to choose a connection type.
New combinations require two members. Existing one-member combinations remain
editable. Their default connection is available under **Advanced options**.
Combining connections does not merge or copy data; each saved connection retains
its own catalog and sign-in. Ordinary connection setup remains the same in Office.


Authenticated catalog discovery uses the VGI extension's named authentication
options, reusing the Office session token or the desktop's per-user encrypted
OAuth store. Each network probe is bounded to 20 seconds. On an authentication
challenge Cupola opens its existing sign-in flow outside that deadline, then
retries discovery. The dialog shows persistent sign-in progress and inline
failures; cancellation returns focus to Find catalogs. Manual entry remains
available. Discovery uses `oauth_cache := 'none'` so the engine does not create
another credential store, and never saves the draft connection.

The Office WASM engine currently uses one execution thread (`maximumThreads: 1`).
The pinned threaded build can stall the next statement after cancelling a pending
query with multiple execution threads. Pending queries still yield for Cancel;
we await its acknowledgement before permitting another run. Keep the live
cancel/reuse/temporary-table test green before raising this setting. This limits
CPU-heavy SQL parallelism in the browser; the Windows native engine is unchanged.

Ask AI cancellation propagates to model streaming/retry waits and SQL/catalog
tools. The desktop bridge uses read-only `query.agent` requests with unique IDs
and waits for cancellation acknowledgement. Office gives Ask AI its own backend
connection and serializes its pending statements so cancellation cannot target an
editor or formula request. The interface shows Stopping until pending work settles;
interrupted tools become Stopped, drafts remain intact, and the conversation can
continue. Cancellation does not dismiss an external OAuth sign-in window. Agent
conversations remain local and are never sent to Sentry.

The desktop host displays a native Cupola opening panel while WebView2 starts.
The workbench and results viewer send `ui.rendered` after their first React paint;
this hides the panel independently of connection discovery. A 45-second startup
timeout or failed navigation exposes recovery instead of leaving a blank window.
