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

These tests require no Excel installation, VGI credentials, or model API key.
The Playwright suite renders both task panes at 300, 320, 340, 400, 720, and
1060 pixels and checks onboarding, connection settings, footer visibility, and
horizontal containment. It also creates multiple Ask AI conversation tabs,
reloads both the desktop Workbench and Microsoft 365 task pane, and verifies
that the selected tab, visible transcript, model history, and connection
scoping survive without persisting API keys or process-local query result IDs.
Accounting numeric tests keep complete DECIMAL columns numeric when every value
fits Excel's 15-significant-digit limit, retain unsafe precision as text, apply
scale-aware number formats, and preserve leading-zero account codes.

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
  -VgiExtensionPath C:\path\to\vgi.duckdb_extension
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
6. MSI database inspection for the two XLLs, Haybarn, and the VGI extension,
   updater/registration files, the full WebView2 managed/native payload, and an
   assertion that no companion executable is present.

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

After installing the XLL, `tests\excel\active-install-smoke.ps1` verifies that
the persistent registration points to a loadable VGI build.

OAuth network exchanges are deliberately mocked in the fast suite. A live
provider test requires a registered test client and redirect URI, so it should
run in a separate credentialed environment rather than the default test job.

`excel\credential-manager-smoke.cmd` runs the desktop bridge tests with a
required Windows Credential Manager round trip. Run it in an interactive user
session; service logons such as OpenSSH do not have a usable credential vault.
