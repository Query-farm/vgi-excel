# AGENTS.md

## Product

This repository builds **Cupola for Excel**, an Excel client for Query Farm's
Vector Gateway Interface (VGI). Preserve the Cupola product name in user-facing
copy. The worksheet-function namespace remains `VGI` for compatibility.

Cupola has two hosts:

- `apps/office`: Microsoft 365 Office add-in using Haybarn WebAssembly.
- `windows/Vgi.ExcelDna` plus `apps/desktop`: classic Windows Excel XLL using
  Excel-DNA, native Haybarn, and an embedded WebView2 interface.

Shared TypeScript logic lives in `packages/core`. Windows packaging lives in
`windows`, `installer`, and `tests`.

## Non-negotiable architecture

- Excel connections use HTTPS VGI endpoints only. Do not add HTTP, subprocess,
  localhost-companion, or local connector transports.
- Connect through DuckDB/Haybarn `ATTACH ... TYPE vgi LOCATION ...`; do not
  replace the VGI catalog with copied HTTP API logic.
- A friendly connection name is the session/workbook identity. Catalog alias,
  HTTPS location, and ATTACH options are separate fields.
- Never put access tokens, refresh tokens, API keys, passwords, or authorization
  headers in workbooks, formulas, Power Query M, connection JSON, logs, or
  process arguments.
- Desktop OAuth refresh sessions are encrypted for the current Windows user.
  Microsoft 365 OAuth material remains in the Office session.
- AI SQL is read-only and must be validated again at the native bridge boundary.
  Workbook writes always require explicit user confirmation.
- Snapshots and Power Query are distinct. Snapshots are Cupola-managed Excel
  tables; Power Query participates in Excel Refresh All through the ODBC driver.

## Power Query contract

The desktop add-in creates DSN-less Power Query sources with:

```text
Driver={Cupola for Excel};CupolaConnection={friendly connection name};
```

`CUPOLA_ODBC_DRIVER_NAME` may override the registered driver display name. The
ODBC driver must resolve `CupolaConnection` through the same per-user Cupola
connection and OAuth stores. Keep endpoint and credential material out of M.

## UI conventions

- Workspace order is Query Editor, Ask AI, Catalog, Connections.
- Prefer inline, persistent state over success toasts. Toasts/notices are for
  errors, retries, or genuinely transient work.
- Use the Cupola mark and existing design tokens/icons; do not introduce generic
  document/database branding where a Cupola asset exists.
- Query tabs and AI conversation tabs persist locally per connection. Process-
  local query-result IDs must not be persisted.
- Normalize AI/user-proposed worksheet names before confirmation and again in
  the Excel bridge. Excel names are at most 31 characters, exclude
  `\ / ? * [ ] :`, and must be made unique rather than rejected.
- The desktop WebView must remain usable down to 360×480 and the Office taskpane
  down to 300 px wide. Keep scrolling inside content panes where practical.

## Versioning and packaging

`package.json` is the source of truth for `version` and `cupolaBuild`. Web builds
read it directly. Native and manifest declarations must match; `npm run check`
enforces this. Increment `cupolaBuild` for every installable update.

Production Office packages are created with:

```sh
npm run package:office -- --base-url=https://your-production-origin
```

The production host must serve WASM correctly and send the COOP/COEP headers
documented in `docs/development.md`.

Windows artifacts are built by `windows/publish.ps1`. Production signing uses
its `-CertificateThumbprint` option. The developer updater is always:

```text
artifacts\xll\Update Cupola for Excel.cmd
```

It must close Excel safely, install the versioned XLL, update registration, and
reopen Excel. Do not rely on hot-reloading an XLL into an existing Excel process.

## Required validation

Run proportional tests for every change. Before a release, run all of:

```sh
npm run check
npm test
npm run build
npm run test:ui
npm run test:office-wasm
```

`test:office-wasm` is a live network integration against an HTTPS VGI catalog,
not a deterministic mock. Override it with `CUPOLA_LIVE_VGI_ENDPOINT` when
needed.

Windows/native changes must also be copied to `europa` and exercised from:

```text
C:\Users\rusty\vgi-excel-cupola-test
```

Use `tests\run-windows.ps1`; add `-SkipWebBuild` when the desktop web bundle was
built on macOS and copied over. The Windows suite covers the .NET policy layer,
native HTTPS, packed XLLs, real Excel formulas/spills, and MSI contents. Run
`tests\excel\credential-manager-smoke.cmd` interactively because an SSH logon
does not have the user's Windows Credential Manager session.

After installing, run `tests\excel\active-install-smoke.ps1` and verify that its
diagnostics report the expected version/build and active version directory.

## Repository hygiene

- Do not commit `node_modules`, build outputs, test reports, artifacts, signing
  material, development certificates, logs, local connection registries, or
  credentials.
- Preserve user changes in a dirty worktree. Avoid destructive Git commands.
- Add regression coverage for bugs at the lowest deterministic layer and at the
  host boundary when relevant.
- Keep README and files under `docs/` and `tests/README.md` current when commands,
  packaging, or supported behavior changes.
