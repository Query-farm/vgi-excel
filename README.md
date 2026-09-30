<p align="center">
  <img src="apps/desktop/public/cupola-mark.svg" alt="Cupola for Excel logo" width="160">
</p>

<h1 align="center">Cupola for Excel</h1>

<p align="center">
  SQL, AI-assisted analysis, and refreshable data in Excel through <a href="https://query.farm/vgi/">VGI</a>.<br>
  Built by <a href="https://query.farm">🚜 Query.Farm</a>
</p>

<p align="center">
  <a href="https://github.com/Query-farm/vgi-excel/releases/latest"><img src="https://img.shields.io/github/v/release/Query-farm/vgi-excel" alt="Latest release"></a>
  <a href="https://github.com/Query-farm/vgi-excel/actions/workflows/windows-release.yml"><img src="https://github.com/Query-farm/vgi-excel/actions/workflows/windows-release.yml/badge.svg?branch=main" alt="Windows release workflow"></a>
  <a href="https://query.farm/vgi/"><img src="https://img.shields.io/badge/protocol-VGI-orange.svg" alt="VGI"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-Query_Farm_Source--Available-blue.svg" alt="Query Farm Source-Available License"></a>
</p>

<p align="center">
  <a href="https://github.com/Query-farm/vgi-excel/releases/latest">Download</a> ·
  <a href="#getting-started">Getting started</a> ·
  <a href="docs/development.md">Development</a> ·
  <a href="docs/github-releases.md">Release guide</a>
</p>

---

Cupola connects Excel to HTTPS data catalogs using Query Farm's **Vector Gateway
Interface (VGI)** and **Haybarn**, Query Farm's DuckDB distribution. Write SQL,
ask an AI assistant to explore your data, browse catalog tables and functions,
and bring results into a workbook.

- **Query Editor** — named SQL tabs, cancellable queries, and paged result previews.
- **Ask AI** — conversations that discover schemas, run read-only SQL, and propose
  Excel tables for your confirmation.
- **Catalog View** — browse the catalogs included in your current workspace.
- **Power Query on Windows** — refreshable tables that participate in Excel
  **Refresh All**, alongside static snapshots and worksheet functions.

## Download

The current public Windows release is
[**0.5.0, build 20260930.1**](https://github.com/Query-farm/vgi-excel/releases/tag/v0.5.0-20260930.1).
Packages are signed as **Query Farm LLC** through Azure Artifact Signing and
include timestamps, SHA-256 checksums, and build provenance.

| Release file | Use |
| --- | --- |
| `CupolaForExcel.msi` | Windows installer |
| `CupolaForExcel-Deployment.zip` | Installer and administration scripts |
| `CupolaForExcel-Updater.zip` | Developer XLL updater and its payload |
| `SHA256SUMS.txt` | Download checksums |
| `release-manifest.json`, `native-provenance.json` | Build identity and file provenance |

The MSI targets **x64 Windows and x64 Microsoft 365 Click-to-Run Excel**. It
requires .NET Framework 4.8 or newer and the machine-wide WebView2 Evergreen
Runtime. Close Excel before installation; machine-wide installation requires
administrator access. The developer updater is a separate installation path:
see [Windows setup](docs/development.md#windows-excel-dna-package).

CI builds, native/package tests, and signing verification passed for this release.
Qualification of the exact signed packages in real Excel and through the complete
installation lifecycle remains outstanding, as recorded in the release notes.
See [Windows deployment qualification](docs/windows-production-readiness.md).

## Getting started

1. Install the Windows package and open **Cupola** from the Excel ribbon. For the
   Microsoft 365 Office add-in, follow the [sideloading instructions](docs/development.md#office-add-in).
2. Open the upper-right **Settings** gear, then **Connections**. Choose **New
   connection**, enter an HTTPS VGI server address, and select **Find catalogs**.
   Choose a catalog, give the connection a friendly name, and save it.
3. On Windows, use **Workspace catalogs** to select the connections to use
   together. Choose a **Default catalog**, then **Apply changes**.
4. Browse **Catalog View**, write SQL in **Query Editor**, or configure an
   Anthropic API key in **Settings → AI settings** and open **Ask AI**.
5. Preview the result and choose **Load into Excel** on Windows for a refreshable
   Power Query table. Confirm the proposed workbook change before it is applied.

The friendly connection name, catalog alias, and HTTPS server address are separate
fields. Cupola connects through Haybarn's VGI `ATTACH` mechanism. Data connections
use HTTPS only; no companion service or pairing step is needed.

### Multiple catalogs on Windows

Query Editor, Ask AI, and Catalog View share the selected workspace's native
session. Catalog View shows all included catalogs. The default catalog determines
where SQL looks when a catalog name is omitted; use qualified names to query
other catalogs or combine them in one query.

A workspace supports up to **16 saved connections with distinct catalog aliases**.
Each connection retains its own endpoint, options, and sign-in session. Saved
profiles remain available under advanced controls.

Changing the workspace does not change existing worksheet formula defaults or
an existing Power Query's selected catalog set. Cupola retains named connection
sets for saved queries; connections referenced by those sets cannot be renamed
or deleted. See [workspace behavior](docs/development.md#desktop-workspace-catalogs).

### Refreshable tables and static snapshots

On Windows, **Load into Excel** creates a DSN-less Power Query using the bundled
Cupola ODBC driver. Its M formula contains the SQL and friendly connection name;
the driver resolves connection details and encrypted OAuth sessions from Cupola's
per-user stores. If Excel asks for ODBC authentication, select **Default or
Custom** and leave credentials blank.

**More → Insert static table** on Windows and **Insert into Excel** in the Office
add-in create ordinary static tables with no refresh metadata. **Open results in
new window** displays a fixed copy of the loaded preview without rerunning SQL.

Existing Cupola-managed tables retain their original refresh behavior under
**Workbook data → Cupola tables**. On Windows, **Create refreshable copy** creates
a separate Power Query table after confirmation. The original remains intact;
**Keep as static table** removes its old refresh metadata only after confirmation.

### Ask AI

Ask AI streams responses from Anthropic, discovers available models with your API
key, and supports model-specific thinking settings. Conversations and named query
tabs persist locally per connection; result rows stay in memory for the session.

Each executed AI query keeps its preview with its answer. **Edit query** or
**Open in Query Editor** opens SQL in a new tab without executing it. AI-generated
queries can supply descriptive names for the Power Queries created in Excel.
On Windows, loading an AI result reruns the full query through Power Query rather
than inserting only its preview rows. Office inserts static results.

The agent can ask for clarification before proceeding. **Stop** or **Escape**
while Ask AI is active cancels the interaction. AI SQL is read-only and is
validated again at the Windows native bridge; workbook writes require explicit
confirmation. See [AI implementation notes](docs/ai-upstream-port.md).

## Worksheet functions

With an Open-Meteo connection configured as the formula default:

```excel
=VGI.QUERY("select * from open_meteo.main.geocoding('Boston') limit 20")
=VGI.VALUE("select count(*) from open_meteo.main.geocoding('Boston')")
=VGI.CALL("open_meteo.main.weather_code_text", A2:A20)
```

`VGI.QUERY` returns a table; `VGI.VALUE` returns one value. Both accept an optional
named connection and refresh key. `VGI.CALL` invokes a catalog-qualified scalar
function using the default connection and supports scalar or equally shaped
range arguments.

The worksheet namespace remains **VGI** for compatibility. Direct XLL use in
Excel 2016–2021 provides `VGI_QUERY`, `VGI_VALUE`, and `VGI_CALL` aliases; this
compatibility path is separate from the MSI's Microsoft 365 installation target.

## Two Excel hosts

| | Windows desktop add-in | Microsoft 365 Office add-in |
| --- | --- | --- |
| Interface | Excel-DNA XLL with embedded WebView2 | Office task pane and shared runtime |
| Query engine | Native Haybarn | Self-hosted Haybarn WebAssembly |
| Excel tables | Power Query or static snapshots | Static snapshots |
| Multiple-catalog workspace controls | Available | Windows-only feature |
| Installation | Signed MSI or developer XLL updater | Hosted web assets and Office manifest |

The Office add-in is packaged separately from the Windows release. Production
hosting requires the WASM content type and COOP/COEP headers documented in the
[Office setup guide](docs/development.md#office-add-in). Track real Office-host
checks in [Office release qualification](docs/office-release-qualification.md).

## Credentials and privacy

Connection credentials and API keys are never stored in workbooks, formulas,
Power Query M, or connection JSON. Windows OAuth refresh sessions are encrypted
for the current Windows user; Office OAuth material stays in the Office session.
A saved Windows Anthropic API key uses Windows Credential Manager.

Production builds send privacy-filtered errors to separate Sentry projects for
Office, the desktop WebView, and the native XLL. Reports exclude SQL, results,
workbook values and names, AI prompts and responses, credentials, and connection
URLs. Tracing, replay, and breadcrumbs are disabled. Local Vite development does
not send telemetry by default. Set `VITE_SENTRY_ENABLED=0` at browser build time
or `VGI_EXCEL_TELEMETRY=0` in Excel's environment to disable error reporting.

## Development

Use **Node.js 22.12 or newer** and install the locked dependencies:

```sh
npm ci
npm run check
npm test
npm run build
npx playwright install chromium
npm run test:ui
npm run test:office-wasm
npm run dev
```

`test:office-wasm` uses a live HTTPS VGI catalog. Set `CUPOLA_LIVE_VGI_ENDPOINT`
to override its endpoint. The Office development server supports sideloading
`apps/office/public/manifest.xml`; see [development and sideloading](docs/development.md).

| Directory | Purpose |
| --- | --- |
| `apps/office` | Microsoft 365 task pane and custom functions |
| `apps/desktop` | Windows WebView2 workbench |
| `packages/core` | Shared query, formula, value-conversion, and agent logic |
| `windows/Vgi.ExcelDna` | Native Excel bridge, ribbon, OAuth, and XLL functions |
| `installer` | Windows MSI and deployment checks |
| `tests` | Browser, native, packaging, and real Excel validation |

Cross-compile native code with `dotnet build windows/Vgi.Excel.sln -c Release`.
Run Windows and real Excel validation with `tests\run-windows.ps1`; the complete
commands and interactive credential checks are in [tests/README.md](tests/README.md).

## Releases

Open [**Actions → Cupola for Excel Windows release**](https://github.com/Query-farm/vgi-excel/actions/workflows/windows-release.yml),
select **Run workflow** on `main`, and choose:

- **signed-draft** — validate, build, sign through Azure, and create a draft release.
- **candidate** — validate and upload an unsigned Windows test package.
- **signing-test** — verify Azure authentication and signing with a disposable file.

GitHub-hosted runners handle the builds; signing uses a dedicated Azure identity
with GitHub OIDC. Increment `cupolaBuild` in `package.json` and keep native version
declarations in sync for every installable update. CI does not run Excel; qualify
the exact signed artifacts on Windows with Excel before publishing a draft.
See the [release guide](docs/github-releases.md) and [Azure signing setup](docs/azure-signing.md).

Production Office packages use a separate command:

```sh
npm run package:office -- --base-url=https://your-production-origin
```

## License

[Query Farm Source-Available License 1.0](LICENSE).

Copyright © 2025–2026 Query Farm LLC.

## Links

- [Query.Farm](https://query.farm) — the team behind Cupola
- [Vector Gateway Interface](https://query.farm/vgi/) — the data protocol
- [Haybarn](https://github.com/Query-farm-haybarn/haybarn) — the query engine
- [Downloads and release notes](https://github.com/Query-farm/vgi-excel/releases)
- [Report an issue](https://github.com/Query-farm/vgi-excel/issues)
