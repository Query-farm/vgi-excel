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
  <a href="https://query.farm/vgi/"><img src="https://img.shields.io/badge/protocol-VGI-orange.svg" alt="VGI"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-Query_Farm_Source--Available-blue.svg" alt="Query Farm Source-Available License"></a>
</p>

<p align="center">
  <a href="https://github.com/Query-farm/vgi-excel/releases/latest">Download</a> ·
  <a href="#getting-started">Getting started</a> ·
  <a href="#what-you-can-do-with-cupola">Features</a> ·
  <a href="#documentation">Documentation</a>
</p>


---

Cupola brings remote data into Excel so you can build reports, explore questions,
and keep your analysis up to date without repeatedly exporting and importing
files. Connect to data catalogs that support Query Farm's **Vector Gateway
Interface (VGI)**, choose the data you need, and use it in familiar Excel tables,
formulas, charts, and PivotTables.

You can work in SQL or ask an AI assistant for help. Preview results in Cupola,
then decide what to add to your workbook.

## What you can do with Cupola

### Build reports you can refresh

On Windows, **Load into Excel** turns a query result into a Power Query table.
Use that table in your workbook and update its data with Excel's **Refresh All**.
The saved query keeps the SQL and connection selection needed to fetch the data
again.

For a fixed record of the data, insert a **static table** instead. Static tables
are available in both the Windows and Microsoft 365 Office add-ins and do not
change when the workbook refreshes.

### Ask questions in plain language

**Ask AI** helps you explore a connected catalog, find relevant tables and
functions, write SQL, and inspect the results. Refine your question in a
conversation, review the SQL it used, or open that SQL in the Query Editor to
adjust it yourself.

Each executed query has a result preview alongside the answer. On Windows, you
can load the full result into a refreshable Power Query table with a descriptive
name. In the Office add-in, you can insert a static result into the worksheet.

Ask AI uses Anthropic models with your own API key. Its SQL is read-only, and
changes to your workbook require your confirmation.

### Explore and query your data

**Catalog View** shows the tables and functions available through your selected
connections. **Query Editor** gives you named SQL tabs, result previews, and
cancellation for running queries. Keep several analyses open and return to your
saved tabs later.

On Windows, select multiple connections under **Workspace catalogs** to browse
and query them together. Use qualified catalog names to combine data from
different catalogs in one SQL query. A **Default catalog** makes queries shorter
when you mainly work with one source.

The workspace supports up to 16 connections with distinct catalog aliases.
Changing the workspace selection preserves the catalog sets already used by
saved Power Queries.

### Use data directly in worksheet formulas

Cupola's worksheet functions let you bring query results into cells and build
calculations around them:

| Function | What it does |
| --- | --- |
| `VGI.QUERY` | Returns a query result as a table of cells |
| `VGI.VALUE` | Returns a single value for use in a calculation |
| `VGI.CALL` | Calls a catalog function with values or ranges from your worksheet |

For example, with an Open-Meteo connection configured as the formula default:

```excel
=VGI.QUERY("select * from open_meteo.main.geocoding('Boston') limit 20")
=VGI.VALUE("select count(*) from open_meteo.main.geocoding('Boston')")
=VGI.CALL("open_meteo.main.weather_code_text", A2:A20)
```

`VGI.QUERY` and `VGI.VALUE` also accept a named connection and a refresh key.
`VGI.CALL` supports scalar or equally shaped range arguments. The function
namespace is **VGI**; direct XLL use in Excel 2016–2021 also provides the
`VGI_QUERY`, `VGI_VALUE`, and `VGI_CALL` aliases.

## Getting started

1. [Download the Windows installer](https://github.com/Query-farm/vgi-excel/releases/latest)
   and open **Cupola** from the Excel ribbon. For the Microsoft 365 Office add-in,
   follow the [Office setup instructions](docs/development.md#office-add-in).
2. Open **Settings → Connections**, choose **New connection**, and enter your
   HTTPS VGI server address. Select **Find catalogs**, choose a catalog, and save
   it with a friendly connection name.
3. On Windows, choose which **Workspace catalogs** to use together and select
   **Apply changes**.
4. Browse **Catalog View**, write SQL in **Query Editor**, or add your API key
   under **Settings → AI settings** and ask a question in **Ask AI**.
5. Preview your results and choose **Load into Excel** for a refreshable Windows
   table, or insert a static table.

You need access to an HTTPS VGI endpoint for the data you want to query. Each
connection has its own catalog and sign-in settings. Cupola keeps connection
credentials outside your workbook.

## Choose your Excel add-in

| Capability | Windows desktop add-in | Microsoft 365 Office add-in |
| --- | --- | --- |
| SQL editor, Ask AI, and catalog browsing | Yes | Yes |
| Worksheet functions | Yes | Yes |
| Static Excel tables | Yes | Yes |
| Power Query tables with Refresh All | Yes | No |
| Multiple-catalog workspace selection | Yes | No |
| Setup | Signed Windows installer | Hosted Office add-in and manifest |

The Windows installer targets **x64 Windows and x64 Microsoft 365 Click-to-Run
Excel**. It requires .NET Framework 4.8 or newer and the machine-wide WebView2
Evergreen Runtime. Close Excel before installing; installation requires
administrator access. See the [Windows setup guide](docs/development.md#windows-excel-dna-package)
for developer XLL installation and compatibility details.

In **Settings → About**, the Windows add-in can check for updates and download
a verified installer. Choose **Install update**, save your work, and close all
Excel windows; the updater waits for Excel to close before installing. Daily
checks run when you open Cupola and can be turned off. See the
[Windows updater guide](docs/windows-updates.md) for managed deployments.

For the Microsoft 365 add-in, download the
[installation manifest](https://cupola.query.farm/manifest.xml) and follow the
[Office installation instructions](docs/cloudflare-hosting.md#installing-in-microsoft-365).
[Release notes](https://github.com/Query-farm/vgi-excel/releases) describe package
availability and validation status.

## Your data and credentials

Passwords, access tokens, and API keys are never stored in workbooks, formulas,
or Power Query definitions. Windows OAuth sessions are encrypted for the current
Windows user, and saved AI keys use Windows Credential Manager. Office sign-in
material stays in the Office session.

If an OAuth session expires or needs your attention, Cupola shows **Sign in
again** for the affected connection. Complete sign-in, then run your query again;
Cupola preserves your draft and never automatically repeats the failed action.
For a workspace with multiple connections, choose **Open connections**, select
the connection that needs attention, and choose **Sign in again**.

Ask AI communicates with Anthropic using your configured key. This is separate
from Cupola's error reporting: diagnostic reports exclude SQL, query results,
workbook contents and names, AI prompts and responses, credentials, and connection
URLs. See [privacy and telemetry settings](docs/development.md#sentry-releases-and-source-maps).

Query Farm publishes its [Privacy Policy](https://query.farm/legal/privacy/) and
[Terms of Service](https://query.farm/legal/terms/). Cupola's software license is
provided in [LICENSE](LICENSE).

## Documentation

- [Development and setup](docs/development.md)
- [Workspace catalog behavior](docs/development.md#desktop-workspace-catalogs)
- [Testing](tests/README.md)
- [Building and signing releases](docs/github-releases.md)
- [Windows deployment qualification](docs/windows-production-readiness.md)
- [Office release qualification](docs/office-release-qualification.md)
- [Report an issue](https://github.com/Query-farm/vgi-excel/issues)

## License

[Query Farm Source-Available License 1.0](LICENSE).

Copyright © 2025–2026 [Query Farm LLC](https://query.farm).

VGI is bundled with each Cupola release and verified against a shared release
lock. See [VGI packaging and upgrades](docs/vgi-packaging.md) for how native and
Office extension versions are selected and qualified.
