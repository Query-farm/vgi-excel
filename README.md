# Cupola for Excel

Cupola for Excel brings Query Farm's Vector Gateway Interface to Excel through a
shared-runtime Office add-in and an Excel-DNA package for classic Windows Excel.
All connections use secure HTTPS VGI endpoints; local command and subprocess
connector locations are intentionally unsupported.

The repository contains:

- `apps/office` — the Microsoft 365 task pane and JavaScript custom functions.
- `packages/core` — runtime-neutral query, formula, value-conversion, and agent logic.
- `apps/desktop` — the embedded WebView2 Cupola and streaming data-agent UI for the Windows XLL.
- `windows/Vgi.ExcelDna` — ribbon, native HTTPS/OAuth bridge, fallback workbench, and equivalent XLL worksheet functions for older Excel.
- `installer` — enterprise Windows installer inputs.

## Developer quick start

Use Node.js 22.12 or newer. Install the versions recorded in the lockfile:

```sh
npm ci
npm run check
npm test
npm run build
npm run test:ui
npm run test:office-wasm
npm run dev
```

Windows builds and real Excel integration tests are driven by
[`tests/run-windows.ps1`](tests/run-windows.ps1). See
[`tests/README.md`](tests/README.md) for the isolated Europa/Windows workflow.

The development manifest is emitted at `apps/office/dist/manifest.xml`. See
[`docs/development.md`](docs/development.md) for sideloading and Windows setup.

When adding a connection, enter the HTTPS VGI endpoint and its catalog name separately from
the friendly connection name. For example, use `open_meteo` for the Open-Meteo
worker. Both runtimes issue an explicit
`ATTACH 'open_meteo' AS "open_meteo"` before querying it.

The Microsoft 365 runtime uses self-hosted Haybarn WebAssembly assets; the live
browser integration suite verifies that the packaged worker attaches the Open
Meteo HTTPS catalog. The Excel-DNA runtime keeps an in-process native Haybarn session per friendly
connection name, reusing its HTTPS VGI attachment across queries. There is no
background service, localhost certificate, or pairing step.

The Windows Cupola experience is an embedded WebView2 application. Its agent streams
directly from Anthropic, supports cancellation, retry/backoff, multi-turn
history, loop guards, schema tools, and paged query inspection. Native C#
revalidates agent SQL, owns OAuth, and requires explicit confirmation before
writing a result to Excel. The Anthropic key stays in protected storage unless
the user explicitly saves it for their Windows account; the XLL stores it
as a Windows generic credential, never in the workbook or connection registry.

On 64-bit Windows Excel, Power Query uses the bundled Cupola-compatible
Haybarn ODBC driver and participates in Refresh All. Before creating the query,
Cupola verifies the connected session's name, VGI catalog, HTTPS endpoint,
authentication mode, and ATTACH options. The workbook M formula stores only
SQL and the friendly connection name. The driver reads the shared per-user
connection registry and encrypted OAuth session. On first use, Excel may ask
for ODBC authentication: choose **Default or Custom**, leaving credentials
blank. **Load into Excel** is the primary Windows insertion action.
**Insert snapshot** creates an ordinary static Excel table with no refresh metadata
in either host. Existing Cupola-managed tables appear under **Workbook data →
Cupola tables** and retain their original refresh behavior.

On Windows, **Create refreshable copy** copies a Cupola query to a new Power Query
table after confirmation. The original table, its refresh information, and formulas
referencing it remain intact. Verify the new table and update dependent formulas
as needed, then choose **Keep as static table** to remove the original's Cupola
refresh metadata without deleting its data. No automatic in-place conversion occurs.

Production builds report privacy-filtered failures to separate Sentry projects
for the Microsoft 365 host, desktop WebView, and native XLL. Error reports carry
only product/build, host, HTTPS transport, operation, exception type, and a
scrubbed stack trace. Cupola does not send SQL, query results, AI prompts or
responses, credentials, connection URLs, catalog/table/sheet/workbook names, or
workbook values. Browser telemetry is disabled on local Vite development servers
by default; all browser telemetry can be disabled with
`VITE_SENTRY_ENABLED=0`, and native telemetry with
`VGI_EXCEL_TELEMETRY=0`.

Ask AI supports Sonnet 5, Opus 5, and Haiku, optional Anthropic workspace IDs,
and model-specific thinking effort. Both web hosts use cache-aware conversations
and VGI metadata discovery. See [AI upstream port](docs/ai-upstream-port.md) for
ported commits, behavior, and the web-only features that remain separate.

## Worksheet API

```excel
=VGI.QUERY("select * from open_meteo.main.geocoding('Boston') limit 20")
=VGI.VALUE("select count(*) from open_meteo.main.geocoding('Boston')")
=VGI.CALL("open_meteo.main.weather_code_text", A2:A20)
```

`VGI.QUERY` and `VGI.VALUE` accept an optional named connection and refresh
key. `VGI.CALL` uses the workbook's default connection and supports scalar or
equally-shaped range arguments.

Excel 2016-2021 users loading the XLL directly use `VGI_QUERY`, `VGI_VALUE`,
and `VGI_CALL`; the dotted namespace is used by the Microsoft 365 add-in.

## License

Cupola for Excel is distributed under the [Query Farm Source-Available License
1.0](LICENSE), the same license used by `vgi-python`.

### Multiple catalogs on Windows

In **Settings → Connections**, save each HTTPS VGI connection first. Under
**Workspace catalogs**, select the connections to use together, choose the
**Default catalog**, and select **Apply changes**. Query Editor, Ask AI, and
Catalog View share that native session; new Power Query tables retain its named
connection set for refresh. Existing formula defaults remain unchanged.
Use qualified names such as `open_meteo.main.weather_code_text(0)` and
`earthquakes.main.recent` in the same SQL query.

Profiles reference up to 16 saved connections with distinct catalog aliases.
Each member keeps its own endpoint, ATTACH options, and encrypted sign-in session.
A missing or invalid member fails the whole connection. Connections used by saved
profiles or retained workspace query sets cannot be deleted or renamed. Saved
profiles remain available under advanced controls on Windows.
Power Query still stores only `Driver={Cupola for Excel};CupolaConnection={Research};`.

Open **Settings** using the gear in the upper-right corner to manage
saved connections. The compact header contains Query Editor, Ask AI, and Catalog View.

Choose **Settings → About** for the version, build, Query.Farm website, and diagnostics.

While a query runs, **Run** becomes **Cancel**. Cancellation preserves your SQL and
previous results; wait for **Query cancelled.** before running another query.

Use **Open results in new window** in the bottom result action bar to inspect a fixed copy
in a separate, resizable window. The viewer pages and copies only the rows already
loaded; it does not rerun SQL or change the workbook. Microsoft 365 uses its dialog
API (DialogApi 1.2); standalone browser testing uses a popup. Results remain in memory.

The results panel appears only when the active query has rows to preview. Empty
queries report “No rows returned” in the editor status line.

Choose **Hide results** in the query editor to give the SQL editor the available
workspace. **Show results** restores the preview and its previous splitter position.
The collapsed preference is remembered locally; running a query does not reopen it.

The bottom result action bar places **Load into Excel** on the right on Windows, creating
a table refreshed through Excel Refresh All. **More** contains **Copy results**
and **Insert static table**. Microsoft 365 uses **Insert into Excel** for a static
table, with copying under **More**. Expanded previews reserve space for visible
rows even in compact windows. The Cupola ribbon no longer duplicates formula
recalculation or exposes a blanket Cupola-table refresh command; existing managed
tables retain their individual refresh action in workbook management.

Settings groups **Connections**, **AI settings**, and **About** in one page.
Open **Cupola** from the Excel ribbon, then use the upper-right gear to manage
connections; the ribbon has no separate Connections button.
AI model, thinking effort, and output limits apply to the named current conversation;
workspace ID and output limits are under Advanced options. Switching to Settings
preserves the conversation and draft, and credentials can be configured before
adding a connection. Query History has been removed; named query tabs remain.
Query execution shows a spinner beside Cancel. Catalog Refresh is in the sidebar
next to Show hidden. Busy indicators honor the system reduced-motion preference.

AI settings discover available models from Anthropic using the current API key
and optional workspace. The list refreshes when opened if its successful cache is
more than a day old; **Refresh models** updates it immediately. Failed requests
keep the saved list and show an inline retry message. The model dropdown shows
the full list; choose **Custom model ID…** to enter an ID manually. Saved
conversation models are never automatically replaced.
Provider metadata determines thinking choices and output limits. Only model
metadata and a credential/workspace fingerprint are cached, never the API key.

Ask AI keeps each executed query's preview with its answer. **Edit query** and
**Open in Query Editor** create a new tab without running SQL or replacing an
existing draft. Result cards open the captured rows in a separate window; Windows
can load a refreshable table, while Office inserts a complete static result. Excel
writes require confirmation. Incomplete Office previews direct users to the
Query Editor for a full-result insert. Result cards are session-only; persisted
SQL remains available in the conversation's tool details.

Cards distinguish the active connection, loaded/preview rows, and initial AI sample
from AI-described dates, filters, and assumptions. When scope is ambiguous, the
agent can pause for a clarification: choose an option or type an answer and press
**Continue**. **Stop** also cancels a pending clarification. Stage and active elapsed
time appear beside Stop; Cupola does not use DuckDB progress percentages.

Both agents limit repeated identical calls, stop SQL execution after three query
failures per turn, and cap a turn at 20 model rounds. **Settings → About → AI
performance diagnostics** exposes local session-only model and tool timings,
input/output token counts and prompt-cache reads/writes. These bounded numerical
records contain no conversation content and are never uploaded to Sentry.

Ask AI labels each new reply with its selected model and retains that attribution
when the conversation changes models. Older unattributed replies say AI assistant.
The transcript preserves the order of model text and tool calls, with each query
preview directly beneath its query. Supported models return a user-visible
thinking summary in a collapsible section; opaque signatures and redacted thinking
are never displayed. Summaries and transcript order are saved locally with the
conversation, while result rows remain session-only. Restored queries indicate
when their preview is no longer available. None of this content is sent to Sentry.

Ask AI uses a compact two-line composer with Send on the right. Press Escape
while Ask AI is active to stop the current interaction, including a pending
clarification, without losing the draft. Escape first dismisses controls that
handle it, such as tab renaming, and does not stop background AI work from Settings
or another workspace. Stop remains available with its keyboard-shortcut tooltip.

In the Windows Ask AI workflow, new-worksheet actions default to refreshable
Power Query tables. The original executed SQL and friendly Cupola connection are
retained with the staged action; **Load into Excel** explicitly confirms creation
and reruns the full query, rather than copying preview rows. Excel Refresh All
then uses the same Cupola ODBC path as Query Editor. Requested worksheet and table
names are normalized and made unique. A query created without a worksheet load
is shown as a partial success with instructions to finish loading the existing
query, avoiding duplicate creation. There is no automatic static fallback.
Explicit static replacements and Microsoft 365 browser insertions remain static
and are labeled accordingly.

### GitHub release builds

Use **Actions → Cupola for Excel Windows release → Run workflow → signed-draft**
to build and sign a release candidate without a local Azure login. The workflow
creates a draft release for Windows/Excel qualification before publication.
See [GitHub release instructions](docs/github-releases.md).
