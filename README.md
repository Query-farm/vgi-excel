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

```sh
npm install
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
Meteo HTTPS catalog. The Excel-DNA runtime starts
the bundled native `haybarn.exe` privately and sends SQL over stdin; there is no
background service, localhost certificate, or pairing step.

The Windows Cupola experience is an embedded WebView2 application. Its agent streams
directly from Anthropic, supports cancellation, retry/backoff, multi-turn
history, loop guards, schema tools, and paged query inspection. Native C#
revalidates agent SQL, owns OAuth, and requires explicit confirmation before
writing a result to Excel. The Anthropic key stays in protected storage unless
the user explicitly saves it for their Windows account; the XLL stores it
as a Windows generic credential, never in the workbook or connection registry.

On Windows, a successful Query Editor result can be loaded as an ordinary Power
Query backed by the Cupola ODBC driver, so it appears in Queries & Connections
and participates in Refresh All. The workbook M formula stores only SQL and the
Cupola connection name; endpoint, ATTACH options, and OAuth credentials remain
in Cupola's per-user stores. Point-in-time inserts remain available as formatted
Excel table snapshots with Cupola-managed refresh metadata.

Production builds report privacy-filtered failures to separate Sentry projects
for the Microsoft 365 host, desktop WebView, and native XLL. Error reports carry
only product/build, host, HTTPS transport, operation, exception type, and a
scrubbed stack trace. Cupola does not send SQL, query results, AI prompts or
responses, credentials, connection URLs, catalog/table/sheet/workbook names, or
workbook values. Browser telemetry is disabled on local Vite development servers
by default; all browser telemetry can be disabled with
`VITE_SENTRY_ENABLED=0`, and native telemetry with
`VGI_EXCEL_TELEMETRY=0`.

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
