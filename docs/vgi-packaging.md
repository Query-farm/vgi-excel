# VGI packaging and upgrades

Cupola releases include a tested VGI extension. End users receive extension
updates through Cupola updates; there is no Settings switch to download the
latest VGI independently. Updating an extension separately could leave Excel's
XLL, Power Query's ODBC driver, and Office using different untested combinations.
Loaded native extensions also require a fresh host process to change versions.

## Selecting the binary

`vgi-extensions.lock.json` is the source of truth. It records the extension
revision, Haybarn engine ABI, WASM npm package version, and the URL, decoded size,
and SHA-256 of each platform artifact. The current selection is VGI `ef59ff8`
for ABI `v1.5.5`, with Haybarn WASM `1.5.5-rc4`:

| Host | Artifact platform |
| --- | --- |
| Windows x64 XLL and ODBC | `windows_amd64` |
| Office baseline WASM | `wasm_mvp` |
| Office exception-handling WASM | `wasm_eh` |
| Office threaded WASM (`coi` engine bundle) | `wasm_threads` |

The builder verifies the checksum and signed footer's revision, ABI and platform.
Haybarn still verifies the original extension signature at load time. Never
modify or Authenticode-sign the extension itself.

## Build and runtime behavior

Run `npm run prepare:vgi` to prepare all platforms, or select Windows only:

```sh
npm run prepare:vgi -- --platform=windows_amd64 --output=artifacts/approved-vgi
```

Pass `artifacts/approved-vgi/windows_amd64/vgi.duckdb_extension` to the native
builder or `windows/publish.ps1`. Both reject a different checksum. Production
provenance and CI native cache keys include the shared VGI lock.

Office development and release builds automatically prepare all three WASM
variants. They ship under `vgi/<sha256>/<platform>/` in the web deployment.
The browser selects the matching variant and loads its explicit URL under the
deployment base. The Windows XLL and ODBC driver load the explicit file beside
their native runtime. Missing files fail locally instead of installing VGI from
community. An existing user-level DuckDB extension cache does not select the
version for these explicit loads.

This follows DuckDB's [explicit extension loading](https://duckdb.org/docs/current/extensions/advanced_installation_methods)
and [WASM extension distribution](https://duckdb.org/docs/current/clients/wasm/extensions).
`extension_directory` selects an installed-extension cache root; it is unnecessary
when loading an explicit packaged path. `custom_extension_repository` selects a
remote repository, so it is also unnecessary for these packaged VGI loads.

Artifacts are cached under ignored `artifacts/vgi/` directories and reverified
before use. The upstream URLs are not immutable: if upstream replaces a file,
a cold build fails its checksum check. Preserve approved binaries in release
storage, or change the lock URL to an approved HTTPS mirror serving the same
bytes. To seed the cache offline:

```sh
npm run prepare:vgi -- --source-dir=/path/to/approved-vgi
```

That directory must contain `<platform>/vgi.duckdb_extension` for Windows and
`<platform>/vgi.duckdb_extension.wasm` for each WASM platform. These must be the
decoded signed files, not gzip archives. Corrupt cache entries fail verification;
remove the affected cache entry and prepare again from an approved copy.

## Approving an update

1. Select artifacts for the exact engine ABI and each supported platform. For
   native candidate discovery, run `FORCE INSTALL vgi FROM community` in the
   matching Haybarn CLI. Plain `INSTALL` does not replace an installed version.
   `SELECT install_path FROM duckdb_extensions() WHERE extension_name='vgi'`
   locates the downloaded candidate. Do this during release preparation, not in
   the installed Cupola runtime.
2. Review the extension revision and record the URLs, decoded sizes and SHA-256
   values in the shared lock. Select the same approved revision for all four
   platforms; coordinate engine changes with both native and WASM locks.
3. Run `npm run prepare:vgi`, `npm run check`, `npm test`, `npm run build`,
   `npm run test:ui`, and `npm run test:office-wasm`. On Windows, build the pinned
   native inputs and run `tests\run-windows.ps1`, including ODBC and Excel tests.
   Qualify OAuth refresh and reauthentication against the supported service.
4. Increment `cupolaBuild`, retain the verified artifacts, and distribute through
   the normal signed Windows and Office release processes. Restart Excel after
   installing the native update; reload Office to use its new deployment.
