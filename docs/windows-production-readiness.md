# Windows production deployment work

Target release: **0.5.0**, initially x64 Windows with x64 Microsoft 365
Click-to-Run Excel. The release is a candidate until the gates below pass.

## Implemented foundation

- Machine-wide COM loader activates the existing Excel-DNA XLL at Excel startup.
  The MSI owns the loader/ODBC registrations and installs code in Program Files;
  it does not create `HKCU` startup entries for the installing SYSTEM account.
- Complete desktop web assets are harvested into the MSI. A build-info file
  prevents packaging a desktop bundle from another version/build.
- Native installer checks detect running Excel and unsupported OS architecture
  before any transaction. MSI launch conditions also require x64 Click-to-Run
  Excel, .NET Framework 4.8+, and machine-wide WebView2 Evergreen Runtime.
- Major-upgrade removal runs inside the rollback transaction. The installer does
  not own or delete per-user connection, OAuth, AI, or browser state.
- Deployment wrapper supports unattended install/repair/uninstall, signature
  validation, logging, and a retry result while Excel is open. Detection checks
  version, files, and machine-wide loader/ODBC registration without opening Excel.
- Per-user connection writes use a cross-process file lock and atomic replacement
  with `.bak` recovery copies. Corrupt JSON is never silently replaced by an empty
  list. IT defaults add missing identities without changing existing definitions
  or the user's selected connection.
- Pinned ODBC source + reviewed Cupola patch, pinned Haybarn engine source, and
  checksum-verified CLI/extension inputs provide a repeatable native build recipe.
- Production publishing requires native provenance and timestamped signatures.
  A manually triggered CI workflow targets a dedicated Windows/Excel runner.

## Release gates still to prove

- Silent SYSTEM MSI install on a clean disposable Windows machine, followed by
  Excel startup as a standard user and a second/new user.
- Upgrade, repair, uninstall, blocked installation with Excel open, downgrade
  rejection, and failed-upgrade rollback; compare connection/OAuth bytes before
  and after each operation. Never run destructive qualification on a working
  customer profile.
- Real Excel startup and Power Query refresh from the actual MSI installation,
  including OAuth-protected catalogs, expired/renewed sessions, and workbook reopen.
- Production signing certificate or signing service, trust deployment, and a
  dedicated release runner. Europa currently has no code-signing certificate and
  no Hyper-V management installation for disposable machine testing.
- Approved distribution of the pinned signed VGI extension. Its current checksum
  is recorded in `windows/native-inputs.lock.json`; the release runner must obtain
  that exact file from an approved artifact store. The tested Haybarn input is
  `1.5.5-rc1`, which must be explicitly qualified for the production release.
- Production artifact publication and pilot assignment to an actual Intune tenant.

## Candidate validation recorded on 2026-09-22

For version 0.5.0 / build 20260922.5:

- `npm run check`, all 121 TypeScript tests, web builds, 27 UI tests, and the
  live Office HTTPS/WASM test passed on macOS. Native cross-compilation passed
  without warnings.
- `tests\run-windows.ps1 -SkipWebBuild` passed on Europa, including the connection
  storage/import regression tests, native HTTPS, Cupola ODBC policy/live query,
  real Excel formulas/spills, MSI contents, and the deployment contract.
  Credential Manager is skipped under SSH and is not covered by that result.
- The deployment wrapper returned `1618` with Excel already running and left
  the existing process running.
- The required Credential Manager round trip passed in Europa's interactive
  user session. That run used captured numeric compatibility fixtures; the
  separate Windows suite above exercised the native engine.
- The isolated developer-registration migration test passed, preserving other
  add-ins' values/types and cleaning stale manager-only Cupola registrations.
- A temporary machine registration probe demonstrated automatic COM loader
  activation with Cupola's per-user OPEN registration removed and restored.
  This is not evidence of an installed MSI's complete lifecycle.

The candidate has not been installed over the user's active developer build.
The remaining release gates above still apply; passing package inspection is
not a substitute for clean-machine installation and upgrade qualification.

The first clean pinned native rebuild exposed a generated-version mismatch:
upstream Git tags labeled the pinned engine source as a development ABI, while
the signed VGI extension requires `v1.5.5`. The recipe now explicitly selects
the same ABI label and source ID as the working driver, rejects mismatched
generated declarations, and overrides inherited version-setting variables.
The corrected clean rebuild passed all 28 native Cupola assertions. The full
Windows suite then passed using those freshly built inputs, including live
VGI extension loading/queries, real Excel formulas/spills, MSI inspection,
deployment contracts, and the isolated migration regression. Europa's local
evidence is in `artifacts/pinned-native-build.log` and
`artifacts/pinned-windows.log`; native checksums are in
`artifacts/native-inputs/provenance.json`.

## Azure approval preparation (build 20260922.6)

Azure Artifact Signing is available through `-AzureSigningConfigPath`, using the
pinned Microsoft `ArtifactSigning` module. The workflow supports GitHub OIDC,
and the adapter verifies timestamped Authenticode signatures against the expected
publisher subject. Configuration examples and the remaining account/profile/RBAC
steps are documented in [Azure signing setup](azure-signing.md).
Offline adapter tests cover failure handling without sending signing requests.
Public Trust approval, profile creation, release-identity role assignment, and a
real signing/trust test remain required before shipping a signed package.

## Persistent native sessions (build 20260925.1)

The XLL now calls the C API in the pinned Haybarn DLL in process and reuses a
serialized session per friendly connection name. It no longer starts a query
process or repeats ATTACH for every metadata query. Attachment settings and
credential changes are detected on the next request; sign-out/removal and
shutdown dispose sessions. Failed SQL is not automatically replayed.

Validation passed: `npm run check`, 121 unit tests, production web builds,
27 UI tests, live Office WASM, cross-platform native compilation, and the full
Europa Windows suite (including packed XLL formulas/spills and MSI inspection).
The new native suite verifies lifecycle/concurrency, actual attachment reuse,
empty-result schemas, numeric precision, NUL/Unicode strings, nested values,
time zones, nanoseconds, and multiple result chunks. Its latest public VGI
function timing was 1,099.6 ms cold, then 38.1/1.1/1.2 ms warm. This is attachment
reuse evidence, not a benchmark of remote data retrieval or the complete AI loop.

The package is staged, not installed over the active user add-in. Windows logs:
`artifacts/native-session-windows.log`. Live OAuth refresh/revocation and managed
deployment qualification remain release gates; this change does not establish
production readiness or complete the signing work.

## Multi-catalog update (20260925.2)

Windows profiles now reference multiple saved VGI connections and retain the
DSN-less Power Query contract. Public Weather/Earthquakes queries pass through
both native persistent sessions and the Cupola ODBC driver, including missing
members, alias collisions, and reordered defaults. The driver was rebuilt from
the pinned sources and patch with verified artifact provenance.

The profile-specific real Excel Power Query test did not complete its first
worksheet load under the noninteractive SSH session. Interactive initial-load,
Refresh All, and workbook-reopen qualification remains outstanding; successful
ODBC queries alone do not establish that Excel qualification. OAuth profiles also
need qualification against authenticated services with independent member sessions.


## Table workflow update (20260925.3)

New Windows refreshable tables use Power Query; new snapshots in both hosts are
static Excel tables. Existing managed tables retain manual refresh and can be
migrated on Windows by creating a separate Power Query copy. Their originals and
formula references remain intact until the user explicitly retires management.
The interactive Power Query qualification noted above remains outstanding; the
workflow change does not resolve Excel's first-use loading/credential prompts.
