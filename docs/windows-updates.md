# Windows updates

MSI installations include an updater in **Settings → About**. Cupola checks
GitHub's latest published stable release at most once daily when the workbench
opens. Users can disable daily checks or check manually. Checks contact GitHub
with Cupola's product version; no workbook or connection data is sent.

Choose **Download update** to download and verify the installer. Choose
**Install update** to open the separate updater, then save your work and close
all Excel windows. It waits for all Excel processes to exit and allows
cancellation while waiting. It never saves workbooks or closes Excel. Windows
may request administrator approval because the MSI installs for all users.
Reopen Excel after installation; if prompted, restart Windows first.

Developer XLL installations continue to use `Update Cupola for Excel.cmd` from
the new developer package. The Microsoft 365 add-in uses its hosted web assets,
not this Windows updater.

## Release requirements

The first release containing this updater must be installed manually. Existing
installations cannot acquire the updater until that MSI is installed.

Publish a stable GitHub release in `Query-farm/vgi-excel` with a tag of
`v{version}-{cupolaBuild}`, `CupolaForExcel.msi`, and `SHA256SUMS.txt`. The existing
Windows release workflow creates these assets. Both the MSI and
`Cupola.Updater.exe` must be signed by Query Farm LLC using production signing.
Unsigned development builds cannot use the installation handoff. The signed
release CI job runs the updater’s publisher checks against both signed payloads
and checks the MSI identity before creating the draft release.

Increment the three-part product version for every MSI upgrade, as well as
`cupolaBuild`. Windows Installer does not use the build identifier to order
upgrades. A newer build with the same product version is shown with release
notes for manual installation; it is not automatically installed. MSI property
`CUPOLABUILD` is checked against the release build, and `npm run check` enforces
its agreement with `package.json`.

Downloads are restricted to the official repository asset URLs and HTTPS
GitHub download hosts, with size limits and timeouts. Before use, Cupola checks
SHA-256, Windows Authenticode trust and revocation, the Query Farm LLC publisher,
and MSI product name, manufacturer, upgrade code, version, and build. The helper
repeats verification and locks the installer against replacement while running
Windows Installer. Downgrades are rejected. The helper runs from a copy in the
per-user cache so MSI can replace the installed helper during an upgrade.

Before publishing, qualify a signed upgrade on Windows, including waiting with
unsaved workbooks, cancellation, UAC denial, installation, reopening Excel,
and `tests\excel\active-install-smoke.ps1`. A successful cross-compilation or
MSI inspection does not establish the full signed upgrade lifecycle.

## Managed deployments and local state

Set the 64-bit machine policy DWORD
`HKLM\SOFTWARE\Policies\QueryFarm\Cupola\DisableUpdates` to `1` to disable
both checks and installation through Cupola. Administrators can deploy signed
MSIs using their existing software management tools. The updater runs only
when a user requests it; it does not install a service or scheduled task.

The user's daily preference and last check time are in
`HKCU\Software\QueryFarm\Cupola\Updates`. Downloads and helper copies are
stored under `%LOCALAPPDATA%\QueryFarm\Cupola\Updates`, separated by release
tag. This cache may be removed when no updater is running. Failed checks show
an error rather than claiming the installed version is current.
