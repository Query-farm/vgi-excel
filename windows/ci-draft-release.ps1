$ErrorActionPreference = 'Stop'
$product = Get-Content package.json -Raw | ConvertFrom-Json
if ($product.version -notmatch '^\d+\.\d+\.\d+$' -or $product.cupolaBuild -notmatch '^\d{8}\.\d+$') { throw 'Unexpected product version or build format.' }
$tag = "v$($product.version)-$($product.cupolaBuild)"
$destination = Join-Path $PWD 'artifacts/release'
New-Item -ItemType Directory -Force $destination | Out-Null
Copy-Item artifacts/deployment/CupolaForExcel.msi $destination
Compress-Archive -Path artifacts/deployment/* -DestinationPath (Join-Path $destination 'CupolaForExcel-Deployment.zip')
Compress-Archive -Path artifacts/xll/* -DestinationPath (Join-Path $destination 'CupolaForExcel-Updater.zip')
Copy-Item artifacts/native-inputs/provenance.json (Join-Path $destination 'native-provenance.json')
Copy-Item artifacts/xll/release-manifest.json $destination
Get-ChildItem $destination -File | Sort-Object Name | ForEach-Object {
    "$((Get-FileHash $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant())  $($_.Name)"
} | Set-Content (Join-Path $destination 'SHA256SUMS.txt') -Encoding ASCII
$notes = @"
Cupola for Excel $($product.version), build $($product.cupolaBuild), Windows x64.

Signed as Query Farm LLC using Azure Artifact Signing and a Microsoft timestamp.
Source commit: $env:GITHUB_SHA
CI run: https://github.com/$env:GITHUB_REPOSITORY/actions/runs/$env:GITHUB_RUN_ID

## Before publishing this draft

GitHub-hosted runners do not include Excel. Validate this exact candidate on a Windows/Excel machine: real formulas and spills, Power Query refresh, interactive Credential Manager, install/upgrade/uninstall, and active-install diagnostics. Record the results before publishing. Do not rebuild the package after qualification.

- CupolaForExcel.msi: installer
- CupolaForExcel-Deployment.zip: installer and deployment scripts
- CupolaForExcel-Updater.zip: developer updater and versioned XLL payload
- SHA256SUMS.txt, release-manifest.json, native-provenance.json: checksums and build provenance
"@
$notesPath = Join-Path $env:RUNNER_TEMP 'cupola-release-notes.md'
$notes | Set-Content $notesPath -Encoding UTF8
$files = @(Get-ChildItem $destination -File | ForEach-Object FullName)
# Fail if this version already exists rather than replacing a qualified release.
gh release create $tag @files --repo $env:GITHUB_REPOSITORY --target $env:GITHUB_SHA --draft --title "Cupola for Excel $($product.version) ($($product.cupolaBuild))" --notes-file $notesPath
if ($LASTEXITCODE) { throw 'Could not create the draft release. Existing releases are never overwritten.' }
"Created signed draft $tag. Complete the Windows/Excel qualification listed in its notes before publishing." | Out-File -FilePath $env:GITHUB_STEP_SUMMARY -Append
