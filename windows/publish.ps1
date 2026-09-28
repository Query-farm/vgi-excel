param(
    [Parameter(Mandatory = $true)] [string] $HaybarnPath,
    [Parameter(Mandatory = $true)] [string] $VgiExtensionPath,
    [ValidateSet('win-x64', 'win-arm64')] [string] $Runtime = 'win-x64',
    [string] $OdbcDriverPath = $env:CUPOLA_ODBC_DRIVER_PATH,
    [switch] $BuildMsi,
    [switch] $Production,
    [string] $NativeProvenancePath,
    [switch] $SkipWebBuild,
    [string] $CertificateThumbprint,
    [string] $AzureSigningConfigPath,
    [string] $TimestampUrl = 'http://timestamp.digicert.com'
)

$ErrorActionPreference = 'Stop'
$signingEnabled = ![string]::IsNullOrWhiteSpace($CertificateThumbprint) -or ![string]::IsNullOrWhiteSpace($AzureSigningConfigPath)
if ($CertificateThumbprint -and $AzureSigningConfigPath) { throw 'Choose either certificate-store signing or Azure signing.' }
if ($Production -and (!$signingEnabled -or !$BuildMsi)) {
    throw 'Production publishing requires -BuildMsi and either -CertificateThumbprint or -AzureSigningConfigPath.'
}
$azureSigning = $null
if ($AzureSigningConfigPath) {
    . (Join-Path $PSScriptRoot 'signing\azure.ps1')
    $azureSigning = Read-CupolaAzureSigningConfig $AzureSigningConfigPath
    Import-Module ArtifactSigning -RequiredVersion 0.1.20 -ErrorAction Stop
    $null = Get-Command az -ErrorAction Stop
}
if ($Production -and $Runtime -ne 'win-x64') { throw 'Production currently supports x64 Windows and x64 Excel only.' }
$haybarnFile = Get-Item $HaybarnPath
if ($haybarnFile.Length -lt 10MB) {
    throw 'HaybarnPath appears to be a uv launcher shim. Pass the native haybarn_cli\_bin\haybarn.exe instead.'
}
$null = Get-Item $VgiExtensionPath
if ([string]::IsNullOrWhiteSpace($OdbcDriverPath)) { throw 'Pass -OdbcDriverPath for the Cupola-compatible Haybarn ODBC DLL.' }
$null = Get-Item $OdbcDriverPath
$repository = Split-Path -Parent $PSScriptRoot
$product = Get-Content (Join-Path $repository 'package.json') -Raw | ConvertFrom-Json
if ($Production) {
    if ([string]::IsNullOrWhiteSpace($NativeProvenancePath)) { $NativeProvenancePath = Join-Path (Split-Path -Parent $OdbcDriverPath) 'provenance.json' }
    $provenance = Get-Content -LiteralPath $NativeProvenancePath -Raw | ConvertFrom-Json
    if ($provenance.lockSha256 -ne (Get-FileHash (Join-Path $PSScriptRoot 'native-inputs.lock.json')).Hash -or $provenance.patchSha256 -ne (Get-FileHash (Join-Path $PSScriptRoot 'odbc\cupola.patch')).Hash) { throw 'Native provenance does not match the reviewed source lock and Cupola patch.' }
    foreach ($input in @(@{Path=$HaybarnPath;Name='haybarn.exe'},@{Path=$OdbcDriverPath;Name='haybarn_odbc.dll'},@{Path=$VgiExtensionPath;Name='vgi.duckdb_extension'})) {
        if ($provenance.files.($input.Name) -ne (Get-FileHash -LiteralPath $input.Path -Algorithm SHA256).Hash) { throw 'A native input does not match its build provenance.' }
    }
}
$artifacts = Join-Path $repository 'artifacts'
$xll = Join-Path $artifacts 'xll'

function Sign-Artifact([string] $Path) {
    if ($azureSigning) { Invoke-CupolaAzureSign $Path $azureSigning; return }
    if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) { return }
    $certificate = @(Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My | Where-Object Thumbprint -eq $CertificateThumbprint | Where-Object HasPrivateKey | Select-Object -First 1)
    if ($certificate.Count -ne 1) { throw 'The signing certificate with its private key must be available in the Windows certificate store.' }
    if ([IO.Path]::GetExtension($Path) -eq '.ps1') {
        $null = Set-AuthenticodeSignature -FilePath $Path -Certificate $certificate[0] -HashAlgorithm SHA256 -TimestampServer $TimestampUrl
    }
    else {
        $signTool = Get-Command signtool.exe -ErrorAction SilentlyContinue
        if ($null -eq $signTool) { throw 'signtool.exe is required when -CertificateThumbprint is supplied.' }
        $storeArguments = @()
        if ($certificate[0].PSParentPath -match 'LocalMachine') { $storeArguments += '/sm' }
        & $signTool.Source sign @storeArguments /sha1 $CertificateThumbprint /fd SHA256 /tr $TimestampUrl /td SHA256 $Path
        if ($LASTEXITCODE -ne 0) { throw "Authenticode signing failed for $Path" }
    }
    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    if ($signature.Status -ne 'Valid' -or $null -eq $signature.TimeStamperCertificate -or $signature.SignerCertificate.Thumbprint -ne $CertificateThumbprint) { throw "Signature, publisher, or timestamp validation failed for $Path" }
}

New-Item -ItemType Directory -Force -Path $xll | Out-Null
if (-not $SkipWebBuild) {
    Push-Location $repository
    try {
        npm run build:desktop
        if ($LASTEXITCODE -ne 0) { throw 'Desktop Workbench web build failed.' }
    }
    finally { Pop-Location }
}
$null = Get-Item (Join-Path $repository 'apps\desktop\dist\index.html')
$webBuild = Get-Content (Join-Path $repository 'apps\desktop\dist\build-info.json') -Raw | ConvertFrom-Json
if ($webBuild.version -ne $product.version -or $webBuild.build -ne $product.cupolaBuild -or $webBuild.host -ne 'desktop') { throw 'The desktop bundle does not match this release. Rebuild and copy the current dist directory.' }
if (Get-ChildItem (Join-Path $repository 'apps\desktop\dist') -Recurse -File -Filter '*.map') {
    throw 'Desktop web bundle contains source maps. Copy a clean production dist directory before packaging.'
}
dotnet build (Join-Path $PSScriptRoot 'Vgi.ExcelDna\Vgi.ExcelDna.csproj') -c Release
if ($LASTEXITCODE -ne 0) { throw 'XLL build failed.' }

dotnet build (Join-Path $PSScriptRoot 'Cupola.ExcelLoader\Cupola.ExcelLoader.csproj') -c Release
if ($LASTEXITCODE -ne 0) { throw 'Excel loader build failed.' }
Copy-Item (Join-Path $PSScriptRoot 'Cupola.ExcelLoader\bin\Release\net48\Cupola.ExcelLoader.dll') $xll -Force

$xllOutput = Join-Path $PSScriptRoot 'Vgi.ExcelDna\bin\Release\net48\publish'
Copy-Item (Join-Path $xllOutput 'Vgi.ExcelDna-packed.xll') $xll -Force
Copy-Item (Join-Path $xllOutput 'Vgi.ExcelDna64-packed.xll') $xll -Force
Copy-Item $haybarnFile.FullName (Join-Path $xll 'haybarn.exe') -Force
Copy-Item $VgiExtensionPath (Join-Path $xll 'vgi.duckdb_extension') -Force
Copy-Item $OdbcDriverPath (Join-Path $xll 'haybarn_odbc.dll') -Force
Copy-Item (Join-Path $PSScriptRoot 'ODBC-NOTICES.txt') $xll -Force
Copy-Item (Join-Path $PSScriptRoot 'register-odbc.ps1') (Join-Path $xll 'register-odbc.ps1') -Force
Copy-Item (Join-Path $PSScriptRoot 'install-xll.ps1') (Join-Path $xll 'install-xll.ps1') -Force
Copy-Item (Join-Path $PSScriptRoot 'update-xll.cmd') (Join-Path $xll 'Update Cupola for Excel.cmd') -Force
$web = Join-Path $xll 'web'
if (Test-Path $web) { Remove-Item $web -Recurse -Force }
Copy-Item (Join-Path $repository 'apps\desktop\dist') $web -Recurse
$loader = Join-Path $PSScriptRoot "Vgi.ExcelDna\bin\Release\net48\runtimes\$Runtime\native\WebView2Loader.dll"
Copy-Item $loader (Join-Path $xll 'WebView2Loader.dll') -Force
foreach ($assembly in @('Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll')) {
    $source = Join-Path $xllOutput $assembly
    if (-not (Test-Path $source)) { $source = Join-Path (Split-Path -Parent $xllOutput) $assembly }
    if (Test-Path $source) { Copy-Item $source (Join-Path $xll $assembly) -Force }
    else { throw "XLL build did not produce required WebView2 assembly $assembly" }
}
foreach ($file in @('Vgi.ExcelDna-packed.xll', 'Vgi.ExcelDna64-packed.xll', 'haybarn.exe', 'haybarn_odbc.dll', 'Cupola.ExcelLoader.dll', 'WebView2Loader.dll', 'Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll', 'install-xll.ps1', 'register-odbc.ps1')) {
    $path = Join-Path $xll $file
    if (Test-Path $path) { Sign-Artifact $path }
}

# The VGI extension carries a Haybarn signature at EOF. Authenticode modification
# would invalidate it; distribute exactly the reviewed, engine-signed bytes.
if ((Get-FileHash (Join-Path $xll 'vgi.duckdb_extension')).Hash -ne (Get-FileHash $VgiExtensionPath).Hash) { throw 'The packaged VGI extension was modified.' }

if (Get-ChildItem $xll -Recurse -File -Filter '*.map') {
    throw 'The staged updater contains source maps. Remove stale build output before packaging.'
}

if ($BuildMsi) {
    cmake -S (Join-Path $repository 'installer\checks') -B (Join-Path $artifacts 'installer-checks') -A x64
    if ($LASTEXITCODE -ne 0) { throw 'Installer checks configuration failed.' }
    cmake --build (Join-Path $artifacts 'installer-checks') --config Release
    if ($LASTEXITCODE -ne 0) { throw 'Installer checks build failed.' }
    Sign-Artifact (Join-Path $artifacts 'installer-checks\Release\CupolaInstallerChecks.dll')
    if ($Runtime -ne 'win-x64') { throw 'The current WiX package targets ProgramFiles64Folder and must be built with -Runtime win-x64.' }
    dotnet build (Join-Path $repository 'installer\VgiExcel.wixproj') -c Release
    if ($LASTEXITCODE -ne 0) { throw 'MSI build failed.' }
    $msiPath = Join-Path $repository 'installer\bin\Release\VgiExcel.msi'
    Sign-Artifact $msiPath
    $deployment = Join-Path $artifacts 'deployment'
    New-Item -ItemType Directory -Force $deployment | Out-Null
    Copy-Item $msiPath (Join-Path $deployment 'CupolaForExcel.msi') -Force
    Copy-Item (Join-Path $PSScriptRoot 'deployment\*.ps1') $deployment -Force
    foreach ($script in @(Get-ChildItem $deployment -Filter '*.ps1')) { Sign-Artifact $script.FullName }
    $product = Get-Content (Join-Path $repository 'package.json') -Raw | ConvertFrom-Json
    [ordered]@{
        version = $product.version
        build = $product.cupolaBuild
        installContext = 'System'
        installCommand = 'powershell.exe -NoProfile -File Deploy-Cupola.ps1 -Action Install'
        uninstallCommand = 'powershell.exe -NoProfile -File Deploy-Cupola.ps1 -Action Uninstall'
        detectionCommand = "powershell.exe -NoProfile -File Detect-Cupola.ps1 -MinimumVersion $($product.version)"
        successCodes = @(0)
        rebootRequiredCodes = @(3010)
        retryCodes = @(1618)
        msiSha256 = (Get-FileHash $msiPath -Algorithm SHA256).Hash.ToLowerInvariant()
        signed = $signingEnabled
    } | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $deployment 'deployment.json') -Encoding UTF8
}

$product = Get-Content (Join-Path $repository 'package.json') -Raw | ConvertFrom-Json
$releaseFiles = @(Get-ChildItem $xll -File -Recurse | Where-Object { $_.FullName -ne (Join-Path $xll 'release-manifest.json') } | Sort-Object FullName | ForEach-Object {
    [ordered]@{
        path = $_.FullName.Substring($xll.Length + 1).Replace('\', '/')
        bytes = $_.Length
        sha256 = (Get-FileHash $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
    }
})
$release = [ordered]@{
    product = 'Cupola for Excel'
    version = $product.version
    build = $product.cupolaBuild
    runtime = $Runtime
    createdUtc = [DateTime]::UtcNow.ToString('O')
    signed = $signingEnabled
    files = $releaseFiles
}
$release | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $xll 'release-manifest.json') -Encoding UTF8

Write-Host "Windows artifacts are ready under $artifacts"
Write-Host "Run 'artifacts\xll\Update Cupola for Excel.cmd' to close Excel, install the update, and reopen Excel."
