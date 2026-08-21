param(
    [Parameter(Mandatory = $true)] [string] $HaybarnPath,
    [Parameter(Mandatory = $true)] [string] $VgiExtensionPath,
    [ValidateSet('win-x64', 'win-arm64')] [string] $Runtime = 'win-x64',
    [switch] $BuildMsi,
    [switch] $SkipWebBuild,
    [string] $CertificateThumbprint,
    [string] $TimestampUrl = 'http://timestamp.digicert.com'
)

$ErrorActionPreference = 'Stop'
$haybarnFile = Get-Item $HaybarnPath
if ($haybarnFile.Length -lt 10MB) {
    throw 'HaybarnPath appears to be a uv launcher shim. Pass the native haybarn_cli\_bin\haybarn.exe instead.'
}
$null = Get-Item $VgiExtensionPath
$repository = Split-Path -Parent $PSScriptRoot
$artifacts = Join-Path $repository 'artifacts'
$xll = Join-Path $artifacts 'xll'

function Sign-Artifact([string] $Path) {
    if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) { return }
    $signTool = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($null -eq $signTool) { throw 'signtool.exe is required when -CertificateThumbprint is supplied.' }
    & $signTool.Source sign /sha1 $CertificateThumbprint /fd SHA256 /tr $TimestampUrl /td SHA256 $Path
    if ($LASTEXITCODE -ne 0) { throw "Authenticode signing failed for $Path" }
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
dotnet build (Join-Path $PSScriptRoot 'Vgi.ExcelDna\Vgi.ExcelDna.csproj') -c Release
if ($LASTEXITCODE -ne 0) { throw 'XLL build failed.' }

$xllOutput = Join-Path $PSScriptRoot 'Vgi.ExcelDna\bin\Release\net48\publish'
Copy-Item (Join-Path $xllOutput 'Vgi.ExcelDna-packed.xll') $xll -Force
Copy-Item (Join-Path $xllOutput 'Vgi.ExcelDna64-packed.xll') $xll -Force
Copy-Item $haybarnFile.FullName (Join-Path $xll 'haybarn.exe') -Force
Copy-Item $VgiExtensionPath (Join-Path $xll 'vgi.duckdb_extension') -Force
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
foreach ($file in @('Vgi.ExcelDna-packed.xll', 'Vgi.ExcelDna64-packed.xll', 'haybarn.exe', 'vgi.duckdb_extension', 'WebView2Loader.dll', 'Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll')) {
    $path = Join-Path $xll $file
    if (Test-Path $path) { Sign-Artifact $path }
}

if ($BuildMsi) {
    if ($Runtime -ne 'win-x64') { throw 'The current WiX package targets ProgramFiles64Folder and must be built with -Runtime win-x64.' }
    dotnet build (Join-Path $repository 'installer\VgiExcel.wixproj') -c Release
    if ($LASTEXITCODE -ne 0) { throw 'MSI build failed.' }
    Sign-Artifact (Join-Path $repository 'installer\bin\Release\VgiExcel.msi')
}

$product = Get-Content (Join-Path $repository 'package.json') -Raw | ConvertFrom-Json
$releaseFiles = @(Get-ChildItem $xll -File -Recurse | Sort-Object FullName | ForEach-Object {
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
    signed = -not [string]::IsNullOrWhiteSpace($CertificateThumbprint)
    files = $releaseFiles
}
$release | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $xll 'release-manifest.json') -Encoding UTF8

Write-Host "Windows artifacts are ready under $artifacts"
Write-Host "Run 'artifacts\xll\Update Cupola for Excel.cmd' to close Excel, install the update, and reopen Excel."
