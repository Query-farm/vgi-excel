param(
    [Parameter(Mandatory = $true)] [string] $HaybarnPath,
    [Parameter(Mandatory = $true)] [string] $VgiExtensionPath,
    [switch] $SkipExcel,
    [switch] $SkipWebBuild
)

$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot

function Run-Step([string] $Name, [scriptblock] $Action) {
    Write-Host "`n== $Name =="
    & $Action
    if ($LASTEXITCODE -ne 0) { throw "$Name failed with exit code $LASTEXITCODE" }
}

Run-Step 'Build XLL and MSI' {
    & (Join-Path $repository 'windows\publish.ps1') -HaybarnPath $HaybarnPath -VgiExtensionPath $VgiExtensionPath -Runtime win-x64 -BuildMsi -SkipWebBuild:$SkipWebBuild
}
Run-Step 'Desktop policy and WebView tests' {
    $env:VGI_EXCEL_WEB_ASSETS_PATH = Join-Path $repository 'apps\desktop\dist'
    try { dotnet run --project (Join-Path $repository 'windows\Vgi.ExcelDna.Tests\Vgi.ExcelDna.Tests.csproj') -c Release }
    finally { Remove-Item Env:\VGI_EXCEL_WEB_ASSETS_PATH -ErrorAction SilentlyContinue }
}
Run-Step 'Native HTTPS integration' {
    & (Join-Path $repository 'tests\engine\haybarn-https.ps1') -HaybarnPath $HaybarnPath -VgiExtensionPath $VgiExtensionPath
}
if (-not $SkipExcel) {
    Run-Step 'Real Excel smoke test' {
        & (Join-Path $repository 'tests\excel\xll-smoke.ps1') -XllPath (Join-Path $repository 'artifacts\xll\Vgi.ExcelDna64-packed.xll')
    }
}
Run-Step 'MSI inspection' {
    & (Join-Path $repository 'tests\packaging\msi-smoke.ps1') -MsiPath (Join-Path $repository 'installer\bin\Release\VgiExcel.msi')
}

Write-Host "`nPASS: all Windows Cupola for Excel tests"
