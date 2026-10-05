param(
    [Parameter(Mandatory = $true)] [string] $HaybarnPath,
    [Parameter(Mandatory = $true)] [string] $VgiExtensionPath,
    [string] $OdbcDriverPath = $env:CUPOLA_ODBC_DRIVER_PATH,
    [switch] $SkipExcel,
    [switch] $SkipWebBuild
)

$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$env:VGI_EXCEL_TELEMETRY = '0'

function Run-Step([string] $Name, [scriptblock] $Action) {
    Write-Host "`n== $Name =="
    & $Action
    if ($LASTEXITCODE -ne 0) { throw "$Name failed with exit code $LASTEXITCODE" }
}

Run-Step 'Build XLL and MSI' {
    & (Join-Path $repository 'windows\publish.ps1') -HaybarnPath $HaybarnPath -VgiExtensionPath $VgiExtensionPath -OdbcDriverPath $OdbcDriverPath -Runtime win-x64 -BuildMsi -SkipWebBuild:$SkipWebBuild
}
Run-Step 'Desktop policy and WebView tests' {
    $previousWebAssets = $env:VGI_EXCEL_WEB_ASSETS_PATH
    $previousHaybarn = $env:VGI_HAYBARN_PATH
    $env:VGI_EXCEL_WEB_ASSETS_PATH = Join-Path $repository 'apps\desktop\dist'
    $env:VGI_HAYBARN_PATH = (Resolve-Path $HaybarnPath).Path
    try { dotnet run --project (Join-Path $repository 'windows\Vgi.ExcelDna.Tests\Vgi.ExcelDna.Tests.csproj') -c Release }
    finally {
        $env:VGI_EXCEL_WEB_ASSETS_PATH = $previousWebAssets
        $env:VGI_HAYBARN_PATH = $previousHaybarn
    }
}
Run-Step 'Persistent native sessions' {
    $previousNative = $env:VGI_HAYBARN_NATIVE_PATH
    $previousExtension = $env:VGI_EXTENSION_PATH
    $env:VGI_HAYBARN_NATIVE_PATH = (Resolve-Path $OdbcDriverPath).Path
    $env:VGI_EXTENSION_PATH = (Resolve-Path $VgiExtensionPath).Path
    try { dotnet run --project (Join-Path $repository 'windows\Vgi.ExcelDna.Tests\Vgi.ExcelDna.Tests.csproj') -c Release -- --native-sessions
        if ($LASTEXITCODE -ne 0) { throw "Native sessions failed" }
        if (-not $SkipExcel) { dotnet run --project (Join-Path $repository 'windows\Vgi.ExcelDna.Tests\Vgi.ExcelDna.Tests.csproj') -c Release -- --workbook-tables }
    }
    finally {
        $env:VGI_HAYBARN_NATIVE_PATH = $previousNative
        $env:VGI_EXTENSION_PATH = $previousExtension
    }
}
Run-Step 'Native HTTPS integration' {
    & (Join-Path $repository 'tests\engine\haybarn-https.ps1') -HaybarnPath $HaybarnPath -VgiExtensionPath $VgiExtensionPath
}
Run-Step 'Cupola ODBC connection integration' {
    & powershell.exe -NoProfile -File (Join-Path $repository 'tests\odbc\cupola-connection.ps1') -DriverPath $OdbcDriverPath -VgiExtensionPath $VgiExtensionPath
}
if (-not $SkipExcel) {
    Run-Step 'Real Excel smoke test' {
        & (Join-Path $repository 'tests\excel\xll-smoke.ps1') -XllPath (Join-Path $repository 'artifacts\xll\Vgi.ExcelDna64-packed.xll')
    }
}
Run-Step 'MSI inspection' {
    & (Join-Path $repository 'tests\packaging\msi-smoke.ps1') -MsiPath (Join-Path $repository 'installer\bin\Release\VgiExcel.msi')
}

Run-Step 'Updater MSI identity' {
    dotnet run --project (Join-Path $repository 'windows\Vgi.ExcelDna.Tests\Vgi.ExcelDna.Tests.csproj') -c Release -- --verify-update-identity (Join-Path $repository 'installer\bin\Release\VgiExcel.msi')
}

Run-Step 'Connection defaults validation' {
    & (Join-Path $repository 'tests\packaging\connection-defaults.ps1')
}
Run-Step 'Enterprise deployment contract' {
    & (Join-Path $repository 'tests\packaging\deployment-contract.ps1') -MsiPath (Join-Path $repository 'installer\bin\Release\VgiExcel.msi')
}
Run-Step 'Developer registration migration' {
    & (Join-Path $repository 'tests\packaging\migration-smoke.ps1')
}

Run-Step 'Azure signing policy (offline)' {
    & (Join-Path $repository 'tests\packaging\azure-signing-smoke.ps1')
}

Write-Host "`nPASS: all Windows Cupola for Excel tests"
