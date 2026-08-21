[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $PackagePath,
    [string] $InstallRoot,
    [ValidateSet('Auto', 'x64', 'x86')] [string] $Architecture = 'Auto',
    [string[]] $OfficeVersion,
    [switch] $RestartExcel,
    [switch] $LaunchExcel
)

$ErrorActionPreference = 'Stop'

function Get-ExcelArchitecture {
    if ($Architecture -ne 'Auto') { return $Architecture }

    $clickToRun = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($null -ne $clickToRun -and $clickToRun.Platform -eq 'x86') { return 'x86' }
    if ($null -ne $clickToRun -and $clickToRun.Platform -eq 'x64') { return 'x64' }

    if ([Environment]::Is64BitOperatingSystem) { return 'x64' }
    return 'x86'
}

function Get-OfficeVersions {
    if ($OfficeVersion.Count -gt 0) { return @($OfficeVersion) }

    $officeRoot = 'HKCU:\Software\Microsoft\Office'
    $versions = @(
        Get-ChildItem $officeRoot -ErrorAction SilentlyContinue |
            Where-Object {
                (Split-Path $_.Name -Leaf) -match '^\d+\.\d+$' -and
                (Test-Path (Join-Path $_.PSPath 'Excel'))
            } |
            ForEach-Object { Split-Path $_.Name -Leaf } |
            Sort-Object -Unique
    )
    if ($versions.Count -eq 0) { return @('16.0') }
    return $versions
}

function Test-IsVgiXllRegistration([string] $Value) {
    return $Value -match '(?i)Vgi\.ExcelDna(?:64)?-packed(?:-[^\\"]+)?\.xll'
}

function Register-VgiXll([string] $Version, [string] $XllPath) {
    $excelRoot = "HKCU:\Software\Microsoft\Office\$Version\Excel"
    $options = Join-Path $excelRoot 'Options'
    $manager = Join-Path $excelRoot 'Add-in Manager'
    New-Item -ItemType Directory -Force $options | Out-Null
    New-Item -ItemType Directory -Force $manager | Out-Null

    $optionItem = Get-ItemProperty $options
    $openProperties = @(
        $optionItem.PSObject.Properties |
            Where-Object { $_.Name -match '^OPEN\d*$' } |
            Sort-Object {
                if ($_.Name -eq 'OPEN') { 0 }
                else { 1 + [int]$_.Name.Substring(4) }
            }
    )
    $preserved = @($openProperties | Where-Object { -not (Test-IsVgiXllRegistration ([string]$_.Value)) } | ForEach-Object Value)
    foreach ($property in $openProperties) {
        Remove-ItemProperty -Path $options -Name $property.Name
    }
    $values = @($preserved) + @('/R "' + $XllPath + '"')
    for ($index = 0; $index -lt $values.Count; $index++) {
        $name = if ($index -eq 0) { 'OPEN' } else { "OPEN$index" }
        New-ItemProperty -Path $options -Name $name -PropertyType String -Value $values[$index] -Force | Out-Null
    }

    $managerItem = Get-ItemProperty $manager
    foreach ($property in $managerItem.PSObject.Properties) {
        if (Test-IsVgiXllRegistration $property.Name) {
            Remove-ItemProperty -Path $manager -Name $property.Name
        }
    }
    New-ItemProperty -Path $manager -Name $XllPath -PropertyType String -Value '' -Force | Out-Null
}

function Stop-ExcelSafely {
    $processes = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
    if ($processes.Count -eq 0) { return $false }

    $excel = $null
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    }
    catch {
        # A hidden orphaned Excel process may no longer be registered in the
        # Running Object Table. It is handled below after visible processes
        # have been protected from forced termination.
    }

    if ($null -ne $excel) {
        try {
            $unsaved = @()
            foreach ($book in @($excel.Workbooks)) {
                if (-not $book.Saved) { $unsaved += $book.Name }
            }
            if ($unsaved.Count -gt 0) {
                throw ('Excel has unsaved workbooks: ' + ($unsaved -join ', ') + '. Save them, then run the updater again.')
            }
            $excel.DisplayAlerts = $false
            $excel.Quit()
        }
        finally {
            if ([Runtime.InteropServices.Marshal]::IsComObject($excel)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
            }
        }
    }

    for ($attempt = 0; $attempt -lt 40; $attempt++) {
        if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0) { return $true }
        Start-Sleep -Milliseconds 250
    }

    $remaining = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
    $visible = @($remaining | Where-Object { $_.MainWindowHandle -ne 0 })
    if ($visible.Count -gt 0) {
        throw 'Excel did not close cleanly. Save and close its visible windows, then run the updater again.'
    }

    # At this point only windowless processes remain. They cannot present save
    # prompts or ribbon UI and would otherwise keep the previous XLL loaded.
    $remaining | Stop-Process -Force -ErrorAction SilentlyContinue
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0) { return $true }
        Start-Sleep -Milliseconds 250
    }
    throw 'A hidden Excel process could not be stopped. Restart Windows before updating Cupola.'
}

$package = (Resolve-Path $PackagePath).Path
foreach ($file in @('Vgi.ExcelDna-packed.xll', 'Vgi.ExcelDna64-packed.xll', 'haybarn.exe', 'vgi.duckdb_extension', 'WebView2Loader.dll', 'web\index.html')) {
    $null = Get-Item (Join-Path $package $file)
}

if ([string]::IsNullOrWhiteSpace($InstallRoot)) {
    $InstallRoot = Join-Path $env:LOCALAPPDATA 'QueryFarm\VgiExcel\AddIn'
}
$InstallRoot = [IO.Path]::GetFullPath($InstallRoot)
$excelWasRunning = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0
if ($excelWasRunning -and $RestartExcel) { $null = Stop-ExcelSafely }

$versionName = (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
$versionDirectory = Join-Path (Join-Path $InstallRoot 'versions') $versionName
New-Item -ItemType Directory -Force $versionDirectory | Out-Null
foreach ($file in @('Vgi.ExcelDna-packed.xll', 'Vgi.ExcelDna64-packed.xll', 'haybarn.exe', 'vgi.duckdb_extension')) {
    Copy-Item (Join-Path $package $file) (Join-Path $versionDirectory $file) -Force
}
Copy-Item (Join-Path $package 'WebView2Loader.dll') (Join-Path $versionDirectory 'WebView2Loader.dll') -Force
Copy-Item (Join-Path $package 'web') (Join-Path $versionDirectory 'web') -Recurse -Force
foreach ($assembly in @('Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll')) {
    $source = Join-Path $package $assembly
    if (Test-Path $source) { Copy-Item $source (Join-Path $versionDirectory $assembly) -Force }
}
Copy-Item $PSCommandPath (Join-Path $InstallRoot 'install-xll.ps1') -Force

# Remove files used only by the retired localhost companion. The direct HTTPS
# add-in uses desktop-connections.json and the Windows credential stores.
$configRoot = Split-Path -Parent $InstallRoot
foreach ($legacy in @('companion-token.bin', 'pairing-code.txt', 'connections.json')) {
    $legacyPath = Join-Path $configRoot $legacy
    if (Test-Path $legacyPath) { Remove-Item $legacyPath -Force }
}

$detectedArchitecture = Get-ExcelArchitecture
$xllName = if ($detectedArchitecture -eq 'x64') { 'Vgi.ExcelDna64-packed.xll' } else { 'Vgi.ExcelDna-packed.xll' }
$installedXll = Join-Path $versionDirectory $xllName
$versions = @(Get-OfficeVersions)
foreach ($version in $versions) { Register-VgiXll $version $installedXll }

Write-Host "Installed Cupola for Excel $detectedArchitecture at $versionDirectory"
Write-Host "Registered $installedXll for Excel $($versions -join ', ')"

if ($RestartExcel -or $LaunchExcel) {
    Start-Process 'excel.exe'
    Write-Host 'Excel restarted. The Cupola tab should appear on the ribbon.'
}
elseif ($excelWasRunning) {
    Write-Host 'Excel is still using the previous add-in. Close every Excel window and reopen Excel to activate this update.' -ForegroundColor Yellow
}
else {
    Write-Host 'Open Excel to activate this update.'
}
