param(
    [string] $ExpectedVersion,
    [string] $ExpectedBuild
)
$ErrorActionPreference = 'Stop'
$env:VGI_EXCEL_TELEMETRY = '0'
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { -not $_.HasExited }).Count) {
    throw 'Close Excel before verifying startup of the active Cupola installation.'
}
$product = Get-Content (Join-Path $PSScriptRoot '..\..\package.json') -Raw | ConvertFrom-Json
if (!$ExpectedVersion) { $ExpectedVersion = $product.version }
if (!$ExpectedBuild) { $ExpectedBuild = $product.cupolaBuild }
$machine = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64')
try {
    $key = $machine.OpenSubKey('SOFTWARE\QueryFarm\Cupola')
    try { $directory = if ($key) { [string]$key.GetValue('InstallDirectory') } else { $null } }
    finally { if ($key) { $key.Dispose() } }
} finally { $machine.Dispose() }
$managed = ![string]::IsNullOrWhiteSpace($directory)
if ($managed) {
    $xllPath = Join-Path $directory 'Vgi.ExcelDna64-packed.xll'
} else {
    $options = Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
    $registration = @($options.PSObject.Properties | Where-Object {
        $_.Name -match '^OPEN\d*$' -and [string]$_.Value -match '(?i)Vgi\.ExcelDna(?:64)?-packed.*\.xll'
    })
    if ($registration.Count -ne 1 -or [string]$registration[0].Value -notmatch '"([^"]+\.xll)"') {
        throw 'Expected exactly one persistent Cupola XLL registration.'
    }
    $xllPath = $Matches[1]
    $directory = Split-Path -Parent $xllPath
}
$null = Get-Item -LiteralPath $xllPath
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    if ($managed) {
        # Do not call RegisterXLL or set Connect here: verify actual machine activation.
        if (!$excel.COMAddIns.Item('QueryFarm.Cupola.ExcelLoader').Connect) {
            throw 'The installed Cupola COM loader did not activate at Excel startup.'
        }
    } elseif (!$excel.RegisterXLL($xllPath)) {
        throw 'Excel rejected the registered developer XLL.'
    }
    $diagnostics = [string]$excel.Run('VGI_DIAGNOSTICS')
    foreach ($expected in @("Product=Cupola for Excel $ExpectedVersion;", "Build=$ExpectedBuild;", 'Transport=HTTPS only;', "Engine=$(Join-Path $directory 'haybarn_odbc.dll');", 'SessionMode=Persistent native;')) {
        if (!$diagnostics.Contains($expected)) { throw "Active Cupola diagnostics did not match $expected" }
    }
    Write-Host 'PASS: active Cupola version, build, installation directory, and HTTPS transport'
    if ($managed) { Write-Host 'PASS: machine-wide loader activated automatically' }
    Write-Host $diagnostics
}
finally {
    if ($null -ne $excel) { $excel.Quit() }
    if ($null -ne $excel -and [Runtime.InteropServices.Marshal]::IsComObject($excel)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}
