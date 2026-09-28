[CmdletBinding()]
param([Parameter(Mandatory=$true)][version] $MinimumVersion)
$ErrorActionPreference = 'Stop'
try {
    $machine = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
    try {
        $product = $machine.OpenSubKey('SOFTWARE\QueryFarm\Cupola')
        if (!$product) { exit 1 }
        try { $version = [version]$product.GetValue('Version'); $directory = [string]$product.GetValue('InstallDirectory') } finally { $product.Dispose() }
        if ($version -lt $MinimumVersion -or !$directory) { exit 1 }
        foreach ($file in @('Cupola.ExcelLoader.dll','Vgi.ExcelDna64-packed.xll','haybarn.exe','haybarn_odbc.dll','vgi.duckdb_extension','WebView2Loader.dll','web\index.html','web\cupola-mark.svg')) {
            if (!(Test-Path -LiteralPath (Join-Path $directory $file) -PathType Leaf)) { exit 1 }
        }
        $odbc = $machine.OpenSubKey('SOFTWARE\ODBC\ODBCINST.INI\Cupola for Excel')
        if (!$odbc) { exit 1 }
        try { if ([string]$odbc.GetValue('Driver') -ne (Join-Path $directory 'haybarn_odbc.dll')) { exit 1 } } finally { $odbc.Dispose() }
        $loader = $machine.OpenSubKey('SOFTWARE\Microsoft\Office\Excel\Addins\QueryFarm.Cupola.ExcelLoader')
        if (!$loader) { exit 1 }
        try { if ($loader.GetValue('LoadBehavior') -ne 3) { exit 1 } } finally { $loader.Dispose() }
        Write-Output "Cupola for Excel $version is installed."
        exit 0
    } finally { $machine.Dispose() }
} catch { exit 1 }
