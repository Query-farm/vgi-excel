param([Parameter(Mandatory = $true)] [string] $MsiPath)

$ErrorActionPreference = 'Stop'

function Assert-True([bool] $Value, [string] $Message) {
    if (-not $Value) { throw "Assertion failed: $Message" }
}
function Invoke-Com($Object, [string] $Name, [string] $Kind, [object[]] $Arguments) {
    return $Object.GetType().InvokeMember($Name, $Kind, $null, $Object, $Arguments)
}

$msi = Get-Item (Resolve-Path $MsiPath)
Assert-True ($msi.Length -gt 10MB) 'the MSI should contain Haybarn and the VGI extension'
$installer = New-Object -ComObject WindowsInstaller.Installer
$database = Invoke-Com $installer 'OpenDatabase' 'InvokeMethod' @($msi.FullName, 0)
$view = Invoke-Com $database 'OpenView' 'InvokeMethod' @('SELECT `FileName` FROM `File`')
$null = Invoke-Com $view 'Execute' 'InvokeMethod' @()
$names = @()
while ($true) {
    $record = Invoke-Com $view 'Fetch' 'InvokeMethod' @()
    if ($null -eq $record) { break }
    $names += [string](Invoke-Com $record 'StringData' 'GetProperty' @(1))
}
$longNames = @($names | ForEach-Object { if ($_ -match '\|') { ($_ -split '\|', 2)[1] } else { $_ } })
foreach ($required in @('Vgi.ExcelDna64-packed.xll', 'Cupola.ExcelLoader.dll', 'Cupola.Updater.exe', 'cupola-mark.svg', 'build-info.json', 'haybarn.exe', 'haybarn_odbc.dll', 'ODBC-NOTICES.txt', 'vgi.duckdb_extension', 'WebView2Loader.dll', 'Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll', 'index.html', 'results.html', 'results.js', 'workbench.js', 'workbench.css')) {
    Assert-True ($longNames -contains $required) "MSI should contain $required"
}
Assert-True (-not ($longNames | Where-Object { $_ -match 'Companion' })) 'MSI must not contain the retired companion'
Assert-True (-not ($longNames | Where-Object { $_ -match '\.map$' })) 'MSI must not contain source maps'
$registryView = Invoke-Com $database 'OpenView' 'InvokeMethod' @('SELECT `Root`, `Key`, `Name`, `Value` FROM `Registry`')
$null = Invoke-Com $registryView 'Execute' 'InvokeMethod' @()
$driverRegistered = $false
$driverListed = $false
while ($true) {
    $record = Invoke-Com $registryView 'Fetch' 'InvokeMethod' @()
    if ($null -eq $record) { break }
    $root = [string](Invoke-Com $record 'StringData' 'GetProperty' @(1))
    $key = [string](Invoke-Com $record 'StringData' 'GetProperty' @(2))
    $name = [string](Invoke-Com $record 'StringData' 'GetProperty' @(3))
    $value = [string](Invoke-Com $record 'StringData' 'GetProperty' @(4))
    if ($root -eq '2' -and $key -eq 'SOFTWARE\ODBC\ODBCINST.INI\Cupola for Excel' -and $name -eq 'Driver' -and $value -eq '[#CupolaOdbcDll]') { $driverRegistered = $true }
    if ($root -eq '2' -and $key -eq 'SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers' -and $name -eq 'Cupola for Excel' -and $value -eq 'Installed') { $driverListed = $true }
}
Assert-True ($driverRegistered -and $driverListed) 'MSI should register its own Cupola ODBC DLL machine-wide'

$product = Get-Content (Join-Path $PSScriptRoot '..\..\package.json') -Raw | ConvertFrom-Json
$buildView = Invoke-Com $database 'OpenView' 'InvokeMethod' @('SELECT `Value` FROM `Property` WHERE `Property` = ''CUPOLABUILD''')
$null = Invoke-Com $buildView 'Execute' 'InvokeMethod' @()
$buildRecord = Invoke-Com $buildView 'Fetch' 'InvokeMethod' @()
Assert-True ($null -ne $buildRecord) 'MSI must identify the build for update verification'
Assert-True ((Invoke-Com $buildRecord 'StringData' 'GetProperty' @(1)) -eq $product.cupolaBuild) 'MSI build must match the release'
Write-Host 'PASS: MSI contents, updater build identity, ODBC registration, companion exclusion, and source-map exclusion'
