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
foreach ($required in @('Vgi.ExcelDna-packed.xll', 'Vgi.ExcelDna64-packed.xll', 'haybarn.exe', 'vgi.duckdb_extension', 'WebView2Loader.dll', 'Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll', 'install-xll.ps1', 'Update Cupola for Excel.cmd', 'index.html', 'workbench.js', 'workbench.css')) {
    Assert-True ($longNames -contains $required) "MSI should contain $required"
}
Assert-True (-not ($longNames | Where-Object { $_ -match 'Companion' })) 'MSI must not contain the retired companion'
Write-Host 'PASS: MSI contents and companion exclusion'
