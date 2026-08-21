$ErrorActionPreference = 'Stop'

$options = Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
$registration = @(
    $options.PSObject.Properties |
        Where-Object { $_.Name -match '^OPEN\d*$' -and [string]$_.Value -match '(?i)Vgi\.ExcelDna(?:64)?-packed.*\.xll' } |
        Select-Object -Last 1
)
if ($registration.Count -ne 1 -or [string]$registration[0].Value -notmatch '"([^"]+\.xll)"') {
    throw 'No persistent VGI XLL registration was found.'
}
$xllPath = $Matches[1]
$null = Get-Item $xllPath

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    if (-not $excel.RegisterXLL($xllPath)) { throw "Excel rejected the registered XLL: $xllPath" }
    $diagnostics = [string]$excel.Run('VGI_DIAGNOSTICS')
    if ($diagnostics -notmatch 'Transport=HTTPS only') {
        throw "VGI_DIAGNOSTICS returned an unexpected result: $diagnostics"
    }
    Write-Host 'PASS: the persistent VGI registration points to a loadable XLL'
    Write-Host $diagnostics
}
finally {
    if ($null -ne $excel) { $excel.Quit() }
    if ($null -ne $excel -and [Runtime.InteropServices.Marshal]::IsComObject($excel)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}
