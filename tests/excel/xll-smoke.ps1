param(
    [Parameter(Mandatory = $true)] [string] $XllPath,
    [string] $Location = 'https://vgi-open-meteo.rusty-bb6.workers.dev',
    [string] $Catalog = 'open_meteo'
)

$ErrorActionPreference = 'Stop'

function Assert-Equal($Expected, $Actual, [string] $Message) {
    if ($Expected -ne $Actual) { throw "Assertion failed: $Message. Expected '$Expected', got '$Actual'." }
}
function Assert-True([bool] $Value, [string] $Message) {
    if (-not $Value) { throw "Assertion failed: $Message" }
}
function Test-IsVgiXllRegistration([string] $Value) {
    return $Value -match '(?i)Vgi\.ExcelDna(?:64)?-packed(?:-[^\\"]+)?\.xll'
}
function Save-ExcelRegistrations {
    $snapshots = @()
    foreach ($versionKey in @(Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue | Where-Object { (Split-Path $_.Name -Leaf) -match '^\d+\.\d+$' })) {
        $excelKey = Join-Path $versionKey.PSPath 'Excel'
        if (-not (Test-Path $excelKey)) { continue }
        $options = Join-Path $excelKey 'Options'
        $manager = Join-Path $excelKey 'Add-in Manager'
        $openValues = @()
        if (Test-Path $options) {
            $openValues = @(
                (Get-ItemProperty $options).PSObject.Properties |
                    Where-Object { $_.Name -match '^OPEN\d*$' } |
                    ForEach-Object { [PSCustomObject]@{ Name = $_.Name; Value = [string]$_.Value } }
            )
        }
        $managerValues = @()
        if (Test-Path $manager) {
            $managerValues = @(
                (Get-ItemProperty $manager).PSObject.Properties |
                    Where-Object { Test-IsVgiXllRegistration $_.Name } |
                    ForEach-Object { [PSCustomObject]@{ Name = $_.Name; Value = [string]$_.Value } }
            )
        }
        $snapshots += [PSCustomObject]@{
            Options = $options
            Manager = $manager
            OpenValues = $openValues
            ManagerValues = $managerValues
        }
    }
    return @($snapshots)
}
function Restore-ExcelRegistrations($Snapshots) {
    foreach ($snapshot in @($Snapshots)) {
        New-Item -ItemType Directory -Force $snapshot.Options | Out-Null
        foreach ($property in @((Get-ItemProperty $snapshot.Options).PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' })) {
            Remove-ItemProperty -Path $snapshot.Options -Name $property.Name
        }
        foreach ($property in @($snapshot.OpenValues)) {
            New-ItemProperty -Path $snapshot.Options -Name $property.Name -PropertyType String -Value $property.Value -Force | Out-Null
        }

        New-Item -ItemType Directory -Force $snapshot.Manager | Out-Null
        foreach ($property in @((Get-ItemProperty $snapshot.Manager).PSObject.Properties | Where-Object { Test-IsVgiXllRegistration $_.Name })) {
            Remove-ItemProperty -Path $snapshot.Manager -Name $property.Name
        }
        foreach ($property in @($snapshot.ManagerValues)) {
            New-ItemProperty -Path $snapshot.Manager -Name $property.Name -PropertyType String -Value $property.Value -Force | Out-Null
        }
    }
}
function Disable-VgiRegistrations($Snapshots) {
    foreach ($snapshot in @($Snapshots)) {
        if (Test-Path $snapshot.Options) {
            $preserved = @($snapshot.OpenValues | Where-Object { -not (Test-IsVgiXllRegistration $_.Value) } | ForEach-Object Value)
            foreach ($property in @((Get-ItemProperty $snapshot.Options).PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' })) {
                Remove-ItemProperty -Path $snapshot.Options -Name $property.Name
            }
            for ($index = 0; $index -lt $preserved.Count; $index++) {
                $name = if ($index -eq 0) { 'OPEN' } else { "OPEN$index" }
                New-ItemProperty -Path $snapshot.Options -Name $name -PropertyType String -Value $preserved[$index] -Force | Out-Null
            }
        }
        if (Test-Path $snapshot.Manager) {
            foreach ($property in @((Get-ItemProperty $snapshot.Manager).PSObject.Properties | Where-Object { Test-IsVgiXllRegistration $_.Name })) {
                Remove-ItemProperty -Path $snapshot.Manager -Name $property.Name
            }
        }
    }
}

$xll = (Resolve-Path $XllPath).Path
$directory = Split-Path -Parent $xll
$null = Get-Item (Join-Path $directory 'haybarn.exe')
$null = Get-Item (Join-Path $directory 'vgi.duckdb_extension')
$root = Join-Path $env:LOCALAPPDATA 'QueryFarm\VgiExcel'
$registry = Join-Path $root 'desktop-connections.json'
$defaultFile = Join-Path $root 'default-connection.txt'
New-Item -ItemType Directory -Force $root | Out-Null
$hadRegistry = Test-Path $registry
$registryBackup = if ($hadRegistry) { [PSCustomObject]@{ Bytes = [IO.File]::ReadAllBytes($registry) } } else { $null }
$hadDefault = Test-Path $defaultFile
$defaultBackup = if ($hadDefault) { [PSCustomObject]@{ Bytes = [IO.File]::ReadAllBytes($defaultFile) } } else { $null }
$existingProcesses = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$excelRegistrations = @(Save-ExcelRegistrations)
Disable-VgiRegistrations $excelRegistrations
$testName = '_vgi_excel_smoke_' + [Guid]::NewGuid().ToString('N')
$badName = $testName + '_http'
$excel = $null
$book = $null
$sheet = $null

try {
    $connections = @()
    if ($hadRegistry) { $connections = @((Get-Content $registry -Raw | ConvertFrom-Json)) }
    $connections = @($connections | Where-Object { $_.Name -ne $testName -and $_.Name -ne $badName })
    $connections += [PSCustomObject]@{ Name = $testName; Catalog = $Catalog; Location = $Location; Authentication = 'anonymous' }
    $connections += [PSCustomObject]@{ Name = $badName; Catalog = 'bad'; Location = 'http://vgi.example.com'; Authentication = 'anonymous' }
    ConvertTo-Json -InputObject @($connections) -Depth 8 | Set-Content -Encoding UTF8 $registry
    Set-Content -Encoding ASCII $defaultFile $testName

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    Assert-True ([bool] $excel.RegisterXLL($xll)) 'Excel should register the packed XLL'

    $sheet.Range('A1').Formula = '=VGI_VALUE("SELECT open_meteo.main.weather_code_text(0)","' + $testName + '",0)'
    $sheet.Range('A2').Formula = '=VGI_CALL("open_meteo.main.weather_code_text",61)'
    $sheet.Range('D1').Formula2 = '=VGI_QUERY("SELECT name,country FROM open_meteo.main.geocoding(''Boston'', count := 1::BIGINT)","' + $testName + '",TRUE,0)'
    $sheet.Range('G1').Formula2 = '=VGI_QUERY("SELECT 99.99::DECIMAL(18,2) AS amount, ''001200''::VARCHAR AS account_code","' + $testName + '",TRUE,0)'
    $excel.CalculateFull()

    Assert-Equal 'Clear sky' $sheet.Range('A1').Value2 'VGI_VALUE result'
    Assert-Equal 'Slight rain' $sheet.Range('A2').Value2 'VGI_CALL result'
    Assert-Equal 'name' $sheet.Range('D1').Value2 'VGI_QUERY first header'
    Assert-Equal 'country' $sheet.Range('E1').Value2 'VGI_QUERY second header'
    Assert-Equal 'Boston' $sheet.Range('D2').Value2 'VGI_QUERY first value'
    Assert-Equal 'United States' $sheet.Range('E2').Value2 'VGI_QUERY second value'
    Assert-Equal 99.99 $sheet.Range('G2').Value2 'safe accounting decimal remains numeric in Excel'
    Assert-Equal '001200' $sheet.Range('H2').Value2 'leading-zero account code remains text in Excel'
    Assert-Equal 'No VGI XLL error has been recorded.' $excel.Run('VGI_LAST_ERROR') 'clean diagnostic before negative test'

    $badResult = $excel.Run('VGI_VALUE', 'SELECT 1', $badName, 0)
    $lastError = [string] $excel.Run('VGI_LAST_ERROR')
    Assert-True ($badResult -ne 1) 'an insecure connection must not execute'
    Assert-True ($lastError -match 'HTTPS VGI endpoints only') 'the insecure connection should return a useful diagnostic'
    Assert-True ([string]$excel.Run('VGI_DIAGNOSTICS') -match 'Transport=HTTPS only') 'diagnostics should report HTTPS-only transport'

    Write-Host 'PASS: real Excel XLL formulas, spill results, and HTTPS rejection'
}
finally {
    if ($book -ne $null) { $book.Close($false) }
    if ($excel -ne $null) { $excel.Quit() }
    foreach ($item in @($sheet, $book, $excel)) {
        if ($item -ne $null -and [Runtime.InteropServices.Marshal]::IsComObject($item)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Restore-ExcelRegistrations $excelRegistrations
    if ($hadRegistry) { [IO.File]::WriteAllBytes($registry, $registryBackup.Bytes) }
    elseif (Test-Path $registry) { Remove-Item -Force $registry }
    if ($hadDefault) { [IO.File]::WriteAllBytes($defaultFile, $defaultBackup.Bytes) }
    elseif (Test-Path $defaultFile) { Remove-Item -Force $defaultFile }
    $newProcesses = @()
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        Start-Sleep -Milliseconds 500
        $newProcesses = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $existingProcesses -notcontains $_.Id })
        if ($newProcesses.Count -eq 0) { break }
    }
    Assert-True ($newProcesses.Count -eq 0) 'the Excel instance created by the smoke test should exit'
}
