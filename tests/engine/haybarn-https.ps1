param(
    [Parameter(Mandatory = $true)] [string] $HaybarnPath,
    [Parameter(Mandatory = $true)] [string] $VgiExtensionPath,
    [string] $Location = 'https://vgi-open-meteo.rusty-bb6.workers.dev',
    [string] $Catalog = 'open_meteo'
)

$ErrorActionPreference = 'Stop'

function Assert-True([bool] $Value, [string] $Message) {
    if (-not $Value) { throw "Assertion failed: $Message" }
}

function Assert-HttpsLocation([string] $Value) {
    $uri = $null
    if (-not [Uri]::TryCreate($Value, [UriKind]::Absolute, [ref] $uri) -or $uri.Scheme -ne 'https') {
        throw 'Cupola for Excel supports HTTPS VGI endpoints only.'
    }
    if (-not [string]::IsNullOrEmpty($uri.UserInfo)) { throw 'Credentials must not be embedded in the VGI URL.' }
}

function Sql-String([string] $Value) { return "'$($Value.Replace("'", "''"))'" }
function Sql-Identifier([string] $Value) { return '"' + $Value.Replace('"', '""') + '"' }

function Invoke-Haybarn([string] $Sql) {
    Assert-HttpsLocation $Location
    $start = New-Object System.Diagnostics.ProcessStartInfo
    $start.FileName = (Resolve-Path $HaybarnPath).Path
    $start.Arguments = '-json'
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardInput = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    Assert-True (-not $start.Arguments.Contains($Location)) 'the VGI URL must not be present in process arguments'
    $script = "LOAD $(Sql-String (Resolve-Path $VgiExtensionPath).Path);`n" +
        "ATTACH $(Sql-String $Catalog) AS $(Sql-Identifier $Catalog) (TYPE vgi, LOCATION $(Sql-String $Location));`n$Sql`n"
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $start
    if (-not $process.Start()) { throw 'Unable to start Haybarn.' }
    try {
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $process.StandardInput.Write($script)
        $process.StandardInput.Close()
        if (-not $process.WaitForExit(300000)) { $process.Kill(); throw 'Haybarn query timed out.' }
        $output = $stdout.GetAwaiter().GetResult()
        $errorText = $stderr.GetAwaiter().GetResult()
        if ($process.ExitCode -ne 0) { throw "Haybarn failed: $errorText" }
        $startIndex = $output.LastIndexOf("`n[")
        $json = if ($startIndex -ge 0) { $output.Substring($startIndex + 1) } else { $output.Trim() }
        return @($json | ConvertFrom-Json)
    }
    finally { $process.Dispose() }
}

$null = Get-Item $HaybarnPath
$extension = Get-Item $VgiExtensionPath
Assert-True ($extension.Length -gt 0) 'the VGI extension must not be empty'

$scalar = Invoke-Haybarn 'SELECT open_meteo.main.weather_code_text(0) AS weather;'
Assert-True ($scalar.Count -eq 1) 'scalar query should return one row'
Assert-True ($scalar[0].weather -eq 'Clear sky') "expected Clear sky, got $($scalar[0].weather)"

$table = Invoke-Haybarn "SELECT name, country FROM open_meteo.main.geocoding('Boston', count := 1::BIGINT);"
Assert-True ($table.Count -eq 1) 'geocoding should return one row'
Assert-True ($table[0].name -eq 'Boston') "expected Boston, got $($table[0].name)"
Assert-True ($table[0].country -eq 'United States') "expected United States, got $($table[0].country)"

$forecast = Invoke-Haybarn "SELECT * FROM open_meteo.main.forecast_current(42.3601::DOUBLE, -71.0589::DOUBLE, temperature_unit := 'fahrenheit') LIMIT 1;"
Assert-True ($forecast.Count -eq 1) 'forecast_current should accept typed positional coordinates and named optional parameters'

$functions = Invoke-Haybarn "SELECT function_name, CAST(to_json(parameters) AS VARCHAR) AS parameters, CAST(to_json(parameter_types) AS VARCHAR) AS parameter_types FROM duckdb_functions() WHERE database_name = $(Sql-String $Catalog) AND function_name = 'geocoding' LIMIT 1;"
Assert-True ($functions.Count -eq 1) 'expected the Open Meteo geocoding function in the function inventory'
$functionParameters = @($functions[0].parameters | ConvertFrom-Json)
$functionParameterTypes = @($functions[0].parameter_types | ConvertFrom-Json)
Assert-True ($functionParameters.Count -gt 0) 'expected geocoding parameter names to be valid JSON'
Assert-True ($functionParameters.Count -eq $functionParameterTypes.Count) 'function parameter names and types should align'

$richArgument = Invoke-Haybarn "SELECT arg_type, is_named, is_positional, arg_choices FROM vgi_function_arguments() WHERE catalog_name = $(Sql-String $Catalog) AND function_name = 'forecast_current' AND arg_name = 'temperature_unit';"
Assert-True ($richArgument.Count -eq 1) 'expected rich metadata for forecast_current.temperature_unit'
Assert-True ([bool]$richArgument[0].is_named) 'temperature_unit should be marked as named'
Assert-True (-not [bool]$richArgument[0].is_positional) 'temperature_unit should not be marked as positional'
Assert-True ([string]$richArgument[0].arg_choices -match 'fahrenheit') 'temperature_unit choices should include fahrenheit'

$missingRichFunctions = Invoke-Haybarn "WITH rich AS (SELECT DISTINCT function_name FROM vgi_function_arguments() WHERE catalog_name = $(Sql-String $Catalog)), flat AS (SELECT DISTINCT function_name FROM duckdb_functions() WHERE database_name = $(Sql-String $Catalog)) SELECT count(*) AS missing FROM rich ANTI JOIN flat USING (function_name);"
Assert-True ([int]$missingRichFunctions[0].missing -eq 0) 'combined discovery should retain every VGI function'

$rejected = $false
try { Assert-HttpsLocation 'uv run worker.py' } catch { $rejected = $_.Exception.Message -match 'HTTPS' }
Assert-True $rejected 'local command locations must be rejected before execution'
$rejected = $false
try { Assert-HttpsLocation 'http://vgi.example.com' } catch { $rejected = $_.Exception.Message -match 'HTTPS' }
Assert-True $rejected 'plain HTTP locations must be rejected before execution'

Write-Host 'PASS: native Haybarn HTTPS integration'
