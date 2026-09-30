param([Parameter(Mandatory=$true)][string] $DriverPath, [Parameter(Mandatory=$true)][string] $VgiExtensionPath)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Security
$testRoot=Join-Path $env:TEMP ('cupola-odbc-'+[guid]::NewGuid().ToString('N'))
$oldRoot=$env:VGI_EXCEL_CONFIG_HOME
$driverName='Cupola Test '+[guid]::NewGuid().ToString('N')
$env:VGI_EXCEL_CONFIG_HOME=$testRoot
$registry=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[Microsoft.Win32.RegistryView]::Registry64)
function Assert-True($Value,[string]$Message) {if (!$Value) {throw $Message}}
function Write-Connections($Values) {ConvertTo-Json -InputObject @($Values) -Depth 8 | Set-Content (Join-Path $testRoot 'desktop-connections.json') -Encoding UTF8}
function Open-Connection([string]$Name,[string]$Extra='') {
 $c=New-Object System.Data.Odbc.OdbcConnection("Driver={$driverName};CupolaConnection={$($Name.Replace('}','}}'))};$Extra")
 try {$c.Open();return $c} catch {$c.Dispose();throw}
}
function Expect-Rejected([scriptblock]$Action,[string]$Message) {
 try {$c=& $Action;if($c){$c.Dispose()}} catch [System.Data.Odbc.OdbcException] {
  Assert-True (!$_.Exception.Message.Contains('secret-fixture')) 'ODBC must not echo credentials'
  if ($Message -eq 'expired decrypted OAuth session') {Assert-True ($_.Exception.Message -match 'expired') 'DPAPI must decrypt the expired session before rejecting it'}
  if ($Message -eq 'corrupt OAuth session') {Assert-True ($_.Exception.Message -match 'decrypt') 'corrupt DPAPI session must be rejected explicitly'}
  return
 }
 throw "Expected rejection: $Message"
}
try {
 New-Item -ItemType Directory -Force $testRoot | Out-Null
 $driverDir=Join-Path $testRoot 'driver';New-Item -ItemType Directory $driverDir | Out-Null
 Copy-Item $DriverPath (Join-Path $driverDir 'haybarn_odbc.dll');Copy-Item $VgiExtensionPath (Join-Path $driverDir 'vgi.duckdb_extension')
 $key=$registry.CreateSubKey("SOFTWARE\ODBC\ODBCINST.INI\$driverName");$key.SetValue('Driver',(Join-Path $driverDir 'haybarn_odbc.dll'));$key.Dispose()
 $list=$registry.CreateSubKey('SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers');$list.SetValue($driverName,'Installed');$list.Dispose()
 $name="Cupola }; ' Ω"
 $connection=@{Name=$name;Catalog='open_meteo';Location='https://vgi-open-meteo.rusty-bb6.workers.dev';Authentication='anonymous';AttachOptions=@{}}
 Write-Connections @($connection)
 Expect-Rejected {Open-Connection 'missing'} 'no fallback to another connection'
 Expect-Rejected {Open-Connection $name 'Database=:memory:;'} 'connection overrides'
 Expect-Rejected {Open-Connection $name 'CupolaConnection=other;'} 'duplicate identities'
 Expect-Rejected {Open-Connection $name 'PWD=secret-fixture;'} 'credentials in connection string'
 $connection.Location='http://example.test';Write-Connections @($connection);Expect-Rejected {Open-Connection $name} 'HTTPS only'
 $connection.Location='https://user:secret-fixture@example.test';Write-Connections @($connection);Expect-Rejected {Open-Connection $name} 'URL credentials'
 $connection.Location='https://vgi-open-meteo.rusty-bb6.workers.dev';$connection.AttachOptions=@{bearer_token='secret-fixture'};Write-Connections @($connection);Expect-Rejected {Open-Connection $name} 'stored credentials'
 $connection.AttachOptions=@{};$connection.Authentication='oauth';Write-Connections @($connection);Expect-Rejected {Open-Connection $name} 'missing OAuth session'
 # Create an expired DPAPI session using exactly the same entropy and filename as the XLL.
 $sha=[Security.Cryptography.SHA256]::Create()
 function Hash([string]$Text) {return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text))).Replace('-','')}
 $target='QueryFarm/VgiExcel/OAuth/'+(Hash $connection.Location)
 $sessions=Join-Path $testRoot 'oauth-sessions';New-Item -ItemType Directory $sessions | Out-Null
 $sessionPath=Join-Path $sessions ((Hash $target)+'.bin')
 $plain=[Text.Encoding]::UTF8.GetBytes('{"access_token":"secret-fixture","ExpiresAtUtc":"2000-01-01T00:00:00Z"}')
 $entropy=[Text.Encoding]::UTF8.GetBytes('QueryFarm.CupolaForExcel.OAuth.v1')
 [IO.File]::WriteAllBytes($sessionPath,[Security.Cryptography.ProtectedData]::Protect($plain,$entropy,[Security.Cryptography.DataProtectionScope]::CurrentUser))
 Expect-Rejected {Open-Connection $name} 'expired decrypted OAuth session'
 [IO.File]::WriteAllBytes($sessionPath,[byte[]](1,2,3));Expect-Rejected {Open-Connection $name} 'corrupt OAuth session'
 $sha.Dispose()
 $connection.Authentication='anonymous';Write-Connections @($connection)
 Write-Host 'Testing single-catalog HTTPS attachment'
 $c=Open-Connection $name
 try {
  $q=$c.CreateCommand();$q.CommandTimeout=60
  $q.CommandText='SELECT connection_name, catalog_alias, location, authentication, attach_options FROM cupola_connection_info()'
  $r=$q.ExecuteReader();Assert-True ($r.Read()) 'driver identity missing';Assert-True ($r.GetString(0) -eq $name) 'escaped name roundtrip';Assert-True ($r.GetString(1) -eq 'open_meteo') 'catalog identity';Assert-True ($r.GetString(2) -eq $connection.Location) 'endpoint identity';Assert-True ($r.GetString(4) -eq '{}') 'options identity';$r.Close()
  $q.CommandText="SELECT count(*) FROM duckdb_databases() WHERE database_name='open_meteo' AND type='vgi'";Assert-True ([int]$q.ExecuteScalar() -eq 1) 'real VGI catalog attachment'
  $q.CommandText="SELECT count(*) FROM open_meteo.main.forecast_current(42.3601::DOUBLE, -71.0589::DOUBLE, temperature_unit := 'fahrenheit')";Assert-True ([int]$q.ExecuteScalar() -gt 0) 'real HTTPS query through ODBC'
  $q.Dispose()
 } finally {$c.Dispose()}
 $earthquakes=@{Name='Earthquakes';Catalog='earthquakes';Location='https://vgi-earthquakes.rusty-bb6.workers.dev';Authentication='anonymous';AttachOptions=@{}}
 $profile=@{IsWorkspaceProfile=$true;Name='Research';Members=@($name,'Earthquakes');Location='';Authentication='anonymous';AttachOptions=@{}}
 Write-Connections @($connection,$earthquakes,$profile)
 Write-Host 'Testing multi-catalog HTTPS attachment'
 $c=Open-Connection 'research'
 try {
  $q=$c.CreateCommand();$q.CommandTimeout=120
  $q.CommandText="SELECT count(*) FROM cupola_connection_info() WHERE contract_version=2 AND connection_name='Research'";Assert-True ([int]$q.ExecuteScalar() -eq 2) 'profile reports both identities'
  $q.CommandText="SELECT current_catalog()";Assert-True ($q.ExecuteScalar() -eq 'open_meteo') 'first member is default catalog'
  $q.CommandText="SELECT open_meteo.main.weather_code_text(0) FROM earthquakes.main.recent LIMIT 1";Assert-True ($q.ExecuteScalar() -eq 'Clear sky') 'one ODBC query uses both live VGI catalogs'
  $q.Dispose()
 } finally {$c.Dispose()}
 Write-Connections @($connection,$profile);Expect-Rejected {Open-Connection 'Research'} 'missing member fails closed'
 $earthquakes.Catalog='OPEN_METEO';Write-Connections @($connection,$earthquakes,$profile);Expect-Rejected {Open-Connection 'Research'} 'catalog alias collision'
 $earthquakes.Catalog='earthquakes';$earthquakes.Members=@($name);Write-Connections @($connection,$earthquakes,$profile);Expect-Rejected {Open-Connection 'Research'} 'nested profiles'
 $earthquakes.Remove('Members');$profile.Members=@($name,$name);Write-Connections @($connection,$earthquakes,$profile);Expect-Rejected {Open-Connection 'Research'} 'duplicate profile members'
 $profile.Members=@('Earthquakes',$name);$earthquakes.Catalog='earthquakes';Write-Connections @($connection,$earthquakes,$profile)
 Write-Host 'Testing reopened profile with changed default catalog'
 $c=Open-Connection 'Research'
 try {
  $q=$c.CreateCommand();$q.CommandText="SELECT current_catalog()";Assert-True ($q.ExecuteScalar() -eq 'earthquakes') 'new ODBC session reads updated member order';$q.Dispose()
 } finally {$c.Dispose()}
 Write-Host 'PASS: Cupola ODBC identity, policy, DPAPI failures, live HTTPS, and multi-catalog profiles'

} finally {
 $env:VGI_EXCEL_CONFIG_HOME=$oldRoot
 $registry.DeleteSubKeyTree("SOFTWARE\ODBC\ODBCINST.INI\$driverName",$false)
 $list=$registry.OpenSubKey('SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers',$true);if($list){$list.DeleteValue($driverName,$false);$list.Dispose()}
 $registry.Dispose()
 # Driver modules may remain loaded until this PowerShell process exits; remove data now.
 Remove-Item (Join-Path $testRoot 'desktop-connections.json') -Force -ErrorAction SilentlyContinue
 Remove-Item (Join-Path $testRoot 'oauth-sessions') -Recurse -Force -ErrorAction SilentlyContinue
}
