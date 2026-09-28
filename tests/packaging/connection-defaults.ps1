$ErrorActionPreference='Stop'
$validate=Join-Path $PSScriptRoot '..\..\windows\deployment\Set-CupolaDefaults.ps1'
$path=Join-Path $env:TEMP ('cupola-defaults-'+[guid]::NewGuid().ToString('N')+'.json')
function Check($Values,[bool]$Accept) {
 ConvertTo-Json -InputObject @($Values) -Depth 8 | Set-Content -LiteralPath $path -Encoding UTF8
 $rejected=$false
 try {& $validate -Path $path -ValidateOnly | Out-Null} catch {$rejected=$true}
 if($rejected -eq $Accept){throw 'Unexpected connection defaults validation result'}
}
try {
 $a=@{Name='Weather';Catalog='weather';Location='https://weather.example.test';Authentication='anonymous';AttachOptions=@{}}
 $b=@{Name='Quakes';Catalog='quakes';Location='https://quakes.example.test';Authentication='anonymous';AttachOptions=@{}}
 $profile=@{Name='Research';Members=@('Weather','Quakes');Location='';Authentication='anonymous';AttachOptions=@{}}
 Check @($a) $true
 Check @($a,$b,$profile) $true
 Check @($a,$profile) $false
 $b.Catalog='WEATHER';Check @($a,$b,$profile) $false;$b.Catalog='quakes'
 $profile.Members=@('Weather','weather');Check @($a,$b,$profile) $false
 $profile.Members=@('Weather','Quakes');$b.Members=@('Weather');Check @($a,$b,$profile) $false;$b.Remove('Members')
 $profile.AttachOptions=@{password='secret-fixture'};Check @($a,$b,$profile) $false
 Write-Host 'PASS: deployment defaults preserve single catalogs and validate profiles without machine writes'
} finally {Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue}
