$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\..\windows\signing\azure.ps1')
$root = Join-Path $env:TEMP ('cupola-signing-tests-' + [Guid]::NewGuid().ToString('N'))
New-Item $root -ItemType Directory | Out-Null
function Assert($condition, $message) { if (!$condition) { throw $message } }
function Reject([scriptblock] $action) { $rejected=$false; try { & $action } catch { $rejected=$true }; Assert $rejected 'Expected signing policy rejection' }
# These doubles never sign, authenticate, or contact Azure.
function Invoke-ArtifactSigning { $script:signArguments = $args; $script:signCalls++ }
function Get-AuthenticodeSignature { param($LiteralPath) return $script:signature }
$script:signCalls=0
try {
    $path=Join-Path $root 'config.json'
    $config=[ordered]@{Endpoint='https://eus.codesigning.azure.net/';CodeSigningAccountName='testaccount';CertificateProfileName='cupola-production';ExpectedSubject='CN=Test Publisher'}
    $config | ConvertTo-Json | Set-Content $path
    $parsed=Read-CupolaAzureSigningConfig $path
    foreach($bad in @('http://eus.codesigning.azure.net/','https://eus.codesigning.azure.net.evil.test/','https://secret@eus.codesigning.azure.net/','https://eus.codesigning.azure.net/?token=fixture')) {
        $config.Endpoint=$bad; $config | ConvertTo-Json | Set-Content $path
        Reject { Read-CupolaAzureSigningConfig $path }
    }
    $config.Endpoint='https://eus.codesigning.azure.net/'
    $config.ClientSecret='fixture'; $config | ConvertTo-Json | Set-Content $path
    Reject { Read-CupolaAzureSigningConfig $path }
    $script:signature=[PSCustomObject]@{Status='Valid';TimeStamperCertificate='timestamp-fixture';SignerCertificate=[PSCustomObject]@{Subject='CN=Test Publisher'}}
    foreach($extension in @('exe','dll','xll','msi','ps1')) {
        $file=Join-Path $root "payload.$extension"; Set-Content $file 'fixture'
        Invoke-CupolaAzureSign $file $parsed
        Assert ($script:signArguments -contains 'SHA256') 'SHA256 required'
        Assert ($script:signArguments -contains 'http://timestamp.acs.microsoft.com') 'Microsoft RFC3161 timestamp required'
        Assert ($script:signArguments -contains '-ExcludeInteractiveBrowserCredential:') 'Unattended signing must not open a browser'
    }
    $file=Join-Path $root 'vgi.duckdb_extension'; Set-Content $file 'engine-signed-fixture'
    Reject { Invoke-CupolaAzureSign $file $parsed }
    Assert ($script:signCalls -eq 5) 'The engine-signed extension must never reach the signing provider'
    $file=Join-Path $root 'payload.dll'
    $script:signature.Status='NotSigned'; Reject { Invoke-CupolaAzureSign $file $parsed }
    $script:signature.Status='Valid'; $script:signature.TimeStamperCertificate=$null
    Reject { Invoke-CupolaAzureSign $file $parsed }
    $script:signature.TimeStamperCertificate='timestamp-fixture';$script:signature.SignerCertificate.Subject='CN=Wrong Publisher'
    Reject { Invoke-CupolaAzureSign $file $parsed }
    Write-Output 'PASS: Azure configuration, payload selection, timestamp, and publisher validation (offline; no signing calls)'
} finally { Remove-Item $root -Recurse -Force }
