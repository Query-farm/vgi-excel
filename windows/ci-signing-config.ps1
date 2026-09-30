param([Parameter(Mandatory = $true)][string] $Path)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'signing\azure.ps1')
@{
    Endpoint = $env:CUPOLA_AZURE_ENDPOINT
    CodeSigningAccountName = $env:CUPOLA_AZURE_ACCOUNT
    CertificateProfileName = $env:CUPOLA_AZURE_PROFILE
    ExpectedSubject = $env:CUPOLA_AZURE_SUBJECT
} | ConvertTo-Json | Set-Content -LiteralPath $Path -Encoding UTF8
$null = Read-CupolaAzureSigningConfig $Path
Install-PackageProvider NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
Install-Module ArtifactSigning -RequiredVersion 0.1.20 -Repository PSGallery -Scope CurrentUser -Force
Import-Module ArtifactSigning -RequiredVersion 0.1.20
