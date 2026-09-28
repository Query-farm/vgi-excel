# Build-time integration only. Configuration contains identifiers, never credentials.
function Read-CupolaAzureSigningConfig([string] $Path) {
    $config = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    $fields = @('Endpoint','CodeSigningAccountName','CertificateProfileName','ExpectedSubject')
    foreach ($field in $config.PSObject.Properties.Name) {
        if ($field -cnotin $fields) { throw 'Unknown Azure signing configuration field. Credentials do not belong in this file.' }
    }
    foreach ($field in $fields) {
        if ($config.$field -isnot [string] -or [string]::IsNullOrWhiteSpace($config.$field) -or $config.$field.Contains('REPLACE')) { throw "Configure Azure signing field $field before publishing." }
    }
    $endpoint = $null
    if (![Uri]::TryCreate($config.Endpoint, [UriKind]::Absolute, [ref]$endpoint) -or
        $endpoint.Scheme -ne 'https' -or $endpoint.Host -notmatch '^[a-z0-9]+\.codesigning\.azure\.net$' -or
        $endpoint.UserInfo -or $endpoint.Query -or $endpoint.Fragment -or $endpoint.AbsolutePath -ne '/' -or !$endpoint.IsDefaultPort) {
        throw 'Azure signing requires the regional HTTPS codesigning.azure.net endpoint.'
    }
    foreach ($field in @('CodeSigningAccountName','CertificateProfileName')) {
        if ($config.$field -notmatch '^[a-zA-Z][a-zA-Z0-9-]{2,99}$') { throw "Invalid Azure signing identifier: $field" }
    }
    return $config
}

function Invoke-CupolaAzureSign([string] $Path, $Config) {
    $file = Get-Item -LiteralPath $Path
    if ($file.Extension -notin @('.exe','.dll','.xll','.msi','.ps1')) { throw 'Unsupported Authenticode payload. Never modify the Haybarn-signed VGI extension.' }
    # The official module accepts comma-delimited paths. Reject ambiguous inputs.
    if ($file.FullName.Contains(',') -or $file.FullName.Contains('"')) { throw 'Signing paths cannot contain commas or quotes.' }
    $parameters = @{
        Endpoint = $Config.Endpoint
        CodeSigningAccountName = $Config.CodeSigningAccountName
        CertificateProfileName = $Config.CertificateProfileName
        Files = $file.FullName
        FileDigest = 'SHA256'
        TimestampRfc3161 = 'http://timestamp.acs.microsoft.com'
        TimestampDigest = 'SHA256'
        Timeout = 300
        # Use only the explicit az login identity (interactive workstation or CI OIDC).
        ExcludeEnvironmentCredential = $true
        ExcludeWorkloadIdentityCredential = $true
        ExcludeManagedIdentityCredential = $true
        ExcludeSharedTokenCacheCredential = $true
        ExcludeVisualStudioCredential = $true
        ExcludeVisualStudioCodeCredential = $true
        ExcludeAzureCliCredential = $false
        ExcludeAzurePowerShellCredential = $true
        ExcludeAzureDeveloperCliCredential = $true
        ExcludeInteractiveBrowserCredential = $true
    }
    Invoke-ArtifactSigning @parameters
    $signature = Get-AuthenticodeSignature -LiteralPath $file.FullName
    if ($signature.Status -ne 'Valid' -or !$signature.TimeStamperCertificate -or
        $signature.SignerCertificate.Subject -cne $Config.ExpectedSubject) {
        throw 'Azure signature, timestamp, or expected publisher validation failed.'
    }
}
