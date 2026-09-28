[CmdletBinding()]
param([Parameter(Mandatory=$true)][string] $Path, [switch] $ValidateOnly)
$ErrorActionPreference = 'Stop'
$file = Get-Item -LiteralPath $Path
if ($file.Length -gt 1MB) { throw 'Connection defaults exceed the size limit.' }
$text = [IO.File]::ReadAllText($file.FullName)
if (!$text.TrimStart().StartsWith('[')) { throw 'Connection defaults must be a JSON array.' }
try { $definitions = ConvertFrom-Json -InputObject $text } catch { throw 'Connection defaults contain invalid JSON.' }
$names = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
foreach ($item in $definitions) {
    if ($null -eq $item) { throw 'A connection definition cannot be null.' }
    foreach ($field in $item.PSObject.Properties.Name) {
        if ($field -cnotin @('Name','Catalog','Location','Authentication','AttachOptions','Members')) { throw 'Connection defaults contain an unsupported field. Credentials are not accepted.' }
    }
    if ($item.Name -isnot [string] -or [string]::IsNullOrWhiteSpace($item.Name)) { throw 'A connection name is required.' }
    if ($null -ne $item.Members -and $item.Members -isnot [array]) { throw 'Profile members must be an array.' }
    if (@($item.Members).Count -gt 0 -and $null -ne $item.Members) {
        if (!$names.Add($item.Name)) { throw 'Connection defaults contain duplicate names.' }
        if ($item.Members.Count -gt 16) { throw 'Profiles support at most 16 members.' }
        if ($item.Location -or ($null -ne $item.Authentication -and $item.Authentication -cne 'anonymous') -or ($null -ne $item.AttachOptions -and ($item.AttachOptions -isnot [PSCustomObject] -or @($item.AttachOptions.PSObject.Properties).Count -gt 0))) { throw "Profiles use their members' endpoints, authentication, and options." }
        continue
    }
    foreach ($field in @('Name','Catalog','Location')) { if ($item.$field -isnot [string] -or [string]::IsNullOrWhiteSpace($item.$field)) { throw 'Connection defaults require name, catalog, and HTTPS location strings.' } }
    if (!$names.Add($item.Name)) { throw 'Connection defaults contain duplicate names.' }
    $uri = $null
    if (![Uri]::TryCreate($item.Location, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https' -or $uri.UserInfo) { throw 'Only HTTPS endpoints without embedded credentials are supported.' }
    if ($null -ne $item.Authentication -and $item.Authentication -cnotin @('anonymous','oauth')) { throw 'Authentication must be anonymous or oauth.' }
    if ($null -ne $item.AttachOptions) {
        if ($item.AttachOptions -isnot [PSCustomObject]) { throw 'ATTACH options must be an object.' }
        foreach ($option in $item.AttachOptions.PSObject.Properties) {
            if ($option.Name -notmatch '^[A-Za-z_][A-Za-z0-9_]*$' -or $option.Name -in @('type','location','access_token','api_key','authorization','bearer_token','client_secret','id_token','oauth_refresh_token','password','refresh_token','secret')) { throw 'ATTACH options contain a reserved or credential field.' }
            if ($null -ne $option.Value -and $option.Value -isnot [string] -and $option.Value -isnot [bool] -and $option.Value -isnot [ValueType]) { throw 'ATTACH option values must be scalars.' }
        }
    }
}
foreach ($item in $definitions) {
    if ($null -eq $item.Members -or $item.Members.Count -eq 0) { continue }
    $memberNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $aliases = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $item.Members) {
        if ($name -isnot [string] -or [string]::IsNullOrWhiteSpace($name) -or !$memberNames.Add($name)) { throw 'Profile members must be distinct connection names.' }
        $matches = @($definitions | Where-Object { [string]::Equals($_.Name,$name,[StringComparison]::OrdinalIgnoreCase) })
        if ($matches.Count -ne 1) { throw 'A profile member is missing or ambiguous.' }
        $member = $matches[0]
        if ($null -ne $member.Members -and $member.Members.Count -gt 0) { throw 'Profiles cannot contain other profiles.' }
        if (!$aliases.Add($member.Catalog)) { throw 'Profile members must have distinct catalog aliases.' }
    }
}
if ($ValidateOnly) { Write-Output 'PASS: connection defaults validated'; return }
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (!$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Provisioning machine defaults requires administrator context.' }
$product = Get-ItemProperty 'HKLM:\SOFTWARE\QueryFarm\Cupola'
$target = Join-Path $product.InstallDirectory 'connection-defaults.json'
$temporary = $target + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
try {
    [IO.File]::WriteAllText($temporary, $text, (New-Object Text.UTF8Encoding($false)))
    if (Test-Path -LiteralPath $target) { [IO.File]::Replace($temporary, $target, $target + '.bak') }
    else { [IO.File]::Move($temporary, $target) }
} finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
Write-Output 'Validated Cupola defaults were installed. Missing connections will be provisioned when each user next starts Excel; existing settings are preserved.'
