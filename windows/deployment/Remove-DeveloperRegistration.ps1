[CmdletBinding()]
param()
$ErrorActionPreference='Stop'

function Remove-CupolaStartupRegistration([string] $OfficeRoot) {
    # Match the Cupola filename, not another add-in that happens to contain it.
    $ownedPattern = '(?i)(?:^|[\\"])Vgi\.ExcelDna(?:64)?-packed(?:-[^\\"]+)?\.xll(?:"|$)'
    foreach ($version in @(Get-ChildItem $OfficeRoot -ErrorAction SilentlyContinue | Where-Object PSChildName -match '^\d+\.\d+$')) {
        $options = Join-Path $version.PSPath 'Excel\Options'
        if (Test-Path $options) {
            $key = Get-Item $options
            try {
                $entries = @($key.GetValueNames() | Where-Object { $_ -match '^OPEN\d*$' } | Sort-Object {
                    if ($_ -eq 'OPEN') { 0 } else { [int]$_.Substring(4) + 1 }
                } | ForEach-Object {
                    [PSCustomObject]@{
                        Name = $_
                        Value = $key.GetValue($_, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                        Kind = $key.GetValueKind($_)
                    }
                })
            } finally { $key.Dispose() }
            if (@($entries | Where-Object { [string]$_.Value -match $ownedPattern }).Count) {
                $retained = @($entries | Where-Object { [string]$_.Value -notmatch $ownedPattern })
                foreach ($entry in $entries) { Remove-ItemProperty $options $entry.Name }
                for ($index = 0; $index -lt $retained.Count; $index++) {
                    $name = if ($index -eq 0) { 'OPEN' } else { "OPEN$index" }
                    New-ItemProperty $options $name -Value $retained[$index].Value -PropertyType $retained[$index].Kind -Force | Out-Null
                }
            }
        }
        # Add-in Manager can retain an old path even without an OPEN entry.
        $manager = Join-Path $version.PSPath 'Excel\Add-in Manager'
        if (Test-Path $manager) {
            foreach ($entry in @((Get-ItemProperty $manager).PSObject.Properties | Where-Object { $_.Name -match $ownedPattern })) {
                Remove-ItemProperty $manager $entry.Name
            }
        }
    }
}

# Dot-sourcing exposes the helper to isolated registry tests without migrating the user.
if ($MyInvocation.InvocationName -ne '.') {
    Remove-CupolaStartupRegistration 'HKCU:\Software\Microsoft\Office'
    Write-Output 'Developer Cupola startup registration was removed for this user. Existing Excel sessions continue until closed; saved connections and sign-ins are unchanged.'
}
