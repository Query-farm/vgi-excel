[CmdletBinding()]
param(
    [ValidateSet('Install','Repair','Uninstall')][string] $Action = 'Install',
    [string] $MsiPath = (Join-Path $PSScriptRoot 'CupolaForExcel.msi'),
    [switch] $AllowUnsigned
)
$ErrorActionPreference = 'Stop'
try {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (!$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { Write-Output 'Cupola deployment requires SYSTEM or administrator context.'; exit 740 }
    # Return a retry code. Never close Excel, dismiss prompts, or reboot a desktop.
    if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0) { Write-Output 'Excel is running. Retry Cupola deployment after Excel closes.'; exit 1618 }
    $msi = (Resolve-Path -LiteralPath $MsiPath).Path
    if (!$AllowUnsigned) {
        $signature = Get-AuthenticodeSignature -LiteralPath $msi
        if ($signature.Status -ne 'Valid' -or !$signature.TimeStamperCertificate) { throw 'Cupola MSI must have a valid timestamped signature.' }
    }
    $log = Join-Path $env:TEMP ('Cupola-' + $Action + '-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.log')
    $mode = switch ($Action) { 'Install' { '/i' } 'Repair' { '/famus' } 'Uninstall' { '/x' } }
    $process = Start-Process (Join-Path $env:SystemRoot 'System32\msiexec.exe') -Wait -PassThru -ArgumentList @($mode, ('"' + $msi + '"'), '/qn', '/norestart', '/L*v', ('"' + $log + '"'))
    Write-Output "Cupola deployment returned $($process.ExitCode). Local installer log: $log"
    exit $process.ExitCode
} catch {
    Write-Error $_ -ErrorAction Continue
    exit 1603
}
