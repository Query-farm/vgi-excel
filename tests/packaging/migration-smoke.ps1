$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\..\windows\deployment\Remove-DeveloperRegistration.ps1')
$root = 'HKCU:\Software\QueryFarm\CupolaMigrationTests\' + [Guid]::NewGuid().ToString('N')
function Assert($condition, $message) { if (!$condition) { throw $message } }
try {
    $options = Join-Path $root '16.0\Excel\Options'
    $manager = Join-Path $root '17.0\Excel\Add-in Manager'
    New-Item $options -Force | Out-Null
    New-Item $manager -Force | Out-Null
    $other = '/R "%ProgramFiles%\Other\other.xll"'
    New-ItemProperty $options OPEN -Value $other -PropertyType ExpandString | Out-Null
    New-ItemProperty $options OPEN1 -Value '/R "C:\Cupola\Vgi.ExcelDna64-packed.xll"' | Out-Null
    New-ItemProperty $options OPEN3 -Value '/R "C:\Other\OtherVgi.ExcelDna64-packed.xll"' | Out-Null
    New-ItemProperty $options Unrelated -Value 'preserve' | Out-Null
    New-ItemProperty $manager 'C:\Cupola\Vgi.ExcelDna64-packed-old.xll' -Value '' | Out-Null
    New-ItemProperty $manager 'C:\Other\other.xll' -Value '' | Out-Null
    Remove-CupolaStartupRegistration $root
    $key = Get-Item $options
    try {
        Assert ($key.GetValueNames().Count -eq 3) 'Only Cupola startup entries should be removed'
        Assert ($key.GetValueKind('OPEN') -eq 'ExpandString') 'Other add-in registry types must be preserved'
        Assert ($key.GetValue('OPEN', $null, 'DoNotExpandEnvironmentNames') -ceq $other) 'Other add-in values must be preserved verbatim'
        Assert ($key.GetValue('OPEN1') -match 'OtherVgi') 'Similar filenames must not be treated as Cupola'
        Assert ($key.GetValue('Unrelated') -eq 'preserve') 'Unrelated settings must survive'
    } finally { $key.Dispose() }
    $key = Get-Item $manager
    try { Assert ($key.GetValueNames().Count -eq 1 -and $key.GetValueNames()[0] -eq 'C:\Other\other.xll') 'Manager-only stale Cupola paths must be removed' }
    finally { $key.Dispose() }
    Remove-CupolaStartupRegistration $root
    Assert ((Get-Item $options).GetValueNames().Count -eq 3) 'Migration must be idempotent'
    Write-Output 'PASS: isolated migration preserves unrelated add-ins, value types, and manager-only registrations'
} finally {
    if (Test-Path $root) { Remove-Item $root -Recurse -Force }
}

# Load only the pure process-detection function; never execute installer actions.
$tokens=$null; $parseErrors=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\..\windows\install-xll.ps1'),[ref]$tokens,[ref]$parseErrors)
Assert (!$parseErrors.Count) 'Updater syntax must be valid'
$detector=$ast.Find({param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-RunningExcel'},$true)
Assert ($null -ne $detector) 'Updater must expose live process detection'
& {
    function Get-Process { param($Name,$ErrorAction) @([pscustomobject]@{Id=1;HasExited=$true},[pscustomobject]@{Id=2;HasExited=$false}) }
    . ([scriptblock]::Create($detector.Extent.Text))
    $running=@(Get-RunningExcel)
    Assert ($running.Count -eq 1 -and $running[0].Id -eq 2) 'Exited Excel entries must not block the updater; live processes remain protected'
}
Write-Output 'PASS: updater distinguishes live Excel from terminated process entries'
