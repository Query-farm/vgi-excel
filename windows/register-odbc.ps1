[CmdletBinding()]
param([Parameter(Mandatory=$true)][string] $DriverPath)
$ErrorActionPreference='Stop'
$driver=(Resolve-Path -LiteralPath $DriverPath).Path
$principal=New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (!$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Elevate only machine-wide driver registration; Excel and the user stores stay in the original session.
    $process=Start-Process powershell.exe -Verb RunAs -Wait -PassThru -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"'+$PSCommandPath+'"'),'-DriverPath',('"'+$driver+'"'))
    if ($process.ExitCode -ne 0) {throw 'Cupola ODBC registration failed.'}
    exit 0
}
$stream=[IO.File]::OpenRead($driver)
try {
    $reader=New-Object IO.BinaryReader($stream)
    $stream.Position=0x3c;$pe=$reader.ReadInt32();$stream.Position=$pe
    if ($reader.ReadUInt32() -ne 0x4550 -or $reader.ReadUInt16() -ne 0x8664) {throw 'The Cupola ODBC driver must be a Windows x64 DLL.'}
} finally {$stream.Dispose()}
$root=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[Microsoft.Win32.RegistryView]::Registry64)
try {
    $key=$root.CreateSubKey('SOFTWARE\ODBC\ODBCINST.INI\Cupola for Excel')
    try {
        $key.SetValue('Driver',$driver);$key.SetValue('APILevel','2');$key.SetValue('ConnectFunctions','YYY');$key.SetValue('DriverODBCVer','03.00')
    } finally {$key.Dispose()}
    $list=$root.CreateSubKey('SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers')
    try {$list.SetValue('Cupola for Excel','Installed')} finally {$list.Dispose()}
} finally {$root.Dispose()}
Write-Host 'Registered the Cupola for Excel x64 ODBC driver.'
