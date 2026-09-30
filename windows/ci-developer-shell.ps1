# Initialize MSVC for this step and subsequent GitHub Actions steps.
$ErrorActionPreference = 'Stop'
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$installation = & $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if (!$installation) { throw 'The x64 Visual C++ build tools are required.' }
Import-Module (Join-Path $installation 'Common7\Tools\Microsoft.VisualStudio.DevShell.dll')
Enter-VsDevShell -VsInstallPath $installation -SkipAutomaticLocation -DevCmdArguments '-arch=x64 -host_arch=x64'
foreach ($name in @('PATH', 'INCLUDE', 'LIB', 'LIBPATH')) {
    $value = [Environment]::GetEnvironmentVariable($name)
    if ($value) { "$name=$value" | Out-File -FilePath $env:GITHUB_ENV -Encoding utf8 -Append }
}
