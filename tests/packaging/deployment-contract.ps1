param([Parameter(Mandatory=$true)][string] $MsiPath)
$ErrorActionPreference='Stop'
function Assert($Value,[string]$Message){if(!$Value){throw $Message}}
function Call($Object,[string]$Name,[string]$Kind,[object[]]$Arguments){$Object.GetType().InvokeMember($Name,$Kind,$null,$Object,$Arguments)}
$installer=New-Object -ComObject WindowsInstaller.Installer
$db=Call $installer OpenDatabase InvokeMethod @((Resolve-Path $MsiPath).Path,0)
function Rows([string]$Query,[int]$Columns){
 $v=Call $db OpenView InvokeMethod @($Query);$null=Call $v Execute InvokeMethod @()
 $rows=@()
 while($true){$r=Call $v Fetch InvokeMethod @();if($null -eq $r){break};$values=@();for($i=1;$i -le $Columns;$i++){$values += [string](Call $r StringData GetProperty @($i))};$rows+=,@($values)}
 return ,$rows
}
$registry=Rows 'SELECT `Root`, `Key`, `Name`, `Value` FROM `Registry`' 4
$loader=@($registry | Where-Object {$_[0] -eq '2' -and $_[1] -eq 'SOFTWARE\Microsoft\Office\Excel\Addins\QueryFarm.Cupola.ExcelLoader' -and $_[2] -eq 'LoadBehavior' -and $_[3] -eq '#3'})
Assert ($loader.Count -eq 1) 'MSI must register the automatic Excel loader machine-wide'
Assert (@($registry | Where-Object {$_[0] -eq '1'}).Count -eq 0) 'Machine installation must not write the SYSTEM account HKCU'
Assert (@($registry | Where-Object {$_[1] -match 'desktop-connections|oauth-sessions|default-connection'}).Count -eq 0) 'Installer must not own user connection or OAuth state'
$launch=Rows 'SELECT `Condition` FROM `LaunchCondition`' 1
foreach($required in @('CUPOLA_PLATFORM_SUPPORTED','CUPOLA_EXCEL_RUNNING','NETFRAMEWORKRELEASE','WEBVIEWVERSION','EXCELPLATFORM')){Assert (@($launch | Where-Object {$_[0].Contains($required)}).Count -gt 0) "Missing deployment guard $required"}
$sequence=Rows 'SELECT `Action`, `Sequence` FROM `InstallExecuteSequence`' 2
$check=@($sequence | Where-Object {$_[0] -eq 'CheckCupolaEnvironment'})
$conditions=@($sequence | Where-Object {$_[0] -eq 'LaunchConditions'})
$initialize=@($sequence | Where-Object {$_[0] -eq 'InstallInitialize'})
$remove=@($sequence | Where-Object {$_[0] -eq 'RemoveExistingProducts'})
Assert ([int]$check[0][1] -lt [int]$conditions[0][1] -and [int]$conditions[0][1] -lt [int]$initialize[0][1]) 'Running Excel must be detected before starting the installer transaction'
Assert ([int]$remove[0][1] -gt [int]$initialize[0][1]) 'Major upgrades must remove the previous version inside the rollback transaction'
$files=Rows 'SELECT `FileName` FROM `File`' 1
foreach($required in @('Cupola.ExcelLoader.dll','cupola-mark.svg','build-info.json')){Assert (@($files | Where-Object {($_[0] -split '\|')[-1] -eq $required}).Count -eq 1) "MSI must contain $required"}
Write-Output 'PASS: machine activation, prerequisite guards, transaction ordering, user-data isolation, and complete web assets'
