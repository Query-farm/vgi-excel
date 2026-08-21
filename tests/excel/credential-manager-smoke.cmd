@echo off
setlocal
set "VGI_EXCEL_REQUIRE_CREDENTIAL_MANAGER=1"
cd /d "%~dp0\..\.."
if not exist artifacts mkdir artifacts
dotnet run --project windows\Vgi.ExcelDna.Tests\Vgi.ExcelDna.Tests.csproj -c Release > artifacts\credential-manager-smoke.log 2>&1
exit /b %errorlevel%
