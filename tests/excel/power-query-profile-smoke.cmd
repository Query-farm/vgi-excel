@echo off
setlocal
cd /d "%~dp0\..\.."
set "VGI_EXCEL_TELEMETRY=0"
set "CUPOLA_TEST_INTERACTIVE_POWERQUERY=1"
set "CUPOLA_ODBC_DRIVER_PATH=%CD%\artifacts\xll\haybarn_odbc.dll"
if not exist "%CUPOLA_ODBC_DRIVER_PATH%" (
  echo Build the Windows artifacts before running this test.
  exit /b 1
)
echo Run this test from an administrator terminal in your Windows desktop session.
echo If Excel asks for ODBC authentication, choose Default or Custom with no credentials.
dotnet run --project windows\Vgi.ExcelDna.Tests\Vgi.ExcelDna.Tests.csproj -c Release -- --power-query-profile
exit /b %errorlevel%
