@echo off
setlocal EnableExtensions

rem Keep the package path free of a trailing backslash. Native argument
rem parsing can otherwise treat the closing quote as part of the path.
set "PACKAGE_DIR=%~dp0."
set "INSTALL_SCRIPT=%~dp0install-xll.ps1"
set "UPDATE_LOG=%TEMP%\Cupola-for-Excel-update.log"

if not exist "%INSTALL_SCRIPT%" (
  echo Cupola for Excel could not find its installer:
  echo %INSTALL_SCRIPT%
  echo.
  pause
  exit /b 2
)

echo Updating Cupola for Excel...
echo Excel will close and reopen. Save any workbook changes before continuing.
echo.
powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%INSTALL_SCRIPT%" -PackagePath "%PACKAGE_DIR%" -RestartExcel > "%UPDATE_LOG%" 2>&1
set "UPDATE_EXIT=%ERRORLEVEL%"
type "%UPDATE_LOG%"

if not "%UPDATE_EXIT%"=="0" (
  echo.
  echo Cupola for Excel was not updated.
  echo Diagnostic log: %UPDATE_LOG%
  pause
  exit /b %UPDATE_EXIT%
)

echo.
echo Update installed successfully. Excel has been reopened with the latest Cupola build.
echo.
pause
exit /b 0
