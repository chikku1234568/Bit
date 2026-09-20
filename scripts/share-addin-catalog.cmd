@echo off
setlocal
echo Creating a normal (non-admin) share so Excel can see Bit...
pushd "%~dp0.."
set "ADDINDIR=%CD%\addin"
popd
net share BitAddin /delete /y >nul 2>&1
net share BitAddin="%ADDINDIR%" /GRANT:Everyone,READ
if errorlevel 1 (
  echo FAILED - right-click this script and Run as administrator.
  echo Folder was: %ADDINDIR%
  pause
  exit /b 1
)
echo.
echo Share OK: \\%COMPUTERNAME%\BitAddin
echo In Excel Trust Center, Catalog Url = \\%COMPUTERNAME%\BitAddin
echo Tick Show in Menu, OK, then fully quit Excel and reopen.
echo SHARED FOLDER tab - Refresh - Bit - Add.
pause
