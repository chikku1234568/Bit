@echo off
echo Creating a normal (non-admin) share so Excel can see Bit...
net share BitAddin /delete /y >nul 2>&1
net share BitAddin="C:\Users\Srikar\Bit\Bit\addin" /GRANT:Everyone,READ
if errorlevel 1 (
  echo FAILED - this window must be Run as administrator.
  pause
  exit /b 1
)
echo.
echo Share OK: \\%COMPUTERNAME%\BitAddin
echo Close this window, then in Excel Trust Center use that path.
pause
