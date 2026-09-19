@echo off
set ROOT=%~dp0..
start "Bit agent" cmd /k "cd /d "%ROOT%" && npm run api"
timeout /t 2 /nobreak >nul
start "Bit web / add-in" cmd /k "cd /d "%ROOT%" && npm run dev:web"
echo Bit agent  http://127.0.0.1:3001
echo Add-in     http://127.0.0.1:5173/addin.html
echo Sideload addin\manifest.xml in Excel (see docs\EXCEL-ADDIN.md)
pause
