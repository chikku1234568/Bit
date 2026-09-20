@echo off
set ROOT=%~dp0..
start "Bit agent" cmd /k "cd /d "%ROOT%" && npm run api"
timeout /t 2 /nobreak >nul
start "Bit web / add-in" cmd /k "cd /d "%ROOT%" && npm run dev:web"
echo.
echo Bit agent   http://127.0.0.1:3001
echo Add-in      http://127.0.0.1:5173/addin.html
echo Browser lab http://127.0.0.1:5173
echo.
echo Keep these two windows open. In Excel: Home - Bit.
echo First time: sideload addin\manifest.xml  (README.md)
echo Colleague: Choose folder = YOUR OneDrive sync path, then Fetch.
echo.
pause
