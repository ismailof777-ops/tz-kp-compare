@echo off
cd /d "%~dp0"
title TZ KP Compare - local server
set "BUNDLED_PY=C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "PYTHONPATH=%CD%\.python_packages;%PYTHONPATH%"
set "APP_PORT=8001"

echo Stopping old local server on port %APP_PORT%...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$serverProcessIds = @(Get-NetTCPConnection -LocalPort %APP_PORT% -State Listen -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique); foreach ($serverProcessId in $serverProcessIds) { if ($serverProcessId -gt 0) { Stop-Process -Id $serverProcessId -Force -ErrorAction SilentlyContinue } }"
echo.
echo Starting local server...
echo Do not close this window while using the service.
echo.
start "" powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 2; Start-Process 'http://127.0.0.1:%APP_PORT%'"
if exist "%BUNDLED_PY%" (
  "%BUNDLED_PY%" web_app.py %APP_PORT%
) else (
  py web_app.py %APP_PORT%
)
pause
