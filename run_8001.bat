@echo off
cd /d "%~dp0"
title TZ KP Compare - port 8001
set "BUNDLED_PY=C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "PYTHONPATH=%CD%\.python_packages;%PYTHONPATH%"

echo Starting service on http://127.0.0.1:8001
echo Keep this window open.
echo.

if exist "%BUNDLED_PY%" (
  "%BUNDLED_PY%" web_app.py 8001
) else (
  py web_app.py 8001
)

echo.
echo Server stopped.
pause
