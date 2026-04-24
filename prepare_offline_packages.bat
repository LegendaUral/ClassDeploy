@echo off
setlocal EnableExtensions
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 (
  set "PY=py -3"
) else (
  set "PY=python"
)
if not exist "vendor\wheels" mkdir "vendor\wheels"
%PY% -m pip download -r requirements.txt -d vendor\wheels
pause
