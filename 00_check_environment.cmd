@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
set "PYTHONUTF8=1"
if not exist ".venv\Scripts\python.exe" (
  echo Python environment is missing.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" "check_environment.py"
set "run_result=%errorlevel%"
pause
exit /b %run_result%
