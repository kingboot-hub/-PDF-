@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PY="
where python.exe >nul 2>nul
if not errorlevel 1 set "PY=python"
if not defined PY (
  where py.exe >nul 2>nul
  if not errorlevel 1 set "PY=py -3"
)
if not defined PY (
  echo [ERROR] Python 3.9+ not found in PATH.
  pause
  exit /b 1
)
echo ================================================
echo   Starting with console - errors will show here
echo ================================================
%PY% "%~dp0PasteCompress.py" --debug %*
echo.
echo (press any key to close this window)
pause >nul
