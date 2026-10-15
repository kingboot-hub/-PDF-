@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PYW="

if exist "%~dp0python\pythonw.exe" set "PYW=%~dp0python\pythonw.exe"

if not defined PYW (
  where pythonw.exe >nul 2>nul
  if not errorlevel 1 set "PYW=pythonw.exe"
)
if not defined PYW (
  where python.exe >nul 2>nul
  if not errorlevel 1 set "PYW=python.exe"
)
if not defined PYW (
  echo.
  echo   [ERROR] Python 3.9+ not found in PATH.
  echo   Install it from https://www.python.org/downloads/
  echo   and check "Add python.exe to PATH" during setup.
  echo.
  echo   Already installed? Try running  ????_????.bat  to see the error.
  echo.
  pause
  exit /b 1
)

start "" "%PYW%" "%~dp0PasteCompress.py" %*
exit /b 0
