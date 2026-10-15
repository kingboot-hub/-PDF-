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
echo Installing pikepdf + pypdfium2 into .\libs  (internet required)...
%PY% -m pip install --target "%~dp0libs" --upgrade pikepdf pypdfium2
echo.
echo Finished. Run the app again.
pause
