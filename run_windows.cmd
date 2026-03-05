@echo off
setlocal EnableExtensions EnableDelayedExpansion

cd /d "%~dp0"

set "MODE=%~1"
set "TARGET_SCRIPT="
set "TARGET_NAME="
set "FORWARD_ARGS="

if "%MODE%"=="" (
  set "MODE=huaining"
  set "TARGET_SCRIPT=huaining\\process_huaining.py"
  set "TARGET_NAME=Huaining"
) else if /I "%MODE%"=="huaining" (
  set "TARGET_SCRIPT=huaining\\process_huaining.py"
  set "TARGET_NAME=Huaining"
  set "FORWARD_ARGS=%2 %3 %4 %5 %6 %7 %8 %9"
) else if /I "%MODE%"=="feixi" (
  set "TARGET_SCRIPT=feixi\\process_feixi.py"
  set "TARGET_NAME=Feixi"
  set "FORWARD_ARGS=%2 %3 %4 %5 %6 %7 %8 %9"
) else if /I "%MODE:~0,2%"=="--" (
  set "MODE=huaining"
  set "TARGET_SCRIPT=huaining\\process_huaining.py"
  set "TARGET_NAME=Huaining"
  set "FORWARD_ARGS=%1 %2 %3 %4 %5 %6 %7 %8 %9"
) else (
  echo [ERROR] Unknown mode: %MODE%
  echo Usage: run_windows.cmd [huaining^|feixi] [processor args]
  exit /b 1
)

set "UV_EXE="
where uv >nul 2>nul && set "UV_EXE=uv"
if not defined UV_EXE if exist "%LOCALAPPDATA%\\Microsoft\\WinGet\\Packages\\astral-sh.uv_Microsoft.Winget.Source_8wekyb3d8bbwe\\uv.exe" set "UV_EXE=%LOCALAPPDATA%\\Microsoft\\WinGet\\Packages\\astral-sh.uv_Microsoft.Winget.Source_8wekyb3d8bbwe\\uv.exe"
if not defined UV_EXE if exist "%USERPROFILE%\\.cargo\\bin\\uv.exe" set "UV_EXE=%USERPROFILE%\\.cargo\\bin\\uv.exe"

if not defined UV_EXE (
  echo [ERROR] uv is not installed or not reachable from CMD.
  echo Install uv first: https://docs.astral.sh/uv/getting-started/installation/
  exit /b 1
)

call "%UV_EXE%" sync
if errorlevel 1 (
  echo [ERROR] uv sync failed.
  exit /b 1
)

call "%UV_EXE%" run python %TARGET_SCRIPT% !FORWARD_ARGS!
if errorlevel 1 (
  echo [ERROR] %TARGET_NAME% processing failed.
  exit /b 1
)

echo [OK] %TARGET_NAME% pipeline finished.
exit /b 0
