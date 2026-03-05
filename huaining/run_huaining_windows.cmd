@echo off
setlocal EnableExtensions

cd /d "%~dp0\.."

set "UV_EXE="
where uv >nul 2>nul && set "UV_EXE=uv"
if not defined UV_EXE if exist "%LOCALAPPDATA%\Microsoft\WinGet\Packages\astral-sh.uv_Microsoft.Winget.Source_8wekyb3d8bbwe\uv.exe" set "UV_EXE=%LOCALAPPDATA%\Microsoft\WinGet\Packages\astral-sh.uv_Microsoft.Winget.Source_8wekyb3d8bbwe\uv.exe"
if not defined UV_EXE if exist "%USERPROFILE%\.cargo\bin\uv.exe" set "UV_EXE=%USERPROFILE%\.cargo\bin\uv.exe"

if not defined UV_EXE (
  echo [ERROR] uv is not installed or not reachable from CMD.
  echo Install: https://docs.astral.sh/uv/getting-started/installation/
  exit /b 1
)

call "%UV_EXE%" sync
if errorlevel 1 (
  echo [ERROR] uv sync failed.
  exit /b 1
)

call "%UV_EXE%" run python huaining\process_huaining.py %*
if errorlevel 1 (
  echo [ERROR] huaining processing failed.
  exit /b 1
)

echo [OK] Huaining files generated.
exit /b 0
