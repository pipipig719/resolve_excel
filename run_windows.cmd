@echo off
setlocal EnableExtensions

cd /d "%~dp0"

if not exist "source\" (
  echo [ERROR] source directory is missing.
  echo Create "source" and place the pharmacy-room source workbook there.
  exit /b 1
)

where uv >nul 2>nul
if errorlevel 1 (
  echo [ERROR] uv is not installed or not in PATH.
  echo Install uv first: https://docs.astral.sh/uv/getting-started/installation/
  exit /b 1
)

call uv sync
if errorlevel 1 (
  echo [ERROR] uv sync failed.
  exit /b 1
)

call uv run python run_pipeline.py %*
if errorlevel 1 (
  echo [ERROR] pipeline failed.
  exit /b 1
)

echo [OK] Pipeline finished.
echo [OK] Root output: final import workbook generated.
echo [OK] Source output: backup import workbook generated.
exit /b 0
