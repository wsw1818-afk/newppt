@echo off
setlocal

set "APP=%~dp0DocumentExtractor_v3.exe"
set "RUNTEST=0"

echo DocumentExtractor security/block diagnosis
echo Target: %APP%
echo.
choice /M "Run a short start test for the EXE"
if errorlevel 2 (
  set "RUNTEST=0"
) else (
  set "RUNTEST=1"
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0diagnose_app_block.ps1" -AppPath "%APP%" -RunStartTest "%RUNTEST%"

echo.
echo Diagnosis complete. Send the report file from your Desktop.
pause
