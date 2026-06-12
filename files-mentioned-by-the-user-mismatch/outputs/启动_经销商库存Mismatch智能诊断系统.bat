@echo off
chcp 65001 >nul
set "APP_DIR=%~dp0"
set "PYTHON_EXE=C:\Users\Leo.Li\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if not exist "%PYTHON_EXE%" (
  echo Cannot find bundled Python runtime:
  echo %PYTHON_EXE%
  pause
  exit /b 1
)

"%PYTHON_EXE%" "%APP_DIR%dealer_mismatch_diagnostic_app.py"
