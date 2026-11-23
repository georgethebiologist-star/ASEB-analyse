@echo off
cd /d "%~dp0"
py -3 src\waste_analysis.py
if errorlevel 1 (
  echo.
  echo Could not update charts. See the error above and README.md.
  pause
)
