@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
cd /d "%~dp0"
echo === [1/2] Build data.js ===
python build_data.py
if errorlevel 1 goto :err
echo.
echo === [2/2] Build trajectory.js (uses cached prices.json) ===
python build_trajectory.py
echo.
echo === Open index.html ===
start "" "index.html"
exit /b 0
:err
echo.
echo [ERROR] Build failed. See messages above.
pause
exit /b 1