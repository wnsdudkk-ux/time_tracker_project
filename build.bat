@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
echo === Build data.js ===
python build_data.py
if errorlevel 1 goto :err
echo.
echo === Open index.html ===
start "" "index.html"
exit /b 0
:err
echo.
echo [ERROR] Build failed. See messages above.
pause
exit /b 1
