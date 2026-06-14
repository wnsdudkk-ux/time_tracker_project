@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
cd /d "%~dp0"
echo === [1/3] Download xlsx ===
python download_timeetf.py
if errorlevel 1 goto :err
echo.
echo === [2/3] Build data.js ===
python build_data.py
if errorlevel 1 goto :err
echo.
echo === [3/3] Build trajectory.js (uses cached prices.json) ===
python build_trajectory.py
echo.
echo (매매궤적 주가 갱신은 prices.bat 실행 — 자세한 내용은 README)
echo.
echo === Open index.html ===
start "" "index.html"
exit /b 0
:err
echo.
echo [ERROR] See messages above.
pause
exit /b 1
