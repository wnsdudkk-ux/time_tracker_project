@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
cd /d "%~dp0"
echo === Fetch yfinance prices (cached; only stale refetched) ===
echo (Full re-fetch: prices.bat --refresh)
python fetch_prices.py %*
if errorlevel 1 goto :err
echo.
echo === Rebuild trajectory.js ===
python build_trajectory.py
echo.
echo Done. Refresh index.html in your browser.
exit /b 0
:err
echo.
echo [ERROR] See messages above.
pause
exit /b 1