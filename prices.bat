@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
cd /d "%~dp0"
echo === yfinance 가격 수집 (캐시 신선도 기반, 오래된 것만 갱신) ===
echo (전체 재수집: prices.bat --refresh)
python fetch_prices.py %*
if errorlevel 1 goto :err
echo.
echo === trajectory.js 재빌드 ===
python build_trajectory.py
echo.
echo 완료. index.html 새로고침하세요.
exit /b 0
:err
echo.
echo [ERROR] See messages above.
pause
exit /b 1
