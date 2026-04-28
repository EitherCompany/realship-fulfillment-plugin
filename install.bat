@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo.
echo ============================================
echo  realship-fulfillment v2 설치
echo ============================================
echo.
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0install.ps1"
echo.
echo 창을 닫아도 됩니다.
pause
