@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ========================================
echo Git コミット履歴
echo ========================================
echo.

git log --oneline -10

echo.
echo ========================================
echo.
pause

