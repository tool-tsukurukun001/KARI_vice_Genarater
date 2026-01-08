@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo Gitリポジトリを初期化しています...
git init

echo.
echo ファイルをステージングしています...
git add .

echo.
echo コミットしています...
git commit -m "Initial commit: ElevenLabs Voice Generator Tool"

echo.
echo 完了しました！
echo.
git log --oneline -1
echo.
pause


