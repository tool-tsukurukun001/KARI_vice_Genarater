@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ========================================
echo Git リポジトリ設定 & プッシュ
echo ========================================
echo.

echo Gitリポジトリを初期化しています...
git init

echo.
echo リモートリポジトリを設定しています...
git remote add origin https://github.com/tool-tsukurukun001/KARI_vice_Genarater.git

echo.
echo ファイルをステージングしています...
git add .

echo.
echo コミットしています...
git commit -m "Initial commit: ElevenLabs Voice Generator Tool"

echo.
echo mainブランチに変更しています...
git branch -M main

echo.
echo GitHubにプッシュしています...
git push -u origin main

echo.
echo ========================================
echo 完了しました！
echo ========================================
echo.
pause

