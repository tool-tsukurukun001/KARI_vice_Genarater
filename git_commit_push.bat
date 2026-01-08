@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ========================================
echo Git コミット & プッシュ
echo ========================================
echo.

echo ファイルをステージングしています...
git add .

echo.
echo コミットしています...
git commit -m "VoiceVox版に変更: 感情自動判定機能追加、プレビュー修正"

echo.
echo GitHubにプッシュしています...
git push origin main

echo.
echo ========================================
echo 完了しました！
echo ========================================
echo.
pause

