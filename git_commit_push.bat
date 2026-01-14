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
git commit -m "VoiceVox版追加、ElevenLabs版を記録用に保存"

echo.
echo GitHubにプッシュしています...
git push origin main

echo.
echo ========================================
echo 完了しました！
echo ========================================
echo.
pause



