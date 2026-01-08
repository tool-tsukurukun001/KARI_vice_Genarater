@echo off
chcp 65001 > nul
echo ========================================
echo FFmpeg インストールスクリプト
echo ========================================
echo.

:: 管理者権限チェック
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo このスクリプトは管理者権限が必要です。
    echo 右クリック → 「管理者として実行」で再度実行してください。
    echo.
    pause
    exit /b 1
)

set FFMPEG_DIR=C:\ffmpeg
set FFMPEG_URL=https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip
set DOWNLOAD_PATH=%TEMP%\ffmpeg.zip

echo FFmpegをダウンロードしています...
echo URL: %FFMPEG_URL%
echo.

powershell -Command "Invoke-WebRequest -Uri '%FFMPEG_URL%' -OutFile '%DOWNLOAD_PATH%'"

if not exist "%DOWNLOAD_PATH%" (
    echo ダウンロードに失敗しました。
    pause
    exit /b 1
)

echo.
echo 解凍しています...

:: 既存のフォルダがあれば削除
if exist "%FFMPEG_DIR%" (
    rmdir /s /q "%FFMPEG_DIR%"
)

:: 一時フォルダに解凍
powershell -Command "Expand-Archive -Path '%DOWNLOAD_PATH%' -DestinationPath '%TEMP%\ffmpeg_temp' -Force"

:: フォルダ名を取得して移動
for /d %%i in ("%TEMP%\ffmpeg_temp\ffmpeg-*") do (
    move "%%i" "%FFMPEG_DIR%"
)

:: 一時ファイルを削除
rmdir /s /q "%TEMP%\ffmpeg_temp" 2>nul
del "%DOWNLOAD_PATH%" 2>nul

echo.
echo 環境変数PATHに追加しています...

:: PATHに追加（既に存在しない場合のみ）
powershell -Command "$oldPath = [Environment]::GetEnvironmentVariable('Path', 'Machine'); if ($oldPath -notlike '*C:\ffmpeg\bin*') { [Environment]::SetEnvironmentVariable('Path', $oldPath + ';C:\ffmpeg\bin', 'Machine'); Write-Host 'PATHに追加しました' } else { Write-Host 'PATHには既に追加されています' }"

echo.
echo ========================================
echo インストール完了！
echo ========================================
echo.
echo FFmpegは C:\ffmpeg にインストールされました。
echo.
echo 新しいコマンドプロンプトを開いて、以下のコマンドで確認できます:
echo   ffmpeg -version
echo.
echo ※現在開いているコマンドプロンプトでは反映されません。
echo   新しいウィンドウを開いてください。
echo.
pause


