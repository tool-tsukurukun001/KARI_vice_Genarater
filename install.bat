@echo off
chcp 65001 > nul
echo ========================================
echo ElevenLabs Voice Generator インストール
echo ========================================
echo.

cd /d "%~dp0"

echo 必要なパッケージをインストールしています...
echo.

pip install -r requirements.txt

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo エラー: パッケージのインストールに失敗しました。
    echo Pythonがインストールされているか確認してください。
    pause
    exit /b 1
)

echo.
echo ========================================
echo インストール完了！
echo ========================================
echo.
echo 次に、デスクトップショートカットを作成しますか？
echo.

set /p CREATE_SHORTCUT="ショートカットを作成する場合は Y を入力してください: "

if /i "%CREATE_SHORTCUT%"=="Y" (
    call create_shortcut.bat
) else (
    echo.
    echo ショートカットの作成をスキップしました。
    echo 後から create_shortcut.bat を実行することで作成できます。
    echo.
)

echo.
echo セットアップが完了しました！
echo voice_generator.py をダブルクリックするか、
echo デスクトップのショートカットから起動できます。
echo.
pause




