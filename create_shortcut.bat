@echo off
chcp 65001 > nul
echo デスクトップにショートカットを作成しています...

set SCRIPT_DIR=%~dp0
set SCRIPT_PATH=%SCRIPT_DIR%voice_generator.py
set SHORTCUT_NAME=ElevenLabs Voice Generator

powershell -Command "$WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\%SHORTCUT_NAME%.lnk'); $Shortcut.TargetPath = 'pythonw.exe'; $Shortcut.Arguments = '\"%SCRIPT_PATH%\"'; $Shortcut.WorkingDirectory = '%SCRIPT_DIR%'; $Shortcut.Save()"

echo.
echo ショートカットを作成しました！
echo デスクトップに「%SHORTCUT_NAME%」が追加されています。
echo.
pause

