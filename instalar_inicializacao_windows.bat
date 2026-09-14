@echo off
title Instalar inicializacao - SP Aguas

cd /d "%~dp0"
set "SP_AGUAS_DIR=%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$startup = [Environment]::GetFolderPath('Startup'); $shortcutPath = Join-Path $startup 'SP Aguas.lnk'; $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcutPath); $shortcut.TargetPath = (Join-Path $env:SP_AGUAS_DIR 'iniciar_sistema.bat'); $shortcut.WorkingDirectory = $env:SP_AGUAS_DIR; $shortcut.WindowStyle = 1; $shortcut.Save()"

if errorlevel 1 (
    echo Nao foi possivel instalar a inicializacao automatica.
    pause
    exit /b 1
)

echo Inicializacao automatica instalada com sucesso.
echo O sistema sera iniciado quando o usuario entrar no Windows.
pause