@echo off
title Sistema de Gestao - SP Aguas

cd /d "%~dp0"

set "EXCEL_DATABASE=C:\Users\ana.silva\OneDrive - PRODESP\SP_AGUAS\Sistema de Gestão de Demandas - SP Aguas.xlsx"

if not exist "%EXCEL_DATABASE%" (
    echo Planilha do OneDrive nao encontrada:
    echo %EXCEL_DATABASE%
    echo Verifique se o OneDrive esta instalado e sincronizado.
    pause
    exit /b 1
)

echo Iniciando sistema...

if not exist ".venv\Scripts\activate.bat" (
    echo Ambiente virtual nao encontrado em .venv\Scripts.
    echo Crie o ambiente virtual e instale as dependencias antes de iniciar.
    pause
    exit /b 1
)

call ".venv\Scripts\activate.bat"

start "" http://localhost:5000

python app.py

pause