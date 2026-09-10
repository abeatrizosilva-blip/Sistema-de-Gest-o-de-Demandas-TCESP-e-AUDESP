@echo off
title Sistema de Gestao - SP Aguas

cd /d "%~dp0"

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