@echo off
setlocal EnableExtensions
chcp 65001 >nul
title Sistema de Gestao - SP Aguas

cd /d "%~dp0"

echo ===============================================
echo   Sistema de Gestao - SP Aguas
echo ===============================================
echo.
echo Pasta do sistema: %CD%
echo.

if not exist "%~dp0app.py" (
    echo ERRO: app.py nao foi encontrado nesta pasta.
    echo Extraia TODOS os arquivos do ZIP antes de executar este arquivo.
    pause
    exit /b 1
)
if not exist "%~dp0requirements.txt" (
    echo ERRO: requirements.txt nao foi encontrado nesta pasta.
    echo Extraia TODOS os arquivos do ZIP antes de executar este arquivo.
    pause
    exit /b 1
)

where py >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON=py"
) else (
    where python >nul 2>&1
    if %errorlevel%==0 (
        set "PYTHON=python"
    ) else (
        echo ERRO: Python nao foi encontrado neste computador.
        echo Instale o Python 3.11 ou superior e tente novamente.
        pause
        exit /b 1
    )
)

if not exist "%~dp0.venv\Scripts\python.exe" (
    echo Ambiente virtual nao encontrado.
    echo Criando o ambiente virtual pela primeira vez...
    echo.
    %PYTHON% -m venv "%~dp0.venv"
    if errorlevel 1 (
        echo.
        echo ERRO ao criar o ambiente virtual.
        pause
        exit /b 1
    )
    echo Ambiente virtual criado com sucesso.
    echo.
)

echo Verificando dependencias...
"%~dp0.venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :erro_pip
"%~dp0.venv\Scripts\python.exe" -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 goto :erro_pip

echo.
echo Dependencias prontas.
echo Iniciando o sistema...
echo.
start "" http://127.0.0.1:5000
"%~dp0.venv\Scripts\python.exe" "%~dp0app.py"
if errorlevel 1 (
    echo.
    echo O sistema foi encerrado com erro.
)
pause
exit /b 0

:erro_pip
echo.
echo ERRO ao instalar/verificar as dependencias.
echo Verifique sua conexao com a internet e tente novamente.
pause
exit /b 1
