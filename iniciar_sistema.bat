@echo off
setlocal EnableExtensions
chcp 65001 >nul
title Sistema de Gestao - SP Aguas

cd /d "%~dp0"

echo ===============================================
echo   Sistema de Gestao - SP Aguas
echo   Modo: servidor para a rede local
echo ===============================================
echo.
echo Pasta do sistema: %CD%
echo.

if not exist "%~dp0app.py" (
    echo ERRO: app.py nao foi encontrado nesta pasta.
    pause
    exit /b 1
)
if not exist "%~dp0requirements.txt" (
    echo ERRO: requirements.txt nao foi encontrado nesta pasta.
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
    echo Criando o ambiente virtual pela primeira vez...
    %PYTHON% -m venv "%~dp0.venv"
    if errorlevel 1 (
        echo ERRO ao criar o ambiente virtual.
        pause
        exit /b 1
    )
)

echo.
echo Verificando dependencias...
"%~dp0.venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :erro_pip
"%~dp0.venv\Scripts\python.exe" -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 goto :erro_pip

echo.
echo ===============================================
echo   SISTEMA DISPONIVEL NA REDE LOCAL
echo ===============================================
echo.
echo Neste computador:
echo   http://127.0.0.1:5000
echo.
echo Enderecos IPv4 deste computador:
ipconfig | findstr /i "IPv4"
echo.
echo Para a outra pessoa acessar, use no navegador:
echo   http://IP-DESTE-COMPUTADOR:5000

echo.
echo IMPORTANTE:
echo   - Nao abra a planilha no Excel enquanto o sistema estiver gravando.
echo   - Este computador precisa permanecer ligado com esta janela aberta.
echo   - O outro usuario NAO precisa instalar o sistema; apenas abrir o navegador.
echo.
start "" http://127.0.0.1:5000

"%~dp0.venv\Scripts\python.exe" -m waitress --host=0.0.0.0 --port=5000 --threads=8 app:app

if errorlevel 1 (
    echo.
    echo O servidor foi encerrado com erro.
)
pause
exit /b 0

:erro_pip
echo.
echo ERRO ao instalar/verificar as dependencias.
echo Verifique sua conexao com a internet e tente novamente.
pause
exit /b 1
