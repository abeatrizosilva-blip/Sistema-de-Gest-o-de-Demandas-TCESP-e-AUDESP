@echo off
chcp 65001 >nul
title Liberar acesso local - SP Aguas

echo Este arquivo cria uma regra no Firewall do Windows para permitir

echo conexoes de entrada na porta 5000 da rede local.
echo.
netsh advfirewall firewall add rule name="SP Aguas - Servidor Local 5000" dir=in action=allow protocol=TCP localport=5000 profile=domain,private
if errorlevel 1 (
    echo.
    echo NAO FOI POSSIVEL CRIAR A REGRA.
    echo Execute este arquivo como Administrador ou solicite ao TI.
) else (
    echo.
    echo REGRA CRIADA COM SUCESSO.
    echo Agora execute iniciar_sistema.bat.
)
pause
