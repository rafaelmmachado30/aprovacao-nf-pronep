@echo off
REM ============================================================
REM  Deploy do Sistema de Aprovacao de NF para Azure SWA
REM  Pre-requisitos:
REM    - instalar.bat ja executado (Node + SWA CLI)
REM    - DEPLOYMENT_TOKEN setado (veja GUIA_AZURE_SWA.md - Etapa 3)
REM ============================================================

cd /d "%~dp0"

if "%DEPLOYMENT_TOKEN%"=="" (
    echo.
    echo *** ERRO: variavel DEPLOYMENT_TOKEN nao setada ***
    echo.
    echo Pegue o token no Azure Portal:
    echo   Static Web App ^> Visao geral ^> Gerenciar token de implantacao
    echo.
    echo Depois rode no PowerShell:
    echo   $env:DEPLOYMENT_TOKEN = "seu_token_aqui"
    echo   .\04_deploy_azure.bat
    echo.
    pause
    exit /b 1
)

echo.
echo === Deploy iniciado em %date% %time% ===
echo.

REM Sobe wwwroot/ (front) + api/ (Functions)
swa deploy ./wwwroot --api-location ./api --deployment-token %DEPLOYMENT_TOKEN% --env production

if errorlevel 1 (
    echo.
    echo *** FALHA NO DEPLOY ***
    echo.
    pause
    exit /b 2
)

echo.
echo === Deploy concluido ===
echo.
echo Acesse: https://aprovacao-nf-pronep.azurestaticapps.net
echo.
pause
