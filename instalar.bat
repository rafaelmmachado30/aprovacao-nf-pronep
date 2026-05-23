@echo off
REM ============================================================
REM  Instalador de dependências — Sistema de Aprovação de NF
REM ============================================================
REM  Pré-requisitos:
REM    - Node.js LTS (https://nodejs.org)
REM    - npm (vem junto)
REM ============================================================

cd /d "%~dp0"

echo.
echo === Verificando Node.js ===
where node >nul 2>nul
if errorlevel 1 (
    echo *** Node.js nao encontrado. Instale em https://nodejs.org ***
    pause
    exit /b 1
)
node --version

echo.
echo === Instalando dependencias do backend (api) ===
cd api
call npm install
if errorlevel 1 (
    echo *** Falha ao instalar dependencias do api ***
    pause
    exit /b 2
)
cd ..

echo.
echo === Instalando Azure Static Web Apps CLI globalmente ===
where swa >nul 2>nul
if errorlevel 1 (
    call npm install -g @azure/static-web-apps-cli
)
swa --version

echo.
echo === Pronto ===
echo.
echo Proximos passos:
echo   1. Copie .env.exemplo para .env e preencha as credenciais
echo   2. Crie os grupos no Entra ID (veja GUIA_ENTRA_ID.md)
echo   3. Crie o Static Web App (veja GUIA_AZURE_SWA.md)
echo   4. Rode 04_deploy_azure.bat para o primeiro deploy
echo.
pause
