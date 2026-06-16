@echo off
REM Aplica versao limpa de contratos-bfs.js (com iniciarPath dentro do IIFE)
SET REPO=C:\Pronep\Aprovacao_NF
SET ORIGEM=%REPO%\contratos-bfs-novo.js
SET DESTINO=%REPO%\wwwroot\contratos-bfs.js

echo === Verificando origem ===
if not exist "%ORIGEM%" (
  echo ERRO: arquivo nao encontrado em %ORIGEM%
  pause
  exit /b 1
)

echo === Copiando ===
copy /Y "%ORIGEM%" "%DESTINO%"
if errorlevel 1 ( echo ERRO ao copiar & pause & exit /b 1 )

echo === Validando sintaxe ===
cd /d "%REPO%\wwwroot"
node -c contratos-bfs.js
if errorlevel 1 ( echo ERRO de sintaxe! & pause & exit /b 1 )

echo === Git commit + push ===
cd /d "%REPO%"
git add wwwroot/contratos-bfs.js wwwroot/index.html
git commit -m "fix: contratos-bfs.js limpo (iniciarPath dentro do IIFE) + botao BFS no modal"
git push origin main

echo === Limpando temporario ===
del "%ORIGEM%"

echo === Pronto. Aguarde deploy e teste o botao Sincronizar via BFS no modal ===
pause
