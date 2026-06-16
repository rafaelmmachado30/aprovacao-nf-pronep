@echo off
REM Aplica versão limpa do SincronizarContratos (sem lixo após persistir)
SET REPO=C:\Pronep\Aprovacao_NF
SET ORIGEM=%REPO%\SincronizarContratos-novo.js
SET DESTINO=%REPO%\api\SincronizarContratos\index.js

echo === Verificando arquivo de origem ===
if not exist "%ORIGEM%" (
  echo ERRO: arquivo nao encontrado em %ORIGEM%
  pause
  exit /b 1
)

echo === Copiando SincronizarContratos/index.js ===
copy /Y "%ORIGEM%" "%DESTINO%"
if errorlevel 1 ( echo ERRO ao copiar & pause & exit /b 1 )

echo === Validando sintaxe ===
cd /d "%REPO%\api"
node -c SincronizarContratos\index.js
if errorlevel 1 ( echo ERRO de sintaxe! Aborte. & pause & exit /b 1 )

echo === Git commit + push ===
cd /d "%REPO%"
git add api/SincronizarContratos/index.js wwwroot/contratos-bfs.js wwwroot/index.html api/ListarSubpastasContratos
git commit -m "fix: remove lixo apos persistir + ListarSubpastas + contratos-bfs.js + index.html ajustes"
git push origin main

echo === Limpando arquivo temporario ===
del "%ORIGEM%"

echo === Pronto. Aguarde deploy verde e teste o Sincronizar TUDO (BFS) ===
pause
