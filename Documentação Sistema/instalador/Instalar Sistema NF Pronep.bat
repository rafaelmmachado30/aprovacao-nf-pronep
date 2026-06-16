@echo off
REM ====================================================================
REM Instalador do Sistema de Aprovacao NF Pronep
REM   - Copia o launcher e o icone pra pasta do usuario
REM   - Cria atalho no Desktop e no Menu Iniciar com icone Pronep
REM   - Abre o sistema ao terminar
REM ====================================================================

title Instalando Sistema NF Pronep
color 0B
cls
echo.
echo   ============================================================
echo     SISTEMA DE APROVACAO DE NF - PRONEP LIFE CARE
echo     Instalador v1.0
echo   ============================================================
echo.
echo   Vou copiar 2 arquivos pra sua pasta de usuario e criar
echo   um atalho no Desktop e no Menu Iniciar com o icone Pronep.
echo.
echo   Voce nao precisa de permissao de administrador.
echo.
pause

set INSTALL_DIR=%USERPROFILE%\Pronep-NF
set BAT_FILE=%INSTALL_DIR%\Sistema-NF-Pronep.bat
set ICO_FILE=%INSTALL_DIR%\Pronep-NF.ico
set DESKTOP=%USERPROFILE%\Desktop
set STARTMENU=%APPDATA%\Microsoft\Windows\Start Menu\Programs

echo.
echo   [1/4] Criando pasta de instalacao...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo   [2/4] Copiando launcher e icone...
copy /Y "%~dp0Sistema-NF-Pronep.bat" "%BAT_FILE%" >nul
copy /Y "%~dp0Pronep-NF.ico" "%ICO_FILE%" >nul

echo   [3/4] Criando atalho no Desktop...
powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell;" ^
  "$sc = $ws.CreateShortcut('%DESKTOP%\Sistema NF Pronep.lnk');" ^
  "$sc.TargetPath = '%BAT_FILE%';" ^
  "$sc.IconLocation = '%ICO_FILE%';" ^
  "$sc.Description = 'Sistema de Aprovacao de Notas Fiscais - Pronep Life Care';" ^
  "$sc.WindowStyle = 7;" ^
  "$sc.Save();"

echo   [4/4] Criando atalho no Menu Iniciar...
powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell;" ^
  "$sc = $ws.CreateShortcut('%STARTMENU%\Sistema NF Pronep.lnk');" ^
  "$sc.TargetPath = '%BAT_FILE%';" ^
  "$sc.IconLocation = '%ICO_FILE%';" ^
  "$sc.Description = 'Sistema de Aprovacao de Notas Fiscais - Pronep Life Care';" ^
  "$sc.WindowStyle = 7;" ^
  "$sc.Save();"

echo.
echo   ============================================================
echo     INSTALACAO CONCLUIDA COM SUCESSO!
echo   ============================================================
echo.
echo   Atalhos criados em:
echo     - Desktop
echo     - Menu Iniciar (procure por "Sistema NF Pronep")
echo.
echo   Pra usar: clique duas vezes no atalho.
echo.
echo   Abrindo o sistema agora...
echo.
timeout /t 2 /nobreak >nul

start "" "%BAT_FILE%"
exit
