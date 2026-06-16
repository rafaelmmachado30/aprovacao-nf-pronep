@echo off
REM ============================================================
REM Sistema de Aprovacao de Notas Fiscais - Pronep Life Care
REM Launcher do sistema (abre em modo app, sem barra do browser)
REM ============================================================

title Sistema de Aprovacao NF Pronep

set URL=https://purple-forest-09588fe10.7.azurestaticapps.net
set EDGE_X86="C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
set EDGE_X64="C:\Program Files\Microsoft\Edge\Application\msedge.exe"
set CHROME_X86="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
set CHROME_X64="C:\Program Files\Google\Chrome\Application\chrome.exe"

REM Tenta abrir no Edge em modo app (visual de aplicativo nativo)
if exist %EDGE_X64% (
  start "" %EDGE_X64% --app=%URL%
  goto :eof
)
if exist %EDGE_X86% (
  start "" %EDGE_X86% --app=%URL%
  goto :eof
)

REM Fallback: Chrome em modo app
if exist %CHROME_X64% (
  start "" %CHROME_X64% --app=%URL%
  goto :eof
)
if exist %CHROME_X86% (
  start "" %CHROME_X86% --app=%URL%
  goto :eof
)

REM Ultimo fallback: abre no browser padrao do sistema
start "" %URL%
