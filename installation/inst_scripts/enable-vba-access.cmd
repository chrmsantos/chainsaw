@echo off
REM =============================================================================
REM CHAINSAW - Habilitar Acesso Programatico ao VBA
REM =============================================================================
REM Versao: 1.0.0
REM Licenca: GNU GPLv3
REM Autor: Christian Martin dos Santos
REM =============================================================================

setlocal enabledelayedexpansion

REM Muda para o diretório do script
cd /d "%~dp0"

REM Executa o PowerShell com política de execução bypass
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0enable-vba-access.ps1" %*

REM Captura o código de saída
set EXITCODE=%ERRORLEVEL%

REM Retorna o código de saída
exit /b %EXITCODE%
