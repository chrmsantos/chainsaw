@echo off
REM =============================================================================
REM CHAINSAW - Launcher para Atualização do Módulo VBA
REM =============================================================================
REM Versão: 1.0.0
REM Licença: GNU GPLv3
REM =============================================================================

setlocal enabledelayedexpansion

REM Define o caminho do script PowerShell
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%update-vba-module.ps1"

REM Verifica se o script existe
if not exist "%PS_SCRIPT%" (
    echo [ERRO] Script nao encontrado: %PS_SCRIPT%
    pause
    exit /b 1
)

REM Executa o script PowerShell com bypass de execução
echo Iniciando atualizacao do modulo VBA...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*

REM Captura o código de saída
set "EXIT_CODE=%ERRORLEVEL%"

REM Pausa apenas se houver erro (código diferente de 0)
if !EXIT_CODE! neq 0 (
    echo.
    echo [ERRO] Falha na atualizacao. Codigo: !EXIT_CODE!
    pause
)

exit /b !EXIT_CODE!
