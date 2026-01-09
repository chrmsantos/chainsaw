@echo off
REM ================================================================
REM CHAINSAW - Executar Testes Automatizados
REM ================================================================
REM
REM Script para executar suite de testes Pester com bypass de ExecutionPolicy
REM Uso: run-tests.cmd [--detailed] [--no-pause]
REM

setlocal
cd /d "%~dp0"

echo.
echo ========================================
echo  CHAINSAW - Testes Automatizados
echo ========================================
echo.

set DETAILED=
set NOPAUSE=

if /I "%1"=="--detailed" set DETAILED=1
if /I "%2"=="--detailed" set DETAILED=1

if /I "%1"=="--no-pause" set NOPAUSE=1
if /I "%2"=="--no-pause" set NOPAUSE=1

if defined DETAILED (
    echo Executando testes em modo detalhado...
    echo.
    powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ".\Run-Tests.ps1" -Detailed -NoProgress
) else (
    echo Executando testes...
    echo.
    powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ".\Run-Tests.ps1" -Output Minimal -NoProgress
)

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [OK] Todos os testes passaram!
    echo.
) else (
    echo.
    echo [ERRO] Alguns testes falharam. Veja detalhes acima.
    echo.
)

REM Evita travar dentro do VS Code (integrated terminal) esperando tecla.
if not defined NOPAUSE (
    if /I "%TERM_PROGRAM%"=="vscode" (
        REM no-op
    ) else (
        pause
    )
)
exit /b %ERRORLEVEL%
