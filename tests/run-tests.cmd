@echo off
REM ================================================================
REM CHAINSAW - Executar Testes Automatizados
REM ================================================================
REM
REM Script para executar suite de testes Pester com bypass de ExecutionPolicy
REM Uso: run-tests.cmd [--detailed]
REM

setlocal
cd /d "%~dp0"

echo.
echo ========================================
echo  CHAINSAW - Testes Automatizados
echo ========================================
echo.

if "%1"=="--detailed" (
    echo Executando testes em modo detalhado...
    echo.
    powershell -NoProfile -ExecutionPolicy Bypass -File ".\Run-Tests.ps1" -Detailed
) else (
    echo Executando testes...
    echo.
    powershell -NoProfile -ExecutionPolicy Bypass -File ".\Run-Tests.ps1"
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

pause
exit /b %ERRORLEVEL%
