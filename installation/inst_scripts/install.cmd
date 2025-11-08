@echo off
REM =============================================================================
REM CHAINSAW - Launcher Seguro para Instalação
REM =============================================================================
REM Este arquivo .cmd garante que o script PowerShell seja executado mesmo com
REM políticas restritivas de execução, usando bypass temporário seguro.
REM =============================================================================

REM Maximiza a janela do console
mode con cols=120 lines=50
if not "%1"=="max" start /MAX cmd /c %0 max & exit/b

setlocal EnableDelayedExpansion

echo.
echo ========================================================================
echo   CHAINSAW - Sistema de Padronizacao de Proposituras Legislativas
echo ========================================================================
echo.
echo [i] Launcher Seguro - Versao 1.0.0
echo.

REM Verifica se install.ps1 existe
if not exist "%~dp0install.ps1" (
    echo [X] ERRO: install.ps1 nao encontrado!
    echo.
    echo     Caminho esperado: %~dp0install.ps1
    echo.
    pause
    exit /b 1
)

echo [*] Seguranca:
echo     - Apenas o install.ps1 sera executado
echo     - A politica do sistema NAO sera alterada
echo     - O bypass expira quando o script terminar
echo     - Nenhum privilegio de administrador e usado
echo.
echo [*] Executando install.ps1 com bypass temporario seguro...
echo.

REM Executa install.ps1 com bypass temporário
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0install.ps1" %*

REM Captura o código de saída
set EXIT_CODE=%ERRORLEVEL%

echo.
if %EXIT_CODE% EQU 0 (
    echo [OK] Instalacao concluida!
) else (
    echo [X] Instalacao falhou com codigo: %EXIT_CODE%
)
echo.

REM Pausa apenas se executado por duplo-clique
echo %CMDCMDLINE% | find /i "%~0" >nul
if not errorlevel 1 (
    pause
)

exit /b %EXIT_CODE%
