@echo off
REM =============================================================================
REM CHAINSAW - Launcher para Atualização do Módulo VBA
REM =============================================================================
REM Versão: 2.0.0
REM Licença: GNU GPLv3
REM =============================================================================

setlocal EnableDelayedExpansion EnableExtensions


REM Maximiza a janela do console (tenta, mas não falha se não conseguir)
mode con cols=120 lines=50 >nul 2>&1
if not "%1"=="__MAXIMIZED__" (
    start /MAX cmd /c "%~f0" __MAXIMIZED__ %* & exit/b
)

REM =============================================================================
REM BANNER
REM =============================================================================
echo.
echo ========================================================================
echo   CHAINSAW - Atualizacao de Modulo VBA
echo ========================================================================
echo.
echo [i] Update Launcher - Versao 2.0.0
echo.

REM =============================================================================
REM VERIFICAÇÕES
REM =============================================================================

REM Verifica se está no Windows
if not "%OS%"=="Windows_NT" (
    echo [X] ERRO: Este script requer Windows NT ou superior!
    echo.
    pause
    exit /b 1
)

REM Define o caminho do script PowerShell
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%update-vba-module.ps1"

REM Verifica se o script PowerShell existe
if not exist "%PS_SCRIPT%" (
    echo [X] ERRO: Script nao encontrado!
    echo.
    echo     Caminho esperado: %PS_SCRIPT%
    echo     Diretorio atual:  %SCRIPT_DIR%
    echo.
    pause
    exit /b 1
)

REM Verifica se PowerShell está disponível
where powershell.exe >nul 2>&1
if errorlevel 1 (
    echo [X] ERRO: PowerShell nao encontrado no sistema!
    echo.
    pause
    exit /b 1
)

REM Verifica se o Word está em execução (aviso)
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>NUL | find /I /N "WINWORD.EXE" >NUL
if not errorlevel 1 (
    echo [!] AVISO: Microsoft Word esta em execucao!
    echo     O atualizador solicitara o fechamento do Word.
    echo.
)

REM =============================================================================
REM EXECUÇÃO
REM =============================================================================

echo [*] Iniciando atualizacao do modulo VBA...
echo.
echo ========================================================================
echo.

REM Executa o script PowerShell com bypass de execução
powershell.exe -NoProfile -NoLogo -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*

REM Captura o código de saída
set "EXIT_CODE=%ERRORLEVEL%"

REM =============================================================================
REM RESULTADO
REM =============================================================================

echo.
echo ========================================================================
echo.

if !EXIT_CODE! EQU 0 (
    echo [OK] Modulo VBA atualizado com sucesso!
    echo.
) else (
    echo [X] Falha na atualizacao. Codigo de erro: !EXIT_CODE!
    echo.
    echo [i] SOLUCOES COMUNS:
    echo     1. Feche completamente o Microsoft Word
    echo     2. Execute novamente o atualizador
    echo     3. Verifique se tem permissoes de escrita
    echo.
)

REM Pausa apenas se executado por duplo-clique
echo %CMDCMDLINE% | find /i "%~0" >nul
if not errorlevel 1 (
    echo.
    pause
)

endlocal
exit /b !EXIT_CODE!
