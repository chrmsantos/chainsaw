@echo off
REM =============================================================================
REM CHAINSAW - Launcher para Restauração de Backup
REM =============================================================================
REM Versão: 1.0.0
REM =============================================================================

setlocal EnableDelayedExpansion EnableExtensions

REM Maximiza a janela do console
mode con cols=120 lines=50 >nul 2>&1
if not "%1"=="__MAXIMIZED__" (
    start /MAX cmd /c "%~f0" __MAXIMIZED__ %* & exit/b
)

echo.
echo ========================================================================
echo   CHAINSAW - Restauracao de Backup
echo ========================================================================
echo.

REM Verifica se restore-backup.ps1 existe
if not exist "%~dp0restore-backup.ps1" (
    echo [X] ERRO: restore-backup.ps1 nao encontrado!
    echo.
    pause
    exit /b 1
)

REM Verifica se PowerShell está disponível
where powershell.exe >nul 2>&1
if errorlevel 1 (
    echo [X] ERRO: PowerShell nao encontrado!
    echo.
    pause
    exit /b 1
)

echo [*] Executando script de restauracao...
echo.
echo ========================================================================
echo.

REM Executa restore-backup.ps1 com bypass temporário
powershell.exe -ExecutionPolicy Bypass -NoProfile -NoLogo -File "%~dp0restore-backup.ps1" %*

set EXIT_CODE=%ERRORLEVEL%

echo.
echo ========================================================================
echo.

if %EXIT_CODE% EQU 0 (
    echo [OK] Operacao concluida
) else (
    echo [X] Operacao falhou com codigo de erro: %EXIT_CODE%
)

echo.
pause

endlocal
exit /b %EXIT_CODE%
