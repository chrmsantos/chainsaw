@echo off
REM =============================================================================
REM CHAINSAW - Exportacao de Personalizacoes do Word
REM =============================================================================

REM Maximiza a janela do console
mode con cols=120 lines=50
if not "%1"=="max" start /MAX cmd /c %0 max & exit/b

setlocal EnableDelayedExpansion

echo.
echo ========================================================================
echo   CHAINSAW - Exportacao de Personalizacoes do Word
echo ========================================================================
echo.

REM Verifica se export-config.ps1 existe
if not exist "%~dp0export-config.ps1" (
    echo [X] ERRO: export-config.ps1 nao encontrado!
    echo.
    pause
    exit /b 1
)

echo [*] Exportando suas personalizacoes do Word...
echo.

REM Executa export-config.ps1 com bypass temporário
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0export-config.ps1" %*

REM Captura o código de saída
set EXIT_CODE=%ERRORLEVEL%

echo.
if %EXIT_CODE% EQU 0 (
    echo [OK] Exportacao concluida!
    echo.
    echo Proximos passos:
    echo   1. Copie a pasta 'exported-config' para a outra maquina
    echo   2. Execute import-config.cmd na maquina de destino
) else (
    echo [X] Exportacao falhou com codigo: %EXIT_CODE%
)
echo.

REM Pausa apenas se executado por duplo-clique
echo %CMDCMDLINE% | find /i "%~0" >nul
if not errorlevel 1 (
    pause
)

exit /b %EXIT_CODE%
