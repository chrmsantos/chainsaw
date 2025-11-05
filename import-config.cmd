@echo off
REM =============================================================================
REM CHAINSAW - Importacao de Personalizacoes do Word
REM =============================================================================

setlocal EnableDelayedExpansion

echo.
echo ========================================================================
echo   CHAINSAW - Importacao de Personalizacoes do Word
echo ========================================================================
echo.

REM Verifica se import-config.ps1 existe
if not exist "%~dp0import-config.ps1" (
    echo [X] ERRO: import-config.ps1 nao encontrado!
    echo.
    pause
    exit /b 1
)

REM Verifica se pasta exported-config existe
if not exist "%~dp0exported-config" (
    echo [X] ERRO: Pasta 'exported-config' nao encontrada!
    echo.
    echo Certifique-se de que a pasta exportada esta neste local:
    echo %~dp0exported-config
    echo.
    pause
    exit /b 1
)

echo [!] IMPORTANTE: Feche o Microsoft Word antes de continuar!
echo.
pause

echo [*] Importando personalizacoes do Word...
echo.

REM Executa import-config.ps1 com bypass temporário
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0import-config.ps1" %*

REM Captura o código de saída
set EXIT_CODE=%ERRORLEVEL%

echo.
if %EXIT_CODE% EQU 0 (
    echo [OK] Importacao concluida!
    echo.
    echo Proximo passo:
    echo   Abra o Microsoft Word para verificar as personalizacoes
) else (
    echo [X] Importacao falhou com codigo: %EXIT_CODE%
)
echo.

REM Pausa apenas se executado por duplo-clique
echo %CMDCMDLINE% | find /i "%~0" >nul
if not errorlevel 1 (
    pause
)

exit /b %EXIT_CODE%
