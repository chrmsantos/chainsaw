@echo off
REM =============================================================================
REM CHAINSAW - Launcher Unificado
REM =============================================================================
REM Versao: 2.0.3
REM Este arquivo executa o script PowerShell unificado com bypass seguro.
REM =============================================================================

setlocal EnableDelayedExpansion EnableExtensions

REM Verifica se foi passado um comando
if "%~1"=="" (
    echo.
    echo ========================================================================
    echo   CHAINSAW - Sistema de Padronizacao de Proposituras Legislativas
    echo ========================================================================
    echo.
    echo [i] Uso: chainsaw.cmd ^<comando^> [opcoes]
    echo.
    echo Comandos disponiveis:
    echo   install      - Instala as configuracoes do Word
    echo   update-vba   - Atualiza apenas o modulo VBA
    echo   export       - Exporta configuracoes atuais
    echo   restore      - Restaura backup anterior
    echo   enable-vba   - Habilita acesso programatico ao VBA
    echo   disable-vba  - Desabilita acesso programatico ao VBA
    echo.
    echo Exemplos:
    echo   chainsaw.cmd install
    echo   chainsaw.cmd update-vba
    echo   chainsaw.cmd export
    echo.
    pause
    exit /b 1
)

REM Valida comando
set "VALID_COMMANDS=install update-vba export restore enable-vba disable-vba"
set "COMMAND=%~1"
set "COMMAND_VALID=0"

for %%c in (%VALID_COMMANDS%) do (
    if /I "%COMMAND%"=="%%c" set "COMMAND_VALID=1"
)

if %COMMAND_VALID%==0 (
    echo [X] Comando invalido: %COMMAND%
    echo.
    echo Comandos validos: %VALID_COMMANDS%
    pause
    exit /b 1
)

REM Verifica PowerShell
where powershell.exe >nul 2>&1
if errorlevel 1 (
    echo [X] ERRO: PowerShell nao encontrado!
    pause
    exit /b 1
)

REM Verifica se chainsaw.ps1 existe
if not exist "%~dp0chainsaw.ps1" (
    echo [X] ERRO: chainsaw.ps1 nao encontrado em %~dp0
    pause
    exit /b 1
)

REM Executa o script PowerShell com o comando
shift
powershell.exe -ExecutionPolicy Bypass -NoProfile -NoLogo -File "%~dp0chainsaw.ps1" %COMMAND% %*

set EXIT_CODE=%ERRORLEVEL%

REM Pausa apenas se executado por duplo-clique
echo %CMDCMDLINE% | find /i "%~0" >nul
if not errorlevel 1 (
    echo.
    pause
)

endlocal
exit /b %EXIT_CODE%
