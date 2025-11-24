@echo off
REM =============================================================================
REM CHAINSAW - Exportacao de Personalizacoes do Word
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
echo   CHAINSAW - Exportacao de Personalizacoes do Word
echo ========================================================================
echo.
echo [i] Export Launcher - Versao 2.0.0
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

REM Verifica se export-config.ps1 existe
if not exist "%~dp0export-config.ps1" (
    echo [X] ERRO: export-config.ps1 nao encontrado!
    echo.
    echo     Caminho esperado: %~dp0export-config.ps1
    echo     Diretorio atual:  %~dp0
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
    echo     Feche o Word antes de exportar para garantir que todas
    echo     as personalizacoes sejam salvas corretamente.
    echo.
    choice /C SN /M "Deseja continuar mesmo assim? (S/N)"
    if errorlevel 2 (
        echo.
        echo [i] Operacao cancelada pelo usuario.
        echo.
        pause
        exit /b 0
    )
    echo.
)

REM =============================================================================
REM EXECUÇÃO
REM =============================================================================

echo [*] Exportando suas personalizacoes do Word...
echo.
echo ========================================================================
echo.

REM Executa export-config.ps1 com bypass temporário
powershell.exe -ExecutionPolicy Bypass -NoProfile -NoLogo -File "%~dp0export-config.ps1" %*

REM Captura o código de saída
set EXIT_CODE=%ERRORLEVEL%

REM =============================================================================
REM RESULTADO
REM =============================================================================

echo.
echo ========================================================================
echo.

if %EXIT_CODE% EQU 0 (
    echo [OK] Exportacao concluida com sucesso!
    echo.
    echo [i] Proximos passos:
    echo     1. Copie a pasta 'exported-config' para a outra maquina
    echo     2. Coloque-a em: chainsaw\installation\
    echo     3. Execute install.cmd na maquina de destino
    echo        (o instalador detectara e importara automaticamente^)
    echo.
) else (
    echo [X] Exportacao falhou com codigo de erro: %EXIT_CODE%
    echo.
    echo [i] SOLUCOES COMUNS:
    echo     1. Feche completamente o Microsoft Word
    echo     2. Execute novamente o exportador
    echo     3. Verifique permissoes de leitura/escrita
    echo.
)

REM Pausa apenas se executado por duplo-clique
echo %CMDCMDLINE% | find /i "%~0" >nul
if not errorlevel 1 (
    echo.
    pause
)

endlocal
exit /b %EXIT_CODE%
