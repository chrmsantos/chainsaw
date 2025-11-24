@echo off
REM =============================================================================
REM CHAINSAW - Launcher Seguro para Instalação
REM =============================================================================
REM Versão: 2.0.0
REM Este arquivo .cmd garante que o script PowerShell seja executado mesmo com
REM políticas restritivas de execução, usando bypass temporário seguro.
REM =============================================================================

setlocal EnableDelayedExpansion EnableExtensions


REM Maximiza a janela do console (tenta, mas não falha se não conseguir)
mode con cols=120 lines=50 >nul 2>&1
if not "%1"=="__MAXIMIZED__" (
    start /MAX cmd /c "%~f0" __MAXIMIZED__ %* & exit/b
)

REM =============================================================================
REM BANNER E INFORMAÇÕES
REM =============================================================================
echo.
echo ========================================================================
echo   CHAINSAW - Sistema de Padronizacao de Proposituras Legislativas
echo ========================================================================
echo.
echo [i] Launcher Seguro - Versao 2.0.0
echo.

REM =============================================================================
REM VERIFICAÇÕES DE SISTEMA
REM =============================================================================

REM Verifica se está no Windows (redundante, mas garante)
if not "%OS%"=="Windows_NT" (
    echo [X] ERRO: Este script requer Windows NT ou superior!
    echo.
    pause
    exit /b 1
)

REM Verifica versão do Windows (Windows 7 ou superior)
REM Windows 7 = 6.1, Windows 8 = 6.2, Windows 8.1 = 6.3, Windows 10 = 10.0
for /f "tokens=4-5 delims=. " %%i in ('ver') do set VERSION=%%i.%%j

REM Extrai apenas o número principal da versão para comparação
for /f "tokens=1 delims=." %%a in ("%VERSION%") do set MAJOR_VERSION=%%a

REM Verifica se é Windows 7 (6.1) ou superior
REM Windows 10+ tem major version >= 10, Windows 7-8.1 tem major version = 6
if %MAJOR_VERSION% LSS 6 (
    echo [X] ERRO: Windows 7 ou superior e necessario!
    echo     Versao detectada: %VERSION%
    echo.
    pause
    exit /b 1
)

REM =============================================================================
REM VERIFICAÇÕES DE ARQUIVOS
REM =============================================================================

REM Verifica se install.ps1 existe no mesmo diretório
if not exist "%~dp0install.ps1" (
    echo [X] ERRO: install.ps1 nao encontrado!
    echo.
    echo     Caminho esperado: %~dp0install.ps1
    echo     Diretorio atual:  %~dp0
    echo.
    echo [i] SOLUCAO: Certifique-se de que install.ps1 esta na mesma pasta.
    echo.
    pause
    exit /b 1
)

REM Verifica se o arquivo não está corrompido (tamanho mínimo)
for %%A in ("%~dp0install.ps1") do set FILE_SIZE=%%~zA
if %FILE_SIZE% LSS 1000 (
    echo [X] ERRO: install.ps1 parece estar corrompido!
    echo     Tamanho do arquivo: %FILE_SIZE% bytes
    echo.
    pause
    exit /b 1
)

REM =============================================================================
REM VERIFICAÇÃO DO POWERSHELL
REM =============================================================================

REM Verifica se PowerShell está disponível
where powershell.exe >nul 2>&1
if errorlevel 1 (
    echo [X] ERRO: PowerShell nao encontrado no sistema!
    echo.
    echo [i] SOLUCAO: Instale o PowerShell 5.1 ou superior.
    echo     Download: https://aka.ms/wmf5download
    echo.
    pause
    exit /b 1
)

REM Verifica versão do PowerShell (5.1 ou superior)
for /f "tokens=*" %%i in ('powershell.exe -NoProfile -Command "$PSVersionTable.PSVersion.Major"') do set PS_VERSION=%%i
if %PS_VERSION% LSS 5 (
    echo [X] ERRO: PowerShell 5.1 ou superior e necessario!
    echo     Versao detectada: %PS_VERSION%.x
    echo.
    echo [i] SOLUCAO: Atualize o PowerShell para versao 5.1 ou superior.
    echo     Download: https://aka.ms/wmf5download
    echo.
    pause
    exit /b 1
)

REM =============================================================================
REM VERIFICAÇÃO DE PRIVILÉGIOS (AVISO)
REM =============================================================================

REM Verifica se está executando como administrador (apenas aviso, não bloqueia)
net session >nul 2>&1
if %errorlevel% EQU 0 (
    echo [!] AVISO: Executando como Administrador
    echo     Este script NAO requer privilegios elevados.
    echo     Recomenda-se executar como usuario normal.
    echo.
    timeout /t 3 >nul
)

REM =============================================================================
REM VERIFICAÇÃO DE WORD EM EXECUÇÃO
REM =============================================================================

REM Verifica se o Word está em execução (aviso prévio)
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>NUL | find /I /N "WINWORD.EXE" >NUL
if not errorlevel 1 (
    echo [!] AVISO: Microsoft Word esta em execucao!
    echo     O instalador solicitara o fechamento do Word.
    echo.
)

REM =============================================================================
REM INFORMAÇÕES DE SEGURANÇA
REM =============================================================================

echo [*] Seguranca:
echo     - Apenas o install.ps1 sera executado
echo     - A politica do sistema NAO sera alterada
echo     - O bypass expira quando o script terminar
echo     - Nenhum privilegio de administrador e necessario
echo     - Codigo-fonte aberto e auditavel
echo.
echo [*] Sistema:
echo     - Windows:    %VERSION%
echo     - PowerShell: %PS_VERSION%.x
echo     - Diretorio:  %~dp0
echo.

REM =============================================================================
REM PREPARAÇÃO PARA EXECUÇÃO
REM =============================================================================

REM Cria timestamp para log
for /f "tokens=2-4 delims=/ " %%a in ('date /t') do set LOGDATE=%%c%%b%%a
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set LOGTIME=%%a%%b
set TIMESTAMP=%LOGDATE%_%LOGTIME%

REM Define arquivo de log de erro
set ERROR_LOG="%TEMP%\chainsaw_install_error_%TIMESTAMP%.log"

echo [*] Executando instalacao...
echo.
echo ========================================================================
echo.

REM =============================================================================
REM EXECUÇÃO DO POWERSHELL
REM =============================================================================

REM Executa install.ps1 com bypass temporário e captura de erros
powershell.exe -ExecutionPolicy Bypass -NoProfile -NoLogo -File "%~dp0install.ps1" %* 2>%ERROR_LOG%

REM Captura o código de saída
set EXIT_CODE=%ERRORLEVEL%

REM =============================================================================
REM TRATAMENTO DE RESULTADO
REM =============================================================================

echo.
echo ========================================================================
echo.

if %EXIT_CODE% EQU 0 (
    echo [OK] Instalacao concluida com sucesso!
    echo.
    
    REM Remove log de erro se não houve erro
    if exist %ERROR_LOG% del /F /Q %ERROR_LOG% >nul 2>&1
) else (
    echo [X] Instalacao falhou com codigo de erro: %EXIT_CODE%
    echo.
    
    REM Verifica se há log de erro
    if exist %ERROR_LOG% (
        for %%A in (%ERROR_LOG%) do set ERROR_SIZE=%%~zA
        if !ERROR_SIZE! GTR 0 (
            echo [i] Detalhes do erro salvos em:
            echo     %ERROR_LOG%
            echo.
            echo [i] Primeiras linhas do erro:
            echo     ----------------------------------------------------------------
            type %ERROR_LOG% | more /E +0
            echo     ----------------------------------------------------------------
            echo.
        )
    )
    
    echo [i] SOLUCOES COMUNS:
    echo     1. Verifique se o Word esta fechado
    echo     2. Execute novamente o instalador
    echo     3. Consulte o log para mais detalhes
    echo     4. Verifique permissoes de escrita no seu perfil
    echo.
)

REM =============================================================================
REM LIMPEZA E SAÍDA
REM =============================================================================

REM Pausa apenas se executado por duplo-clique (não em linha de comando)
echo %CMDCMDLINE% | find /i "%~0" >nul
if not errorlevel 1 (
    echo.
    pause
)

REM Restaura codepage original (se necessário)
REM chcp 850 >nul 2>&1

endlocal
exit /b %EXIT_CODE%
