@echo off
REM =============================================================================
REM CHAINSAW - Instalador Automático
REM =============================================================================
REM Baixa o código-fonte completo do GitHub e executa a instalação
REM =============================================================================

setlocal enabledelayedexpansion

REM Configurações
set "REPO_URL=https://github.com/chrmsantos/chainsaw/archive/refs/heads/main.zip"
set "INSTALL_DIR=%USERPROFILE%\chainsaw"
set "TEMP_ZIP=%TEMP%\chainsaw-main.zip"
set "TEMP_EXTRACT=%TEMP%\chainsaw-extract"
set "SAFE_INSTALL_DIR=%TEMP%\chainsaw-install-temp"

REM Configuração de logs
for /f "tokens=2-4 delims=/ " %%a in ('date /t') do set "DATESTAMP=%%c%%b%%a"
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set "TIMESTAMP=%%a%%b"
set "DATETIME=%DATESTAMP%_%TIMESTAMP%"
set "DATETIME=%DATETIME::=%"

REM Log no mesmo diretório do chainsaw_installer.cmd
set "SCRIPT_DIR=%~dp0"
set "LOG_FILE=%SCRIPT_DIR%chainsaw_installer_%DATETIME%.log"

REM Função para log (redireciona saída para arquivo e console)
call :LogInit

REM =============================================================================
REM PROTECAO CRITICA: Detecta se esta sendo executado de dentro da pasta destino
REM =============================================================================

set "CURRENT_DIR=%CD%"
set "NORMALIZED_CURRENT=%CURRENT_DIR%"
set "NORMALIZED_INSTALL=%INSTALL_DIR%"

REM Normaliza os caminhos (remove trailing backslash)
if "%NORMALIZED_CURRENT:~-1%"=="\" set "NORMALIZED_CURRENT=%NORMALIZED_CURRENT:~0,-1%"
if "%NORMALIZED_INSTALL:~-1%"=="\" set "NORMALIZED_INSTALL=%NORMALIZED_INSTALL:~0,-1%"

REM Compara se esta dentro da pasta de destino
echo %NORMALIZED_CURRENT% | findstr /I /C:"%NORMALIZED_INSTALL%" >nul
if not errorlevel 1 (
    REM Cria diretorio temporario
    if not exist "%SAFE_INSTALL_DIR%" mkdir "%SAFE_INSTALL_DIR%"
    
    REM Copia o instalador para local seguro
    copy "%~f0" "%SAFE_INSTALL_DIR%\chainsaw_installer.cmd" >nul
    
    REM Executa de la
    cd /d "%SAFE_INSTALL_DIR%"
    call "%SAFE_INSTALL_DIR%\chainsaw_installer.cmd"
    
    REM Limpa e sai
    cd /d "%USERPROFILE%"
    rd /s /q "%SAFE_INSTALL_DIR%" >nul 2>&1
    exit /b 0
)

echo.
call :Log "================================================================================"
call :Log "  CHAINSAW - Instalador Automatico"
call :Log "================================================================================"
echo.
call :Log " Este instalador ira:"
call :Log "  1. Baixar o codigo-fonte completo do GitHub"
call :Log "  2. Extrair para: %INSTALL_DIR%"
call :Log "  3. Executar a instalacao automaticamente"
echo.
call :Log "================================================================================"
call :Log " Log sendo salvo em: %LOG_FILE%"
call :Log "================================================================================"
echo.

REM Verifica se PowerShell está disponível
where powershell >nul 2>&1
if errorlevel 1 (
    call :Log "[ERRO] PowerShell nao encontrado!"
    call :Log "       Este instalador requer PowerShell 5.1 ou superior."
    pause
    exit /b 1
)

echo.
call :Log "================================================================================"
call :Log "  ETAPA 1: Download do Codigo-Fonte"
call :Log "================================================================================"
echo.
call :Log " Baixando de: %REPO_URL%"
call :Log " Destino temporario: %TEMP_ZIP%"
echo.

REM Remove arquivo ZIP antigo se existir
if exist "%TEMP_ZIP%" del /f /q "%TEMP_ZIP%" >nul 2>&1

REM Baixa o arquivo ZIP usando PowerShell
set "PS_SCRIPT=%TEMP%\chainsaw_download.ps1"
(
echo $ProgressPreference='SilentlyContinue'
echo try {
echo     Write-Host '[INFO] Iniciando download...' -ForegroundColor Cyan
echo     Invoke-WebRequest -Uri '%REPO_URL%' -OutFile '%TEMP_ZIP%' -UseBasicParsing
echo     Write-Host '[OK] Download concluido!' -ForegroundColor Green
echo     exit 0
echo } catch {
echo     Write-Host "[ERRO] Falha no download: $_" -ForegroundColor Red
echo     exit 1
echo }
) > "%PS_SCRIPT%"

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
set DOWNLOAD_EXIT=%ERRORLEVEL%
del /f /q "%PS_SCRIPT%" >nul 2>&1

if %DOWNLOAD_EXIT% neq 0 (
    echo.
    call :Log "[ERRO] Falha ao baixar o codigo-fonte do GitHub."
    call :Log "       Verifique sua conexao com a internet e tente novamente."
    pause
    exit /b 1
)

REM Verifica se o arquivo foi baixado
if not exist "%TEMP_ZIP%" (
    call :Log "[ERRO] Arquivo ZIP nao encontrado apos download."
    pause
    exit /b 1
)

call :Log "[OK] Download verificado com sucesso."

REM =============================================================================
REM ETAPA 1.5: VALIDAÇÃO COMPLETA DO ARQUIVO BAIXADO
REM =============================================================================
echo.
call :Log "================================================================================"
call :Log "  ETAPA 1.5: Validacao do Arquivo Baixado"
call :Log "================================================================================"
echo.

REM Verifica tamanho mínimo (arquivo válido deve ter pelo menos 100KB)
for %%A in ("%TEMP_ZIP%") do set "ZIP_SIZE=%%~zA"
call :Log "[INFO] Tamanho do arquivo ZIP: %ZIP_SIZE% bytes"

if %ZIP_SIZE% LSS 102400 (
    call :Log "[ERRO] Arquivo ZIP muito pequeno (possivelmente corrompido)"
    call :Log "       Tamanho esperado: > 100KB, recebido: %ZIP_SIZE% bytes"
    del /f /q "%TEMP_ZIP%" >nul 2>&1
    pause
    exit /b 1
)

call :Log "[OK] Tamanho do arquivo validado."

REM Testa a integridade do ZIP ANTES de fazer qualquer modificação
call :Log "[INFO] Testando integridade do arquivo ZIP..."
set "PS_TEST=%TEMP%\chainsaw_test_zip.ps1"
(
echo $ProgressPreference='SilentlyContinue'
echo try {
echo     Add-Type -AssemblyName System.IO.Compression.FileSystem
echo     $zip = [System.IO.Compression.ZipFile]::OpenRead('%TEMP_ZIP%'^)
echo     $entryCount = $zip.Entries.Count
echo     $zip.Dispose(^)
echo     if ($entryCount -lt 10^) {
echo         Write-Host "[ERRO] ZIP contem muito poucos arquivos: $entryCount" -ForegroundColor Red
echo         exit 1
echo     }
echo     Write-Host "[OK] ZIP valido com $entryCount arquivos" -ForegroundColor Green
echo     exit 0
echo } catch {
echo     Write-Host "[ERRO] ZIP corrompido ou invalido: $_" -ForegroundColor Red
echo     exit 1
echo }
) > "%PS_TEST%"

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_TEST%"
set "ZIP_TEST_EXIT=%ERRORLEVEL%"
del /f /q "%PS_TEST%" >nul 2>&1

if %ZIP_TEST_EXIT% neq 0 (
    call :Log "[ERRO] Arquivo ZIP corrompido ou invalido!"
    call :Log "       Nao e seguro continuar. Abortando instalacao."
    del /f /q "%TEMP_ZIP%" >nul 2>&1
    pause
    exit /b 1
)

call :Log "[OK] Integridade do ZIP validada com sucesso!"

echo.
call :Log "================================================================================"
call :Log "  ETAPA 2: Extracao e Validacao dos Arquivos"
call :Log "================================================================================"
echo.
call :Log "[INFO] Preparando extracao dos arquivos..."

REM Remove pasta temporária de extração se existir
if exist "%TEMP_EXTRACT%" rd /s /q "%TEMP_EXTRACT%" >nul 2>&1

REM Extrai o ZIP usando PowerShell
set "PS_SCRIPT=%TEMP%\chainsaw_extract.ps1"
(
echo $ProgressPreference='SilentlyContinue'
echo try {
echo     Write-Host '[INFO] Extraindo arquivos...' -ForegroundColor Cyan
echo     Expand-Archive -Path '%TEMP_ZIP%' -DestinationPath '%TEMP_EXTRACT%' -Force
echo     Write-Host '[OK] Extracao concluida!' -ForegroundColor Green
echo     exit 0
echo } catch {
echo     Write-Host "[ERRO] Falha na extracao: $_" -ForegroundColor Red
echo     exit 1
echo }
) > "%PS_SCRIPT%"

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
set EXTRACT_EXIT=%ERRORLEVEL%
del /f /q "%PS_SCRIPT%" >nul 2>&1

if %EXTRACT_EXIT% neq 0 (
    call :Log "[ERRO] Falha ao extrair o arquivo ZIP."
    pause
    exit /b 1
)

REM =============================================================================
REM VALIDAÇÃO CRÍTICA DO CONTEÚDO EXTRAÍDO
REM =============================================================================

REM O GitHub cria uma pasta "chainsaw-main" dentro do ZIP
set "SOURCE_DIR=%TEMP_EXTRACT%\chainsaw-main"

call :Log "[CRITICO] Validando conteudo extraido ANTES de instalar..."

if not exist "%SOURCE_DIR%" (
    call :Log "[ERRO CRITICO] Pasta extraida nao encontrada em: %SOURCE_DIR%"
    call :Log "[ERRO] Estrutura do ZIP inesperada. Abortando instalacao."
    if exist "%TEMP_EXTRACT%" rd /s /q "%TEMP_EXTRACT%" >nul 2>&1
    pause
    exit /b 1
)

REM Valida presença de pastas essenciais
set "VALIDATION_FAILED=0"

if not exist "%SOURCE_DIR%\installation" (
    call :Log "[ERRO] Pasta 'installation' nao encontrada no conteudo extraido!"
    set "VALIDATION_FAILED=1"
)

if not exist "%SOURCE_DIR%\installation\inst_scripts" (
    call :Log "[ERRO] Pasta 'installation\inst_scripts' nao encontrada!"
    set "VALIDATION_FAILED=1"
)

if not exist "%SOURCE_DIR%\installation\inst_scripts\install.cmd" (
    call :Log "[ERRO] Script 'install.cmd' nao encontrado!"
    set "VALIDATION_FAILED=1"
)

if not exist "%SOURCE_DIR%\installation\inst_scripts\install.ps1" (
    call :Log "[ERRO] Script 'install.ps1' nao encontrado!"
    set "VALIDATION_FAILED=1"
)

if not exist "%SOURCE_DIR%\installation\inst_configs" (
    call :Log "[ERRO] Pasta 'installation\inst_configs' nao encontrada!"
    set "VALIDATION_FAILED=1"
)

if %VALIDATION_FAILED% equ 1 (
    call :Log "[ERRO CRITICO] Conteudo extraido INVALIDO ou INCOMPLETO!"
    call :Log "[ERRO] NAO E SEGURO instalar arquivos incompletos."
    call :Log "[ERRO] Instalacao ABORTADA para proteger sua instalacao atual."
    echo.
    if exist "%TEMP_EXTRACT%" rd /s /q "%TEMP_EXTRACT%" >nul 2>&1
    pause
    exit /b 1
)

call :Log "[OK] Conteudo extraido validado com sucesso!"

REM Conta arquivos no conteúdo extraído usando PowerShell (mais rápido)
for /f %%i in ('powershell -NoProfile -Command "(Get-ChildItem -Path '%SOURCE_DIR%' -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count"') do set EXTRACTED_FILE_COUNT=%%i

call :Log "[INFO] Total de arquivos extraidos: %EXTRACTED_FILE_COUNT%"

if %EXTRACTED_FILE_COUNT% LSS 20 (
    call :Log "[ERRO] Conteudo extraido contem muito poucos arquivos: %EXTRACTED_FILE_COUNT%"
    call :Log "[ERRO] Download pode estar incompleto. Abortando."
    if exist "%TEMP_EXTRACT%" rd /s /q "%TEMP_EXTRACT%" >nul 2>&1
    pause
    exit /b 1
)

call :Log "[OK] Validacao completa! Seguro para instalar."

REM =============================================================================
REM Move os arquivos validados para o destino final
REM =============================================================================

echo.
call :Log "[INFO] Preparando instalacao em %INSTALL_DIR%..."

if exist "%INSTALL_DIR%" (
    REM PROTECAO CRITICA: Preservar pasta .git se existir
    set "GIT_BACKUP_DIR=%TEMP%\chainsaw-git-backup-%RANDOM%"
    if exist "%INSTALL_DIR%\.git" (
        call :Log "[CRITICO] Detectada pasta .git - criando backup temporario..."
        mkdir "%GIT_BACKUP_DIR%" >nul 2>&1
        xcopy "%INSTALL_DIR%\.git" "%GIT_BACKUP_DIR%\.git\" /E /H /C /I /Y >nul 2>&1
        if exist "%GIT_BACKUP_DIR%\.git" (
            call :Log "[OK] Backup do .git criado em: %GIT_BACKUP_DIR%"
        ) else (
            call :Log "[ERRO] Falha ao criar backup do .git!"
        )
    )
    
    REM Deletar conteúdo da pasta (EXCETO .git)
    call :Log "[INFO] Removendo conteudo antigo (preservando .git)..."
    
    REM Remove apenas arquivos na raiz
    del /f /q "%INSTALL_DIR%\*.*" >nul 2>&1
    
    REM Remove pastas EXCETO .git
    for /d %%p in ("%INSTALL_DIR%\*") do (
        if /I not "%%~nxp"==".git" (
            rd /s /q "%%p" >nul 2>&1
        )
    )
    
    call :Log "[OK] Conteudo removido (exceto .git)"
)

call :Log "[INFO] Instalando novos arquivos validados em %INSTALL_DIR%..."

REM Cria pasta de destino se não existir
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

REM Copia todos os arquivos e pastas
xcopy "%SOURCE_DIR%\*" "%INSTALL_DIR%\" /E /H /C /I /Y >nul
set "COPY_EXIT=%ERRORLEVEL%"

if %COPY_EXIT% neq 0 (
    call :Log "[ERRO CRITICO] Falha ao copiar arquivos para o destino (erro %COPY_EXIT%)!"
    call :Log "[ERRO] Instalacao nao pode ser concluida."
    pause
    exit /b 1
)

call :Log "[OK] Arquivos copiados com sucesso!"

REM Restaura backup do .git se existir
if exist "%GIT_BACKUP_DIR%\.git" (
    call :Log "[INFO] Restaurando backup do .git..."
    xcopy "%GIT_BACKUP_DIR%\.git" "%INSTALL_DIR%\.git\" /E /H /C /I /Y >nul 2>&1
    if exist "%INSTALL_DIR%\.git" (
        call :Log "[OK] .git restaurado com sucesso!"
    ) else (
        call :Log "[ERRO] Falha ao restaurar .git - backup em: %GIT_BACKUP_DIR%"
    )
    REM Limpa backup temporário
    rd /s /q "%GIT_BACKUP_DIR%" >nul 2>&1
)

REM =============================================================================
REM VALIDAÇÃO FINAL DA INSTALAÇÃO
REM =============================================================================

call :Log "[INFO] Validando instalacao final..."

set "FINAL_VALIDATION_FAILED=0"

if not exist "%INSTALL_DIR%\installation\inst_scripts\install.cmd" (
    call :Log "[ERRO] install.cmd nao encontrado apos instalacao!"
    set "FINAL_VALIDATION_FAILED=1"
)

if not exist "%INSTALL_DIR%\installation\inst_scripts\install.ps1" (
    call :Log "[ERRO] install.ps1 nao encontrado apos instalacao!"
    set "FINAL_VALIDATION_FAILED=1"
)

if %FINAL_VALIDATION_FAILED% equ 1 (
    call :Log "[ERRO CRITICO] Validacao final FALHOU!"
    call :Log "[ERRO] Instalacao nao pode ser concluida."
    pause
    exit /b 1
)

call :Log "[OK] Validacao final concluida com sucesso!"

REM Cria arquivo de versão local
call :Log "[INFO] Criando arquivo de versao local..."
set "LOCAL_VERSION_FILE=%INSTALL_DIR%\installation\inst_configs\version.json"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$version = @{" ^
    "  version = '2.0.2';" ^
    "  installedDate = (Get-Date -Format 'yyyy-MM-dd');" ^
    "  installPath = '%INSTALL_DIR%';" ^
    "  lastUpdateCheck = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" ^
    "} | ConvertTo-Json -Depth 3 | Out-File -FilePath '%LOCAL_VERSION_FILE%' -Encoding UTF8"
if not errorlevel 1 (
    call :Log "[OK] Arquivo de versao criado: %LOCAL_VERSION_FILE%"
)

REM Limpeza de arquivos temporários
call :Log "[INFO] Limpando arquivos temporarios..."
if exist "%TEMP_ZIP%" del /f /q "%TEMP_ZIP%" >nul 2>&1
if exist "%TEMP_EXTRACT%" rd /s /q "%TEMP_EXTRACT%" >nul 2>&1

REM Copia o log para a pasta de instalação
set "INSTALL_LOG_DIR=%INSTALL_DIR%\installation\inst_docs\inst_logs"
if not exist "%INSTALL_LOG_DIR%" mkdir "%INSTALL_LOG_DIR%"
copy "%LOG_FILE%" "%INSTALL_LOG_DIR%\chainsaw_installer_%DATETIME%.log" >nul 2>&1
if not errorlevel 1 (
    call :Log "[INFO] Log copiado para: %INSTALL_LOG_DIR%"
)

echo.
call :Log "================================================================================"
call :Log "  ETAPA 4: Executando Instalacao"
call :Log "================================================================================"
echo.

REM Verifica se install.cmd existe
set "INSTALL_CMD=%INSTALL_DIR%\installation\inst_scripts\install.cmd"
if not exist "%INSTALL_CMD%" (
    call :Log "[ERRO] Script de instalacao nao encontrado em:"
    call :Log "       %INSTALL_CMD%"
    echo.
    call :Log "       Verifique se o download foi concluido corretamente."
    pause
    exit /b 1
)

REM Executa o instalador
call :Log "[INFO] Iniciando instalacao do CHAINSAW..."
echo.

cd /d "%INSTALL_DIR%\installation\inst_scripts"
call install.cmd

REM Captura o código de saída do instalador
set INSTALL_EXIT_CODE=%errorlevel%

echo.
call :Log "================================================================================"
call :Log "  Instalacao Concluida"
call :Log "================================================================================"
echo.

if %INSTALL_EXIT_CODE% equ 0 (
    call :Log "[OK] Instalacao concluida com sucesso!"
    echo.
    call :Log " O CHAINSAW foi instalado em: %INSTALL_DIR%"
    call :Log " Log salvo em: %LOG_FILE%"
    call :Log " Log copiado para: %INSTALL_LOG_DIR%"
    echo.
) else (
    call :Log "[AVISO] A instalacao foi concluida com codigo de saida: %INSTALL_EXIT_CODE%"
    call :Log "        Verifique os logs para mais detalhes."
    echo.
)

call :Log "Processo finalizado em: %DATE% %TIME%"
echo.
echo Pressione qualquer tecla para sair...
pause >nul

exit /b %INSTALL_EXIT_CODE%

REM =============================================================================
REM Funções auxiliares
REM =============================================================================

:LogInit
REM Inicializa o arquivo de log
echo ================================================================================ > "%LOG_FILE%"
echo  CHAINSAW - Log de Instalacao >> "%LOG_FILE%"
echo ================================================================================ >> "%LOG_FILE%"
echo  Data/Hora de inicio: %DATE% %TIME% >> "%LOG_FILE%"
echo  Sistema: %OS% >> "%LOG_FILE%"
echo  Usuario: %USERNAME% >> "%LOG_FILE%"
echo  Diretorio de instalacao: %INSTALL_DIR% >> "%LOG_FILE%"
echo ================================================================================ >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"
goto :eof

:Log
REM Escreve mensagem no console e no arquivo de log
set "MSG=%~1"
echo %MSG%
echo %MSG% >> "%LOG_FILE%"
goto :eof
