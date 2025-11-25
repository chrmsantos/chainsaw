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

echo.
call :Log "================================================================================"
call :Log "  CHAINSAW - Instalador Automatico"
call :Log "================================================================================"
echo.
call :Log " Este instalador ira:"
call :Log "  1. Baixar o codigo-fonte completo do GitHub"
call :Log "  2. Criar backup da instalacao existente (se houver)"
call :Log "  3. Extrair para: %INSTALL_DIR%"
call :Log "  4. Executar a instalacao automaticamente"
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
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ProgressPreference = 'SilentlyContinue'; " ^
    "try { " ^
    "    Write-Host '[INFO] Iniciando download...' -ForegroundColor Cyan; " ^
    "    Invoke-WebRequest -Uri '%REPO_URL%' -OutFile '%TEMP_ZIP%' -UseBasicParsing; " ^
    "    Write-Host '[OK] Download concluido!' -ForegroundColor Green; " ^
    "    exit 0; " ^
    "} catch { " ^
    "    Write-Host '[ERRO] Falha no download: ' $_.Exception.Message -ForegroundColor Red; " ^
    "    exit 1; " ^
    "}"

if errorlevel 1 (
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

REM Download bem-sucedido! Agora cria backup da pasta antiga se existir
echo.
call :Log "================================================================================"
call :Log "  ETAPA 2: Backup da Instalacao Existente"
call :Log "================================================================================"
echo.

if exist "%INSTALL_DIR%" (
    set "BACKUP_DIR=%USERPROFILE%\chainsaw_backup_%DATETIME%"
    call :Log "[INFO] Pasta existente encontrada. Criando backup..."
    call :Log "[INFO] Destino do backup: !BACKUP_DIR!"
    
    REM Cria backup completo da pasta existente
    xcopy "%INSTALL_DIR%\*" "!BACKUP_DIR!\" /E /H /C /I /Y >nul 2>&1
    if errorlevel 1 (
        call :Log "[AVISO] Falha ao criar backup completo. Tentando backup seletivo..."
        
        REM Tenta backup apenas de arquivos críticos
        if not exist "!BACKUP_DIR!" mkdir "!BACKUP_DIR!"
        if exist "%INSTALL_DIR%\installation" xcopy "%INSTALL_DIR%\installation\*" "!BACKUP_DIR!\installation\" /E /H /C /I /Y >nul 2>&1
        if exist "%INSTALL_DIR%\source" xcopy "%INSTALL_DIR%\source\*" "!BACKUP_DIR!\source\" /E /H /C /I /Y >nul 2>&1
        
        call :Log "[OK] Backup seletivo criado."
    ) else (
        call :Log "[OK] Backup completo criado com sucesso!"
    )
    
    echo.
    call :Log "[INFO] Removendo pasta antiga..."
    
    REM Tentativa 1: Deletar pasta completa
    rd /s /q "%INSTALL_DIR%" >nul 2>&1
    if not exist "%INSTALL_DIR%" (
        call :Log "[OK] Pasta antiga removida com sucesso."
        goto :extract
    )
    
    REM Tentativa 2: Deletar conteúdo da pasta
    call :Log "[AVISO] Nao foi possivel remover a pasta. Tentando limpar conteudo..."
    del /f /s /q "%INSTALL_DIR%\*.*" >nul 2>&1
    for /d %%p in ("%INSTALL_DIR%\*") do rd /s /q "%%p" >nul 2>&1
    
    REM Tentativa 3: Força substituição durante extração
    if exist "%INSTALL_DIR%\*" (
        call :Log "[AVISO] Alguns arquivos nao puderam ser removidos."
        call :Log "[AVISO] A extracao ira substituir os arquivos existentes."
    ) else (
        call :Log "[OK] Conteudo da pasta removido."
    )
) else (
    call :Log "[INFO] Nenhuma instalacao anterior encontrada. Continuando..."
)

:extract

echo.
call :Log "================================================================================"
call :Log "  ETAPA 3: Extracao dos Arquivos"
call :Log "================================================================================"
echo.

REM Remove pasta temporária de extração se existir
if exist "%TEMP_EXTRACT%" rd /s /q "%TEMP_EXTRACT%" >nul 2>&1

REM Extrai o ZIP usando PowerShell
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ProgressPreference = 'SilentlyContinue'; " ^
    "try { " ^
    "    Write-Host '[INFO] Extraindo arquivos...' -ForegroundColor Cyan; " ^
    "    Expand-Archive -Path '%TEMP_ZIP%' -DestinationPath '%TEMP_EXTRACT%' -Force; " ^
    "    Write-Host '[OK] Extracao concluida!' -ForegroundColor Green; " ^
    "    exit 0; " ^
    "} catch { " ^
    "    Write-Host '[ERRO] Falha na extracao: ' $_.Exception.Message -ForegroundColor Red; " ^
    "    exit 1; " ^
    "}"

if errorlevel 1 (
    call :Log "[ERRO] Falha ao extrair o arquivo ZIP."
    pause
    exit /b 1
)

REM Move os arquivos extraídos para o destino final
call :Log "[INFO] Movendo arquivos para %INSTALL_DIR%..."

REM O GitHub cria uma pasta "chainsaw-main" dentro do ZIP
set "SOURCE_DIR=%TEMP_EXTRACT%\chainsaw-main"

if not exist "%SOURCE_DIR%" (
    call :Log "[ERRO] Pasta extraida nao encontrada em: %SOURCE_DIR%"
    pause
    exit /b 1
)

REM Cria pasta de destino se não existir
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

REM Copia todos os arquivos e pastas
xcopy "%SOURCE_DIR%\*" "%INSTALL_DIR%\" /E /H /C /I /Y >nul
if errorlevel 1 (
    call :Log "[ERRO] Falha ao copiar arquivos para o destino."
    pause
    exit /b 1
)

call :Log "[OK] Arquivos copiados com sucesso!"

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
