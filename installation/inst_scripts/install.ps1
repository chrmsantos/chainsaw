# =============================================================================
# CHAINSAW - Script de Instalação de Configurações do Word
# =============================================================================
# Versão: 2.0.0
# Licença: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
# Compatibilidade: Windows 10+, PowerShell 5.1+
# Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
# =============================================================================

<#
.SYNOPSIS
    Instala as configurações do Word do chainsaw para o usuário atual.

.DESCRIPTION
    Este script realiza as seguintes operações:
    1. Copia o arquivo stamp.png para a pasta do usuário
    2. Faz backup da pasta Templates atual
    3. Copia os novos Templates
    4. Detecta e importa personalizações do Word (se encontradas)
    5. Registra todas as operações em arquivo de log
    
    Se uma pasta 'exported-config' for encontrada no diretório do script,
    as personalizações do Word (Ribbon, Partes Rápidas, etc.) serão 
    automaticamente importadas.

.PARAMETER SourcePath
    Caminho base dos arquivos. Padrão: pasta onde o script está localizado

.PARAMETER Force
    Força a instalação sem confirmação do usuário.

.PARAMETER NoBackup
    Não cria backup da pasta Templates existente (não recomendado).

.PARAMETER SkipCustomizations
    Não importa personalizações do Word mesmo se encontradas.

.EXAMPLE
    .\install.ps1
    Executa a instalação com confirmação do usuário.

.EXAMPLE
    .\install.ps1 -Force
    Executa a instalação sem confirmação.

.EXAMPLE
    .\install.ps1 -SkipCustomizations
    Instala apenas Templates, sem importar personalizações.

.NOTES
    Requer permissões de escrita nas pastas do usuário.
    Não requer privilégios de administrador.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$SourcePath = "",
    
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$NoBackup,
    
    [Parameter()]
    [switch]$SkipCustomizations,
    
    [Parameter(DontShow)]
    [switch]$BypassedExecution
)

# Remove argumento de maximização se presente
if ($SourcePath -eq "__MAXIMIZED__") {
    $SourcePath = ""
}

# Maximiza a janela do PowerShell
if ($Host.Name -eq "ConsoleHost") {
    $psWindow = (Get-Host).UI.RawUI
    $newSize = $psWindow.BufferSize
    $newSize.Width = 120
    $newSize.Height = 9999
    try {
        $psWindow.BufferSize = $newSize
        $psWindow.WindowSize = $psWindow.MaxPhysicalWindowSize
    }
    catch {
        # Ignora erros se não for possível maximizar
    }
}

# Define o caminho padrão como a pasta onde o script está localizado
if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $SourcePath = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($SourcePath)) {
        # Fallback se PSScriptRoot não estiver disponível
        $SourcePath = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
}

# =============================================================================
# AUTO-RELANÇAMENTO COM BYPASS DE EXECUÇÃO
# =============================================================================
# Este bloco garante que o script seja executado com a política de execução
# adequada, sem modificar permanentemente as configurações do sistema.
# Extremamente seguro: apenas este script é executado com bypass temporário.
# =============================================================================

if (-not $BypassedExecution) {
    Write-Host " Verificando política de execução..." -ForegroundColor Cyan
    
    # Captura a política atual para documentação no log
    $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
    Write-Host "   Política atual (CurrentUser): $currentPolicy" -ForegroundColor Gray
    
    # Verifica se precisa de bypass
    $needsBypass = $false
    try {
        # Tenta uma operação de script simples
        $null = [ScriptBlock]::Create("1 + 1").Invoke()
    }
    catch [System.Management.Automation.PSSecurityException] {
        $needsBypass = $true
    }
    
    if ($needsBypass -or $currentPolicy -eq "Restricted" -or $currentPolicy -eq "AllSigned") {
        Write-Host "[AVISO]  Política de execução restritiva detectada." -ForegroundColor Yellow
        Write-Host " Relançando script com bypass temporário..." -ForegroundColor Cyan
        Write-Host ""
        Write-Host "ℹ  SEGURANÇA:" -ForegroundColor Green
        Write-Host "   • Apenas ESTE script será executado com bypass" -ForegroundColor Gray
        Write-Host "   • A política do sistema NÃO será alterada" -ForegroundColor Gray
        Write-Host "   • O bypass expira quando o script terminar" -ForegroundColor Gray
        Write-Host "   • Nenhum privilégio de administrador é usado" -ForegroundColor Gray
        Write-Host ""
        
        # Constrói argumentos para o relançamento
        $arguments = @(
            "-ExecutionPolicy", "Bypass",
            "-NoProfile",
            "-File", "`"$PSCommandPath`"",
            "-BypassedExecution"
        )
        
        # Adiciona parâmetros originais
        # SourcePath é sempre definido automaticamente, então não precisa passar
        if ($Force) {
            $arguments += "-Force"
        }
        if ($NoBackup) {
            $arguments += "-NoBackup"
        }
        
        # Relança o script com bypass temporário
        $processInfo = Start-Process -FilePath "powershell.exe" `
            -ArgumentList $arguments `
            -Wait `
            -NoNewWindow `
            -PassThru
        
        # Retorna o código de saída do processo relançado
        exit $processInfo.ExitCode
    }
    else {
        Write-Host "[OK] Política de execução adequada: $currentPolicy" -ForegroundColor Green
        Write-Host ""
    }
}
else {
    Write-Host "[OK] Executando com bypass temporário (seguro)" -ForegroundColor Green
    Write-Host ""
}

# =============================================================================
# CONFIGURAÇÕES E CONSTANTES
# =============================================================================

$ErrorActionPreference = "Stop"
$script:LogFile = $null
$script:WarningCount = 0
$script:ErrorCount = 0
$script:SuccessCount = 0

# Cores para output
$ColorSuccess = "Green"
$ColorWarning = "Yellow"
$ColorError = "Red"
$ColorInfo = "Cyan"

# =============================================================================
# FUNÇÕES DE LOG
# =============================================================================

function Initialize-LogFile {
    <#
    .SYNOPSIS
        Inicializa o arquivo de log.
    #>
    try {
        # Log na nova estrutura: installation/inst_docs/inst_logs
        $projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
        $logDir = Join-Path $projectRoot "installation\inst_docs\inst_logs"
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $script:LogFile = Join-Path $logDir "install_$timestamp.log"
        
        $header = @"
================================================================================
CHAINSAW - Log de Instalação
================================================================================
Data/Hora Início: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")
Usuário: $env:USERNAME
Computador: $env:COMPUTERNAME
Sistema: $([Environment]::OSVersion.VersionString)
PowerShell: $($PSVersionTable.PSVersion)
Caminho de Origem: $SourcePath
================================================================================

"@
        Add-Content -Path $script:LogFile -Value $header
        return $true
    }
    catch {
        Write-Warning "Não foi possível criar arquivo de log: $_"
        return $false
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Escreve mensagem no log e na tela.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR")]
        [string]$Level = "INFO",
        
        [Parameter()]
        [switch]$NoConsole
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Escreve no arquivo de log
    if ($script:LogFile) {
        try {
            Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
        }
        catch {
            # Ignora erros de escrita no log para não interromper o processo
        }
    }
    
    # Escreve no console
    if (-not $NoConsole) {
        switch ($Level) {
            "SUCCESS" {
                Write-Host "[OK] $Message" -ForegroundColor $ColorSuccess
                $script:SuccessCount++
            }
            "WARNING" {
                Write-Host "[AVISO] $Message" -ForegroundColor $ColorWarning
                $script:WarningCount++
            }
            "ERROR" {
                Write-Host "[ERRO] $Message" -ForegroundColor $ColorError
                $script:ErrorCount++
            }
            default {
                Write-Host "ℹ $Message" -ForegroundColor $ColorInfo
            }
        }
    }
}

# =============================================================================
# NORMALIZAÇÃO DO SOURCEPATH
# =============================================================================

# Normaliza/resolve o SourcePath quando fornecido como relativo/alias
if (-not [string]::IsNullOrWhiteSpace($SourcePath)) {
    try {
        # Tenta resolver caminhos absolutos ou relativos existentes
        $resolved = Resolve-Path -Path $SourcePath -ErrorAction Stop
        $SourcePath = $resolved.ProviderPath
    }
    catch {
        # Se não for um caminho absoluto, tente relative ao diretório do script
        try {
            if (-not [IO.Path]::IsPathRooted($SourcePath) -and -not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
                $candidate = Join-Path $PSScriptRoot $SourcePath
                if (Test-Path $candidate) {
                    $SourcePath = (Resolve-Path -Path $candidate).ProviderPath
                }
            }
        }
        catch {
            # ignora - deixamos o valor original para que as validações posteriores loguem
        }
    }
}

# =============================================================================
# FUNÇÕES DE VALIDAÇÃO
# =============================================================================

function Test-Prerequisites {
    <#
    .SYNOPSIS
        Verifica pré-requisitos para instalação.
    #>
    Write-Log "Verificando pré-requisitos..." -Level INFO
    
    $allOk = $true
    
    # Verifica versão do Windows
    $osVersion = [Environment]::OSVersion.Version
    if ($osVersion.Major -lt 10) {
        Write-Log "Windows 10 ou superior é necessário. Versão detectada: $($osVersion.ToString())" -Level ERROR
        $allOk = $false
    }
    else {
        Write-Log "Sistema operacional: Windows $($osVersion.Major).$($osVersion.Minor) [OK]" -Level SUCCESS
    }
    
    # Verifica versão do PowerShell
    $psVersion = $PSVersionTable.PSVersion
    if ($psVersion.Major -lt 5) {
        Write-Log "PowerShell 5.1 ou superior é necessário. Versão detectada: $($psVersion.ToString())" -Level ERROR
        $allOk = $false
    }
    else {
        Write-Log "PowerShell versão: $($psVersion.ToString()) [OK]" -Level SUCCESS
    }
    
    # Verifica acesso ao diretório de origem
    Write-Log "Verificando acesso ao diretório de origem: $SourcePath" -Level INFO
    if (-not (Test-Path $SourcePath)) {
        Write-Log "Não foi possível acessar o diretório de origem: $SourcePath" -Level ERROR
        Write-Log "Verifique se o diretório existe e você tem permissões de acesso." -Level ERROR
        $allOk = $false
    }
    else {
        Write-Log "Acesso ao diretório de origem confirmado [OK]" -Level SUCCESS
    }
    
    # Verifica permissões de escrita no perfil do usuário
    $testFile = Join-Path $env:USERPROFILE "CHAINSAW_test_$(Get-Date -Format 'yyyyMMddHHmmss').tmp"
    try {
        [System.IO.File]::WriteAllText($testFile, "test")
        Remove-Item $testFile -Force -ErrorAction SilentlyContinue
        Write-Log "Permissões de escrita no perfil do usuário confirmadas [OK]" -Level SUCCESS
    }
    catch {
        Write-Log "Sem permissões de escrita no perfil do usuário: $env:USERPROFILE" -Level ERROR
        $allOk = $false
    }
    
    return $allOk
}

function Test-SourceFiles {
    <#
    .SYNOPSIS
        Verifica se os arquivos de origem existem.
    #>
    param(
        [ref]$SourceStampFile,
        [ref]$SourceTemplatesFolder,
        [ref]$ProjectRoot
    )
    
    Write-Log "Verificando arquivos de origem..." -Level INFO
    
    $allOk = $true
    
    # Detecta a raiz do projeto procurando por assets\stamp.png
    $detectedProjectRoot = $SourcePath
    $currentPath = $SourcePath
    $maxLevels = 5
    $level = 0
    
    while ($level -lt $maxLevels) {
        $testStampPath = Join-Path $currentPath "assets\stamp.png"
        if (Test-Path $testStampPath) {
            $detectedProjectRoot = $currentPath
            Write-Log "Raiz do projeto detectada: $detectedProjectRoot" -Level INFO
            break
        }
        $parentPath = Split-Path $currentPath -Parent
        if ([string]::IsNullOrEmpty($parentPath) -or $parentPath -eq $currentPath) {
            break
        }
        $currentPath = $parentPath
        $level++
    }
    
    # Validação: se não encontrou a raiz do projeto nas iterações acima, usa o SourcePath como fallback
    if ($detectedProjectRoot -eq $SourcePath -and -not (Test-Path (Join-Path $SourcePath "assets\stamp.png"))) {
        Write-Log "Aviso: Não foi possível detectar a raiz do projeto automaticamente" -Level WARNING
        Write-Log "Usando caminho de origem como fallback: $SourcePath" -Level WARNING
    }
    
    # Retorna a raiz do projeto detectada (após validação)
    # Verifica se $ProjectRoot é uma referência válida antes de tentar definir Value
    if ($null -ne $ProjectRoot) {
        try {
            $ProjectRoot.Value = $detectedProjectRoot
            Write-Log "Raiz do projeto armazenada: $detectedProjectRoot" -Level INFO
        }
        catch {
            Write-Log "Erro ao armazenar raiz do projeto: $_" -Level ERROR
        }
    }
    else {
        Write-Log "AVISO: ProjectRoot é null, não foi possível armazenar" -Level WARNING
    }
    
    # Verifica arquivo stamp.png
    $stampPath = Join-Path $detectedProjectRoot "assets\stamp.png"
    if (Test-Path $stampPath) {
        if ($null -ne $SourceStampFile) {
            try {
                $SourceStampFile.Value = $stampPath
            }
            catch {
                Write-Log "Erro ao armazenar caminho stamp.png: $_" -Level ERROR
            }
        }
        Write-Log "Arquivo stamp.png encontrado [OK]" -Level SUCCESS
    }
    else {
        Write-Log "Arquivo não encontrado: $stampPath" -Level ERROR
        $allOk = $false
    }
    
    # Verifica pasta Templates usando a raiz do projeto detectada
    $templatesPath = Join-Path $detectedProjectRoot "installation\inst_configs\Templates"
    if (Test-Path $templatesPath) {
        if ($null -ne $SourceTemplatesFolder) {
            try {
                $SourceTemplatesFolder.Value = $templatesPath
            }
            catch {
                Write-Log "Erro ao armazenar caminho Templates: $_" -Level ERROR
            }
        }
        Write-Log "Pasta Templates encontrada [OK]" -Level SUCCESS
    }
    else {
        Write-Log "Pasta não encontrada: $templatesPath" -Level ERROR
        $allOk = $false
    }
    
    return $allOk
}

# =============================================================================
# FUNÇÕES AUXILIARES
# =============================================================================

function Test-WordRunning {
    <#
    .SYNOPSIS
        Verifica se o Microsoft Word está em execução.
    #>
    $wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
    return ($null -ne $wordProcesses -and $wordProcesses.Count -gt 0)
}

# =============================================================================
# FUNÇÕES DE BACKUP
# =============================================================================

function Backup-TemplatesFolder {
    <#
    .SYNOPSIS
        Cria backup da pasta Templates existente.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$SourceFolder
    )
    
    if (-not (Test-Path $SourceFolder)) {
        Write-Log "Pasta Templates não existe, backup não necessário." -Level INFO
        return $null
    }
    
    # Verifica se o Word está aberto
    if (Test-WordRunning) {
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "║                  [AVISO] MICROSOFT WORD ABERTO [AVISO]                    ║" -ForegroundColor Yellow
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "O Microsoft Word está em execução e deve ser fechado antes de" -ForegroundColor Yellow
        Write-Host "continuar com a instalação." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Por favor:" -ForegroundColor White
        Write-Host "  1. Salve todos os documentos abertos no Word" -ForegroundColor Gray
        Write-Host "  2. Feche completamente o Microsoft Word" -ForegroundColor Gray
        Write-Host "  3. Pressione qualquer tecla para continuar" -ForegroundColor Gray
        Write-Host ""
        
        Write-Log "Aguardando fechamento do Word..." -Level WARNING
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
        # Verifica novamente
        if (Test-WordRunning) {
            Write-Log "Word ainda está aberto - abortando instalação" -Level ERROR
            throw "Microsoft Word deve ser fechado antes da instalação."
        }
        
        Write-Host "[OK] Word fechado, continuando..." -ForegroundColor Green
        Write-Host ""
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupName = "Templates_backup_$timestamp"
    $backupPath = Join-Path (Split-Path $SourceFolder -Parent) $backupName
    
    Write-Log "Criando backup da pasta Templates..." -Level INFO
    Write-Log "Origem: $SourceFolder" -Level INFO
    Write-Log "Destino: $backupPath" -Level INFO
    
    try {
        # Tenta usar Rename-Item primeiro (mais rápido)
        Rename-Item -Path $SourceFolder -NewName $backupName -Force -ErrorAction Stop
        Write-Log "Backup criado com sucesso: $backupName [OK]" -Level SUCCESS
        return $backupPath
    }
    catch [System.IO.IOException] {
        Write-Log "Erro de acesso ao renomear (possível arquivo em uso)" -Level WARNING
        Write-Log "Tentando método alternativo (cópia)..." -Level INFO
        
        try {
            # Método alternativo: copiar e depois deletar
            Copy-Item -Path $SourceFolder -Destination $backupPath -Recurse -Force -ErrorAction Stop
            
            # Aguarda um pouco para liberar arquivos
            Start-Sleep -Seconds 1
            
            # Remove a pasta original
            Remove-Item -Path $SourceFolder -Recurse -Force -ErrorAction Stop
            
            Write-Log "Backup criado com sucesso (método cópia): $backupName [OK]" -Level SUCCESS
            return $backupPath
        }
        catch {
            Write-Log "Erro ao criar backup com método alternativo: $_" -Level ERROR
            throw "Não foi possível criar backup. Certifique-se de que o Word está fechado e que não há arquivos em uso na pasta Templates."
        }
    }
    catch {
        Write-Log "Erro ao criar backup: $_" -Level ERROR
        throw
    }
}

function Remove-OldBackups {
    <#
    .SYNOPSIS
        Remove backups antigos mantendo apenas os mais recentes.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$BackupFolder,
        
        [Parameter()]
        [int]$KeepCount = 5
    )
    
    $backupParent = Split-Path $BackupFolder -Parent
    $backups = Get-ChildItem -Path $backupParent -Directory -Filter "Templates_backup_*" |
    Sort-Object Name -Descending
    
    if ($backups.Count -gt $KeepCount) {
        $toRemove = $backups | Select-Object -Skip $KeepCount
        
        Write-Log "Removendo backups antigos (mantendo os $KeepCount mais recentes)..." -Level INFO
        
        foreach ($backup in $toRemove) {
            try {
                Remove-Item -Path $backup.FullName -Recurse -Force -ErrorAction Stop
                Write-Log "Backup removido: $($backup.Name)" -Level INFO
            }
            catch {
                Write-Log "Erro ao remover backup $($backup.Name): $_" -Level WARNING
            }
        }
    }
}

# =============================================================================
# FUNÇÕES DE INSTALAÇÃO
# =============================================================================

function Copy-StampFile {
    <#
    .SYNOPSIS
        Copia o arquivo stamp.png para a pasta do usuário.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$SourceFile
    )
    
    $destFolder = Join-Path $env:USERPROFILE "CHAINSAW\assets"
    $destFile = Join-Path $destFolder "stamp.png"
    
    Write-Log "Copiando arquivo stamp.png..." -Level INFO
    Write-Log "Origem: $SourceFile" -Level INFO
    Write-Log "Destino: $destFile" -Level INFO
    
    try {
        # Verifica se origem e destino são o mesmo arquivo
        $sourceFullPath = (Resolve-Path $SourceFile).Path
        $destFullPath = if (Test-Path $destFile) { (Resolve-Path $destFile).Path } else { $null }
        
        if ($sourceFullPath -eq $destFullPath) {
            Write-Log "Arquivo já está no local correto (origem = destino), pulando cópia" -Level INFO
            Write-Log "Arquivo stamp.png já está instalado [OK]" -Level SUCCESS
            return $true
        }
        
        # Cria pasta de destino se não existir
        if (-not (Test-Path $destFolder)) {
            New-Item -Path $destFolder -ItemType Directory -Force | Out-Null
            Write-Log "Pasta criada: $destFolder" -Level INFO
        }
        
        # Copia o arquivo
        Copy-Item -Path $SourceFile -Destination $destFile -Force -ErrorAction Stop
        
        # Verifica se o arquivo foi copiado corretamente
        if (Test-Path $destFile) {
            $sourceSize = (Get-Item $SourceFile).Length
            $destSize = (Get-Item $destFile).Length
            
            if ($sourceSize -eq $destSize) {
                Write-Log "Arquivo stamp.png copiado com sucesso [OK]" -Level SUCCESS
                return $true
            }
            else {
                Write-Log "Tamanhos diferentes: origem=$sourceSize, destino=$destSize" -Level WARNING
                return $false
            }
        }
        else {
            Write-Log "Arquivo não foi copiado corretamente" -Level ERROR
            return $false
        }
    }
    catch {
        Write-Log "Erro ao copiar stamp.png: $_" -Level ERROR
        throw
    }
}

function Copy-TemplatesFolder {
    <#
    .SYNOPSIS
        Copia a pasta Templates do diretório de origem para o perfil do usuário.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$SourceFolder,
        
        [Parameter(Mandatory)]
        [string]$DestFolder
    )
    
    Write-Log "Copiando pasta Templates..." -Level INFO
    Write-Log "Origem: $SourceFolder" -Level INFO
    Write-Log "Destino: $DestFolder" -Level INFO
    
    try {
        # Verifica se origem e destino são o mesmo local
        $sourceFullPath = (Resolve-Path $SourceFolder).Path.TrimEnd('\')
        $destFullPath = if (Test-Path $DestFolder) { (Resolve-Path $DestFolder).Path.TrimEnd('\') } else { $null }
        
        if ($sourceFullPath -eq $destFullPath) {
            Write-Log "A pasta Templates já está no local correto (origem = destino), pulando cópia" -Level INFO
            Write-Log "Pasta Templates já está instalada [OK]" -Level SUCCESS
            return $true
        }
        
        # Cria pasta de destino
        if (-not (Test-Path $DestFolder)) {
            New-Item -Path $DestFolder -ItemType Directory -Force | Out-Null
        }
        
        # Copia todos os arquivos e subpastas
        $itemsToCopy = Get-ChildItem -Path $SourceFolder -Recurse
        $totalItems = $itemsToCopy.Count
        $copiedItems = 0
        
        Write-Log "Total de itens a copiar: $totalItems" -Level INFO
        
        foreach ($item in $itemsToCopy) {
            $relativePath = $item.FullName.Substring($SourceFolder.Length + 1)
            $destPath = Join-Path $DestFolder $relativePath
            
            if ($item.PSIsContainer) {
                # É uma pasta
                if (-not (Test-Path $destPath)) {
                    New-Item -Path $destPath -ItemType Directory -Force | Out-Null
                }
            }
            else {
                # É um arquivo
                $destDir = Split-Path $destPath -Parent
                if (-not (Test-Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                }
                Copy-Item -Path $item.FullName -Destination $destPath -Force
                $copiedItems++
            }
            
            # Progress
            if ($copiedItems % 10 -eq 0) {
                Write-Progress -Activity "Copiando Templates" -Status "$copiedItems de $totalItems arquivos copiados" -PercentComplete (($copiedItems / $totalItems) * 100)
            }
        }
        
        Write-Progress -Activity "Copiando Templates" -Completed
        Write-Log "Pasta Templates copiada com sucesso ($copiedItems arquivos) [OK]" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Erro ao copiar pasta Templates: $_" -Level ERROR
        throw
    }
}

# =============================================================================
# FUNÇÕES DE IMPORTAÇÃO DE PERSONALIZAÇÕES
# =============================================================================

function Test-CustomizationsAvailable {
    param([string]$ImportPath)
    
    if (-not (Test-Path $ImportPath)) {
        return $false
    }
    
    # Verifica se há um manifesto ou arquivos para importar
    $manifestPath = Join-Path $ImportPath "MANIFEST.json"
    $hasManifest = Test-Path $manifestPath
    
    if ($hasManifest) {
        try {
            $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
            Write-Log "Manifesto encontrado: $($manifest.TotalItems) itens" -Level INFO
            Write-Log "Exportado em: $($manifest.ExportDate) por $($manifest.UserName)" -Level INFO
        }
        catch {
            Write-Log "Erro ao ler manifesto: $_" -Level WARNING
        }
    }
    
    return $true
}

function Backup-CompleteConfiguration {
    <#
    .SYNOPSIS
        Cria um backup completo de TODAS as configurações antes da instalação.
    .DESCRIPTION
        Este backup permite restaurar completamente o estado anterior à instalação.
        Inclui: Templates, Normal.dotm, personalizações UI, stamp.png, e metadados.
    #>
    
    if ($NoBackup) {
        Write-Log "Backup completo desabilitado (-NoBackup)" -Level WARNING
        return $null
    }
    
    Write-Log "Criando backup completo da configuração atual..." -Level INFO
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupBasePath = Join-Path $env:USERPROFILE "CHAINSAW\backups"
    $backupPath = Join-Path $backupBasePath "full_backup_$timestamp"
    
    try {
        # Cria estrutura de diretórios
        if (-not (Test-Path $backupPath)) {
            New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
        }
        
        $backupManifest = @{
            Timestamp = $timestamp
            Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            User = $env:USERNAME
            Computer = $env:COMPUTERNAME
            Items = @{}
        }
        
        # 1. Backup completo da pasta Templates
        $templatesPath = Join-Path $env:APPDATA "Microsoft\Templates"
        if (Test-Path $templatesPath) {
            $templatesBackupPath = Join-Path $backupPath "Templates"
            Write-Log "Fazendo backup da pasta Templates..." -Level INFO
            Copy-Item -Path $templatesPath -Destination $templatesBackupPath -Recurse -Force -ErrorAction Stop
            
            $templatesSize = (Get-ChildItem -Path $templatesBackupPath -Recurse -File | Measure-Object -Property Length -Sum).Sum / 1MB
            $backupManifest.Items.Templates = @{
                Path = $templatesBackupPath
                SizeMB = [math]::Round($templatesSize, 2)
                Files = (Get-ChildItem -Path $templatesBackupPath -Recurse -File).Count
            }
            Write-Log "Templates backup: $([math]::Round($templatesSize, 2)) MB, $($backupManifest.Items.Templates.Files) arquivos" -Level INFO
        }
        else {
            Write-Log "Pasta Templates não existe - pulando" -Level INFO
        }
        
        # 2. Backup do stamp.png (se existir)
        $stampPath = Join-Path $env:USERPROFILE "CHAINSAW\assets\stamp.png"
        if (Test-Path $stampPath) {
            $stampBackupPath = Join-Path $backupPath "stamp.png"
            Copy-Item -Path $stampPath -Destination $stampBackupPath -Force -ErrorAction Stop
            $backupManifest.Items.Stamp = @{
                Path = $stampBackupPath
                SizeKB = [math]::Round((Get-Item $stampBackupPath).Length / 1KB, 2)
            }
            Write-Log "stamp.png backup criado: $($backupManifest.Items.Stamp.SizeKB) KB" -Level INFO
        }
        
        # 3. Backup de personalizações (Normal.dotm, UI customizations)
        $customBackupPath = Join-Path $backupPath "Customizations"
        New-Item -Path $customBackupPath -ItemType Directory -Force | Out-Null
        
        # Normal.dotm (cópia adicional)
        $normalPath = Join-Path $templatesPath "Normal.dotm"
        if (Test-Path $normalPath) {
            $normalBackupPath = Join-Path $customBackupPath "Templates"
            New-Item -Path $normalBackupPath -ItemType Directory -Force | Out-Null
            Copy-Item -Path $normalPath -Destination $normalBackupPath -Force
            Write-Log "Normal.dotm backup adicional criado" -Level INFO
        }
        
        # Personalizações UI do Office
        $localAppDataPath = $env:LOCALAPPDATA
        $uiPath = Join-Path $localAppDataPath "Microsoft\Office"
        $uiFiles = Get-ChildItem -Path $uiPath -Filter "*.officeUI" -Recurse -ErrorAction SilentlyContinue
        
        if ($uiFiles.Count -gt 0) {
            $destUI = Join-Path $customBackupPath "OfficeCustomUI"
            New-Item -Path $destUI -ItemType Directory -Force | Out-Null
            foreach ($file in $uiFiles) {
                Copy-Item -Path $file.FullName -Destination (Join-Path $destUI $file.Name) -Force
            }
            $backupManifest.Items.Customizations = @{
                UIFiles = $uiFiles.Count
            }
            Write-Log "Personalizações UI backup: $($uiFiles.Count) arquivos" -Level INFO
        }
        
        # 4. Salva manifesto do backup
        $manifestPath = Join-Path $backupPath "backup_manifest.json"
        $backupManifest | ConvertTo-Json -Depth 10 | Out-File -FilePath $manifestPath -Encoding UTF8
        
        # 5. Cria arquivo README no backup
        $readmePath = Join-Path $backupPath "README.txt"
        $readmeContent = @"
================================================================================
CHAINSAW - BACKUP COMPLETO
================================================================================

Data do Backup: $($backupManifest.Date)
Usuário: $($backupManifest.User)
Computador: $($backupManifest.Computer)

Este backup contém uma cópia completa das configurações do Word antes da
instalação do CHAINSAW.

================================================================================
CONTEÚDO DO BACKUP
================================================================================

"@
        
        if ($backupManifest.Items.Templates) {
            $readmeContent += @"

[Templates]
  - Pasta completa: Templates\
  - Tamanho: $($backupManifest.Items.Templates.SizeMB) MB
  - Arquivos: $($backupManifest.Items.Templates.Files)
  - Inclui: Normal.dotm, Building Blocks, Temas, Estilos, etc.

"@
        }
        
        if ($backupManifest.Items.Stamp) {
            $readmeContent += @"

[Stamp]
  - Arquivo: stamp.png
  - Tamanho: $($backupManifest.Items.Stamp.SizeKB) KB

"@
        }
        
        if ($backupManifest.Items.Customizations) {
            $readmeContent += @"

[Personalizações]
  - Pasta: Customizations\
  - Arquivos UI: $($backupManifest.Items.Customizations.UIFiles)
  - Inclui: Ribbon customizations, Quick Access Toolbar, etc.

"@
        }
        
        $readmeContent += @"

================================================================================
RESTAURAÇÃO
================================================================================

Para restaurar este backup, execute:

    .\restore-backup.cmd

Ou manualmente:

    powershell.exe -ExecutionPolicy Bypass -File restore-backup.ps1 -BackupPath "$backupPath"

================================================================================
IMPORTANTE
================================================================================

- Este backup é criado AUTOMATICAMENTE antes de cada instalação
- Backups antigos são mantidos por segurança
- Para liberar espaço, remova manualmente backups antigos desta pasta:
  $backupBasePath

================================================================================
"@
        
        $readmeContent | Out-File -FilePath $readmePath -Encoding UTF8
        
        Write-Log "Backup completo criado com sucesso em: $backupPath [OK]" -Level SUCCESS
        Write-Log "Manifesto salvo em: $manifestPath" -Level INFO
        
        return $backupPath
    }
    catch {
        Write-Log "Erro ao criar backup completo: $_" -Level ERROR
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
        return $null
    }
}

function Backup-WordCustomizations {
    param([string]$BackupReason = "pré-importação")
    
    if ($NoBackup) {
        Write-Log "Backup de personalizações desabilitado (-NoBackup)" -Level WARNING
        return $null
    }
    
    Write-Log "Criando backup das personalizações do Word ($BackupReason)..." -Level INFO
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = Join-Path $env:USERPROFILE "CHAINSAW\backups\word-customizations_$timestamp"
    
    try {
        if (-not (Test-Path $backupPath)) {
            New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
        }
        
        $templatesPath = Join-Path $env:APPDATA "Microsoft\Templates"
        $localAppDataPath = $env:LOCALAPPDATA
        
        # Backup do Normal.dotm
        $normalPath = Join-Path $templatesPath "Normal.dotm"
        if (Test-Path $normalPath) {
            $destNormal = Join-Path $backupPath "Templates"
            New-Item -Path $destNormal -ItemType Directory -Force | Out-Null
            Copy-Item -Path $normalPath -Destination $destNormal -Force
            Write-Log "Normal.dotm backup criado" -Level INFO
        }
        
        # Backup de personalizações UI
        $uiPath = Join-Path $localAppDataPath "Microsoft\Office"
        $uiFiles = Get-ChildItem -Path $uiPath -Filter "*.officeUI" -Recurse -ErrorAction SilentlyContinue
        if ($uiFiles.Count -gt 0) {
            $destUI = Join-Path $backupPath "OfficeCustomUI"
            New-Item -Path $destUI -ItemType Directory -Force | Out-Null
            foreach ($file in $uiFiles) {
                Copy-Item -Path $file.FullName -Destination (Join-Path $destUI $file.Name) -Force
            }
            Write-Log "Personalizações UI backup criado: $($uiFiles.Count) arquivos" -Level INFO
        }
        
        Write-Log "Backup de personalizações criado em: $backupPath [OK]" -Level SUCCESS
        return $backupPath
    }
    catch {
        Write-Log "Erro ao criar backup de personalizações: $_" -Level ERROR
        return $null
    }
}

function Import-NormalTemplate {
    param([string]$ImportPath)
    
    Write-Log "Importando Normal.dotm..." -Level INFO
    
    $sourcePath = Join-Path $ImportPath "Templates\Normal.dotm"
    $templatesPath = Join-Path $env:APPDATA "Microsoft\Templates"
    $destPath = Join-Path $templatesPath "Normal.dotm"
    
    if (-not (Test-Path $sourcePath)) {
        Write-Log "Normal.dotm não encontrado no pacote de importação" -Level WARNING
        return $false
    }
    
    try {
        if (-not (Test-Path $templatesPath)) {
            New-Item -Path $templatesPath -ItemType Directory -Force | Out-Null
        }
        
        Copy-Item -Path $sourcePath -Destination $destPath -Force
        Write-Log "Normal.dotm importado com sucesso [OK]" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Erro ao importar Normal.dotm: $_" -Level ERROR
        return $false
    }
}

function Import-BuildingBlocks {
    param([string]$ImportPath)
    
    Write-Log "Importando Building Blocks..." -Level INFO
    
    $templatesPath = Join-Path $env:APPDATA "Microsoft\Templates"
    $sourceManaged = Join-Path $ImportPath "Templates\LiveContent\16\Managed\Word Document Building Blocks"
    $sourceUser = Join-Path $ImportPath "Templates\LiveContent\16\User\Word Document Building Blocks"
    
    $destManaged = Join-Path $templatesPath "LiveContent\16\Managed\Word Document Building Blocks"
    $destUser = Join-Path $templatesPath "LiveContent\16\User\Word Document Building Blocks"
    
    $importedCount = 0
    
    # Importa Building Blocks gerenciados
    if (Test-Path $sourceManaged) {
        try {
            if (-not (Test-Path $destManaged)) {
                New-Item -Path $destManaged -ItemType Directory -Force | Out-Null
            }
            
            $files = Get-ChildItem -Path $sourceManaged -Recurse -File
            foreach ($file in $files) {
                $relativePath = $file.FullName.Substring($sourceManaged.Length + 1)
                $destFile = Join-Path $destManaged $relativePath
                $destDir = Split-Path $destFile -Parent
                
                if (-not (Test-Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                }
                
                Copy-Item -Path $file.FullName -Destination $destFile -Force
                $importedCount++
            }
            
            Write-Log "Building Blocks gerenciados importados: $($files.Count) arquivos" -Level INFO
        }
        catch {
            Write-Log "Erro ao importar Building Blocks gerenciados: $_" -Level WARNING
        }
    }
    
    # Importa Building Blocks do usuário
    if (Test-Path $sourceUser) {
        try {
            if (-not (Test-Path $destUser)) {
                New-Item -Path $destUser -ItemType Directory -Force | Out-Null
            }
            
            $files = Get-ChildItem -Path $sourceUser -Recurse -File
            foreach ($file in $files) {
                $relativePath = $file.FullName.Substring($sourceUser.Length + 1)
                $destFile = Join-Path $destUser $relativePath
                $destDir = Split-Path $destFile -Parent
                
                if (-not (Test-Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                }
                
                Copy-Item -Path $file.FullName -Destination $destFile -Force
                $importedCount++
            }
            
            Write-Log "Building Blocks do usuário importados: $($files.Count) arquivos" -Level INFO
        }
        catch {
            Write-Log "Erro ao importar Building Blocks do usuário: $_" -Level WARNING
        }
    }
    
    if ($importedCount -gt 0) {
        Write-Log "Building Blocks importados: $importedCount arquivos [OK]" -Level SUCCESS
        return $true
    }
    else {
        Write-Log "Nenhum Building Block para importar" -Level INFO
        return $false
    }
}

function Import-DocumentThemes {
    param([string]$ImportPath)
    
    Write-Log "Importando temas de documentos..." -Level INFO
    
    $templatesPath = Join-Path $env:APPDATA "Microsoft\Templates"
    $sourceManaged = Join-Path $ImportPath "Templates\LiveContent\16\Managed\Document Themes"
    $sourceUser = Join-Path $ImportPath "Templates\LiveContent\16\User\Document Themes"
    
    $destManaged = Join-Path $templatesPath "LiveContent\16\Managed\Document Themes"
    $destUser = Join-Path $templatesPath "LiveContent\16\User\Document Themes"
    
    $importedCount = 0
    
    # Temas gerenciados
    if (Test-Path $sourceManaged) {
        try {
            if (-not (Test-Path $destManaged)) {
                New-Item -Path $destManaged -ItemType Directory -Force | Out-Null
            }
            
            $files = Get-ChildItem -Path $sourceManaged -Recurse -File
            foreach ($file in $files) {
                $relativePath = $file.FullName.Substring($sourceManaged.Length + 1)
                $destFile = Join-Path $destManaged $relativePath
                $destDir = Split-Path $destFile -Parent
                
                if (-not (Test-Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                }
                
                Copy-Item -Path $file.FullName -Destination $destFile -Force
                $importedCount++
            }
        }
        catch {
            Write-Log "Erro ao importar temas gerenciados: $_" -Level WARNING
        }
    }
    
    # Temas do usuário
    if (Test-Path $sourceUser) {
        try {
            if (-not (Test-Path $destUser)) {
                New-Item -Path $destUser -ItemType Directory -Force | Out-Null
            }
            
            $files = Get-ChildItem -Path $sourceUser -Recurse -File
            foreach ($file in $files) {
                $relativePath = $file.FullName.Substring($sourceUser.Length + 1)
                $destFile = Join-Path $destUser $relativePath
                $destDir = Split-Path $destFile -Parent
                
                if (-not (Test-Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                }
                
                Copy-Item -Path $file.FullName -Destination $destFile -Force
                $importedCount++
            }
        }
        catch {
            Write-Log "Erro ao importar temas do usuário: $_" -Level WARNING
        }
    }
    
    if ($importedCount -gt 0) {
        Write-Log "Temas importados: $importedCount arquivos [OK]" -Level SUCCESS
        return $true
    }
    else {
        Write-Log "Nenhum tema para importar" -Level INFO
        return $false
    }
}

function Import-RibbonCustomization {
    param([string]$ImportPath)
    
    Write-Log "Importando personalização da Faixa de Opções..." -Level INFO
    
    $sourcePath = Join-Path $ImportPath "RibbonCustomization"
    
    if (-not (Test-Path $sourcePath)) {
        Write-Log "Nenhuma personalização do Ribbon para importar" -Level INFO
        return $false
    }
    
    try {
        $files = Get-ChildItem -Path $sourcePath -Filter "*.officeUI" -ErrorAction SilentlyContinue
        
        if ($files.Count -eq 0) {
            Write-Log "Nenhum arquivo de personalização Ribbon encontrado" -Level INFO
            return $false
        }
        
        foreach ($file in $files) {
            # Tenta os locais possíveis
            $possibleDests = @(
                (Join-Path $env:LOCALAPPDATA "Microsoft\Office"),
                (Join-Path $env:APPDATA "Microsoft\Office")
            )
            
            foreach ($destPath in $possibleDests) {
                if (-not (Test-Path $destPath)) {
                    New-Item -Path $destPath -ItemType Directory -Force | Out-Null
                }
                
                $destFile = Join-Path $destPath $file.Name
                Copy-Item -Path $file.FullName -Destination $destFile -Force
                Write-Log "Ribbon importado para: $destFile" -Level INFO
            }
        }
        
        Write-Log "Personalização do Ribbon importada: $($files.Count) arquivos [OK]" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Erro ao importar Ribbon: $_" -Level ERROR
        return $false
    }
}

function Import-OfficeCustomUI {
    param([string]$ImportPath)
    
    Write-Log "Importando personalizações da interface..." -Level INFO
    
    $sourcePath = Join-Path $ImportPath "OfficeCustomUI"
    
    if (-not (Test-Path $sourcePath)) {
        Write-Log "Nenhuma personalização UI para importar" -Level INFO
        return $false
    }
    
    try {
        $files = Get-ChildItem -Path $sourcePath -Filter "*.officeUI" -ErrorAction SilentlyContinue
        
        if ($files.Count -eq 0) {
            Write-Log "Nenhum arquivo de personalização UI encontrado" -Level INFO
            return $false
        }
        
        $destPath = Join-Path $env:LOCALAPPDATA "Microsoft\Office"
        if (-not (Test-Path $destPath)) {
            New-Item -Path $destPath -ItemType Directory -Force | Out-Null
        }
        
        foreach ($file in $files) {
            $destFile = Join-Path $destPath $file.Name
            Copy-Item -Path $file.FullName -Destination $destFile -Force
        }
        
        Write-Log "Personalizações UI importadas: $($files.Count) arquivos [OK]" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Erro ao importar personalizações UI: $_" -Level ERROR
        return $false
    }
}

function Import-WordCustomizations {
    param([string]$ImportPath)
    
    Write-Log "=== Iniciando importação de personalizações ===" -Level INFO
    
    # Verifica se o Word está em execução
    if (Test-WordRunning) {
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "║                  [AVISO] MICROSOFT WORD ABERTO [AVISO]                    ║" -ForegroundColor Yellow
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "O Microsoft Word está em execução e deve ser fechado antes de" -ForegroundColor Yellow
        Write-Host "importar as personalizações." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Por favor:" -ForegroundColor White
        Write-Host "  1. Salve todos os documentos abertos no Word" -ForegroundColor Gray
        Write-Host "  2. Feche completamente o Microsoft Word" -ForegroundColor Gray
        Write-Host "  3. Execute este script novamente" -ForegroundColor Gray
        Write-Host ""
        
        Write-Log "Importação abortada: Word está em execução" -Level WARNING
        return $false
    }
    
    # Cria backup
    $backupPath = Backup-WordCustomizations -BackupReason "pré-importação de personalizações"
    if ($null -eq $backupPath -and -not $NoBackup) {
        Write-Host ""
        Write-Host "[AVISO] Falha ao criar backup das personalizações atuais." -ForegroundColor Yellow
        
        if (-not $Force) {
            $response = Read-Host "Continuar mesmo assim? (S/N)"
            if ($response -notmatch '^[Ss]$') {
                Write-Log "Importação cancelada: falha no backup" -Level WARNING
                return $false
            }
        }
    }
    
    Write-Host ""
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
    Write-Host "  ETAPA 6: Importação de Personalizações do Word" -ForegroundColor White
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
    Write-Host ""
    
    # Importações
    $importedCount = 0
    
    if (Import-NormalTemplate -ImportPath $ImportPath) { $importedCount++ }
    if (Import-BuildingBlocks -ImportPath $ImportPath) { $importedCount++ }
    if (Import-DocumentThemes -ImportPath $ImportPath) { $importedCount++ }
    if (Import-RibbonCustomization -ImportPath $ImportPath) { $importedCount++ }
    if (Import-OfficeCustomUI -ImportPath $ImportPath) { $importedCount++ }
    
    if ($importedCount -gt 0) {
        Write-Log "Total de categorias de personalizações importadas: $importedCount [OK]" -Level SUCCESS
        return $true
    }
    else {
        Write-Log "Nenhuma personalização foi importada" -Level WARNING
        return $false
    }
}

# =============================================================================
# GERENCIAMENTO DO WORD
# =============================================================================

function Test-WordRunning {
    <#
    .SYNOPSIS
        Verifica se há processos do Word em execução.
    #>
    $wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
    return ($null -ne $wordProcesses -and $wordProcesses.Count -gt 0)
}

function Stop-WordProcesses {
    <#
    .SYNOPSIS
        Fecha forçadamente todos os processos do Word.
    .DESCRIPTION
        Encerra apenas processos WINWORD.EXE (Microsoft Word), sem afetar
        outros aplicativos do Office como Excel (EXCEL.EXE) ou PowerPoint (POWERPNT.EXE).
    #>
    param(
        [Parameter()]
        [switch]$Force
    )
    
    try {
        $wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
        
        if ($null -eq $wordProcesses -or $wordProcesses.Count -eq 0) {
            Write-Log "Nenhum processo do Word em execução" -Level INFO
            return $true
        }
        
        Write-Log "Encontrados $($wordProcesses.Count) processo(s) do Word em execução" -Level INFO
        
        foreach ($process in $wordProcesses) {
            try {
                Write-Log "Encerrando processo Word (PID: $($process.Id))..." -Level INFO
                
                if ($Force) {
                    # Encerra forçadamente
                    $process.Kill()
                    $process.WaitForExit(5000) # Aguarda até 5 segundos
                }
                else {
                    # Tenta encerrar graciosamente primeiro
                    $process.CloseMainWindow() | Out-Null
                    Start-Sleep -Milliseconds 500
                    
                    if (-not $process.HasExited) {
                        $process.Kill()
                        $process.WaitForExit(5000)
                    }
                }
                
                Write-Log "Processo Word (PID: $($process.Id)) encerrado com sucesso" -Level SUCCESS
            }
            catch {
                Write-Log "Erro ao encerrar processo Word (PID: $($process.Id)): $_" -Level WARNING
            }
        }
        
        # Aguarda um momento e verifica se todos foram fechados
        Start-Sleep -Milliseconds 1000
        $remainingProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
        
        if ($null -ne $remainingProcesses -and $remainingProcesses.Count -gt 0) {
            Write-Log "Ainda há $($remainingProcesses.Count) processo(s) do Word em execução" -Level WARNING
            return $false
        }
        
        Write-Log "Todos os processos do Word foram encerrados com sucesso" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Erro ao encerrar processos do Word: $_" -Level ERROR
        return $false
    }
}

function Confirm-CloseWord {
    <#
    .SYNOPSIS
        Solicita que o usuário salve e feche o Word, ou cancela a operação.
    .DESCRIPTION
        Exibe aviso ao usuário e aguarda confirmação antes de fechar o Word forçadamente.
        Retorna $true se o usuário autorizar, $false se cancelar.
    #>
    
    # Verifica se Word está em execução
    if (-not (Test-WordRunning)) {
        Write-Log "Word não está em execução - prosseguindo..." -Level SUCCESS
        return $true
    }
    
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "║                          [AVISO] ATENÇÃO [AVISO]                          ║" -ForegroundColor Yellow
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "O Microsoft Word está atualmente em execução!" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "IMPORTANTE:" -ForegroundColor Red
    Write-Host "  • SALVE todos os seus documentos abertos no Word" -ForegroundColor White
    Write-Host "  • FECHE o Word completamente" -ForegroundColor White
    Write-Host "  • Outros aplicativos do Office (Excel, PowerPoint) NÃO serão afetados" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Se você continuar, o Word será FECHADO FORÇADAMENTE e" -ForegroundColor Red
    Write-Host "qualquer trabalho não salvo SERÁ PERDIDO!" -ForegroundColor Red
    Write-Host ""
    
    Write-Log "Word em execução - solicitando confirmação do usuário" -Level WARNING
    
    # Aguarda confirmação
    $response = Read-Host "Deseja FECHAR o Word e continuar a instalação? (S/N)"
    
    if ($response -notmatch '^[Ss]$') {
        Write-Host ""
        Write-Host "[OK] Instalação cancelada pelo usuário" -ForegroundColor Cyan
        Write-Host "  Salve seus documentos e execute o script novamente quando estiver pronto." -ForegroundColor Gray
        Write-Host ""
        Write-Log "Instalação cancelada - usuário optou por não fechar o Word" -Level WARNING
        return $false
    }
    
    # Usuário confirmou - fecha o Word
    Write-Host ""
    Write-Host "Fechando Microsoft Word..." -ForegroundColor Cyan
    Write-Log "Usuário autorizou o fechamento do Word" -Level INFO
    
    if (Stop-WordProcesses -Force) {
        Write-Host "[OK] Word fechado com sucesso" -ForegroundColor Green
        Write-Host ""
        # Aguarda um pouco para garantir que recursos foram liberados
        Start-Sleep -Seconds 2
        return $true
    }
    else {
        Write-Host "[ERRO] Não foi possível fechar o Word completamente" -ForegroundColor Red
        Write-Host ""
        Write-Log "Falha ao fechar Word - cancelando instalação" -Level ERROR
        
        $retry = Read-Host "Deseja tentar novamente? (S/N)"
        if ($retry -match '^[Ss]$') {
            return Confirm-CloseWord # Recursão para tentar novamente
        }
        return $false
    }
}

# =============================================================================
# FUNÇÃO PRINCIPAL
# =============================================================================

function Install-CHAINSAWConfig {
    <#
    .SYNOPSIS
        Função principal de instalação.
    #>
    
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║          CHAINSAW - Instalação de Configurações do Word       ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    
    # Inicializa log
    if (-not (Initialize-LogFile)) {
        Write-Warning "Continuando sem arquivo de log..."
    }
    else {
        Write-Host " Arquivo de log: $script:LogFile" -ForegroundColor Gray
        Write-Host ""
    }
    
    # Loga o SourcePath resolvido
    Write-Log "Caminho de Origem (resolvido): $SourcePath" -Level INFO
    
    $startTime = Get-Date
    Write-Log "=== INÍCIO DA INSTALAÇÃO ===" -Level INFO
    
    try {
        # 0. Verificar e fechar Word se necessário
        Write-Host ""
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host "  ETAPA 0: Verificação do Microsoft Word" -ForegroundColor White
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host ""
        
        if (-not (Confirm-CloseWord)) {
            Write-Log "Instalação cancelada - Word não foi fechado" -Level WARNING
            return
        }
        
        # 1. Verificar pré-requisitos
        Write-Host ""
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host "  ETAPA 1: Verificação de Pré-requisitos" -ForegroundColor White
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host ""
        
        if (-not (Test-Prerequisites)) {
            throw "Pré-requisitos não atendidos. Verifique os erros acima."
        }
        
        # 2. Verificar arquivos de origem
        Write-Host ""
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host "  ETAPA 2: Verificação de Arquivos de Origem" -ForegroundColor White
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host ""
        
        $sourceStampFile = $null
        $sourceTemplatesFolder = $null
        $projectRoot = $null
        
        if (-not (Test-SourceFiles -SourceStampFile ([ref]$sourceStampFile) -SourceTemplatesFolder ([ref]$sourceTemplatesFolder) -ProjectRoot ([ref]$projectRoot))) {
            throw "Arquivos de origem não encontrados. Verifique os erros acima."
        }
        
        # 3. Backup completo da configuração atual (ANTES de qualquer modificação)
        $fullBackupPath = $null
        if (-not $NoBackup) {
            Write-Host ""
            Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
            Write-Host "  ETAPA 3: Backup Completo da Configuração Atual" -ForegroundColor White
            Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
            Write-Host ""
            Write-Host "ℹ Criando backup completo antes da instalação..." -ForegroundColor Cyan
            Write-Host "  Este backup permitirá restaurar completamente o estado atual" -ForegroundColor Gray
            Write-Host ""
            
            $fullBackupPath = Backup-CompleteConfiguration
            
            if ($fullBackupPath) {
                Write-Host ""
                Write-Host "✓ Backup completo criado com sucesso!" -ForegroundColor Green
                Write-Host "  Localização: $fullBackupPath" -ForegroundColor Gray
                Write-Host ""
                Write-Host "  Para restaurar este backup, execute:" -ForegroundColor Yellow
                Write-Host "    .\restore-backup.cmd" -ForegroundColor White
                Write-Host ""
            }
            else {
                Write-Host ""
                Write-Host "⚠ Não foi possível criar backup completo" -ForegroundColor Yellow
                Write-Host "  A instalação pode continuar, mas não será possível restaurar" -ForegroundColor Yellow
                Write-Host ""
                
                if (-not $Force) {
                    $response = Read-Host "Deseja continuar mesmo assim? (S/N)"
                    if ($response -notmatch '^[Ss]$') {
                        Write-Log "Instalação cancelada pelo usuário (falha no backup)" -Level WARNING
                        return
                    }
                }
            }
        }
        
        # 4. Confirmação do usuário
        if (-not $Force) {
            Write-Host ""
            Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
            Write-Host "  CONFIRMAÇÃO" -ForegroundColor White
            Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
            Write-Host ""
            Write-Host "As seguintes operações serão realizadas:" -ForegroundColor Yellow
            Write-Host "  1. Copiar stamp.png para: $env:USERPROFILE\chainsaw\assets\" -ForegroundColor White
            Write-Host "  2. Fazer backup da pasta Templates atual (se existir)" -ForegroundColor White
            Write-Host "  3. Copiar nova pasta Templates do diretório local" -ForegroundColor White
            Write-Host ""
            
            if ($fullBackupPath) {
                Write-Host "✓ Backup completo já foi criado em:" -ForegroundColor Green
                Write-Host "  $fullBackupPath" -ForegroundColor Gray
                Write-Host ""
            }
            
            $response = Read-Host "Deseja continuar? (S/N)"
            if ($response -notmatch '^[Ss]$') {
                Write-Log "Instalação cancelada pelo usuário." -Level WARNING
                return
            }
        }
        
        # 5. Copiar arquivo stamp.png
        Write-Host ""
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host "  ETAPA 4: Cópia do Arquivo stamp.png" -ForegroundColor White
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host ""
        
        Copy-StampFile -SourceFile $sourceStampFile | Out-Null
        
        # 6. Backup da pasta Templates
        $templatesPath = Join-Path $env:APPDATA "Microsoft\Templates"
        $backupPath = $null
        
        if (-not $NoBackup) {
            Write-Host ""
            Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
            Write-Host "  ETAPA 5: Backup da Pasta Templates" -ForegroundColor White
            Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
            Write-Host ""
            
            $backupPath = Backup-TemplatesFolder -SourceFolder $templatesPath
            
            if ($backupPath) {
                Remove-OldBackups -BackupFolder $backupPath -KeepCount 5
            }
        }
        else {
            Write-Log "Backup desabilitado pelo parâmetro -NoBackup" -Level WARNING
        }
        
        # 7. Copiar pasta Templates
        Write-Host ""
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host "  ETAPA 6: Cópia da Pasta Templates" -ForegroundColor White
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host ""
        
        Copy-TemplatesFolder -SourceFolder $sourceTemplatesFolder -DestFolder $templatesPath | Out-Null
        
        # 8. Atualizar módulo VBA no Normal.dotm
        Write-Host ""
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host "  ETAPA 7: Atualização do Módulo VBA" -ForegroundColor White
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host ""
        
        # Usa a raiz do projeto detectada para encontrar o módulo VBA
        $vbaModulePath = Join-Path $projectRoot "source\main\monolithicMod.bas"
        if (Test-Path $vbaModulePath) {
            Write-Log "Módulo VBA encontrado: $vbaModulePath" -Level INFO
            Write-Host " Importando módulo VBA mais recente..." -ForegroundColor Cyan
            
            try {
                $normalDotmPath = Join-Path $templatesPath "Normal.dotm"
                
                if (-not (Test-Path $normalDotmPath)) {
                    Write-Log "Normal.dotm não encontrado em: $normalDotmPath" -Level ERROR
                    Write-Host "[ERRO] Normal.dotm não encontrado!" -ForegroundColor Red
                    Write-Host "  O módulo VBA precisa ser importado manualmente." -ForegroundColor Yellow
                }
                else {
                    # Cria objeto Word
                    $word = New-Object -ComObject Word.Application
                    $word.Visible = $false
                    $word.DisplayAlerts = 0  # wdAlertsNone
                    
                    # Abre Normal.dotm
                    $doc = $word.Documents.Open($normalDotmPath, $false, $false)
                    $vbProject = $doc.VBProject
                    
                    # Remove módulos antigos
                    $oldModuleNames = @("Módulo1", "Module1", "monolithicMod", "Mod_Main", "Chainsaw", "CHAINSAW_MODX", "Chainsaw_ModX", "chainsawModX")
                    
                    foreach ($moduleName in $oldModuleNames) {
                        try {
                            $module = $vbProject.VBComponents.Item($moduleName)
                            if ($module) {
                                # Faz backup do módulo antigo (nova estrutura)
                                $backupDir = Join-Path $projectRoot "source\backups"
                                if (-not (Test-Path $backupDir)) {
                                    New-Item -Path $backupDir -ItemType Directory -Force | Out-Null
                                }
                                $backupPath = Join-Path $backupDir "backup_$moduleName`_$(Get-Date -Format 'yyyyMMdd_HHmmss').bas"
                                $module.Export($backupPath)
                                Write-Log "Backup do módulo '$moduleName' criado: $backupPath" -Level INFO
                                
                                # Remove o módulo
                                $vbProject.VBComponents.Remove($module)
                                Write-Log "Módulo '$moduleName' removido" -Level INFO
                            }
                        }
                        catch {
                            # Módulo não existe, continua
                        }
                    }
                    
                    # Importa novo módulo
                    $vbProject.VBComponents.Import($vbaModulePath) | Out-Null
                    Write-Log "Módulo 'monolithicMod' importado com sucesso" -Level SUCCESS
                    
                    # Salva e fecha
                    $doc.Save()
                    $doc.Close($false)
                    $word.Quit()
                    
                    # Libera objetos COM
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                    
                    Write-Host "[OK] Módulo VBA atualizado com sucesso!" -ForegroundColor Green
                    Write-Log "Módulo VBA importado e Normal.dotm salvo" -Level SUCCESS
                }
            }
            catch {
                Write-Log "Erro ao importar módulo VBA: $_" -Level ERROR
                Write-Host "[AVISO] Não foi possível importar o módulo VBA automaticamente." -ForegroundColor Yellow
                Write-Host ""
                Write-Host "  Importação Manual:" -ForegroundColor Cyan
                Write-Host "    1. Abra o Word" -ForegroundColor Gray
                Write-Host "    2. Pressione Alt + F11" -ForegroundColor Gray
                Write-Host "    3. Arquivo > Importar Arquivo" -ForegroundColor Gray
                Write-Host "    4. Selecione: $vbaModulePath" -ForegroundColor Gray
                Write-Host ""
                
                # Cleanup
                if ($word) {
                    try { $word.Quit() } catch {}
                    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
                }
            }
        }
        else {
            Write-Log "Módulo VBA não encontrado em: $vbaModulePath" -Level WARNING
            Write-Host "[AVISO] Módulo VBA (monolithicMod.bas) não encontrado." -ForegroundColor Yellow
            Write-Host "  Localização esperada: $vbaModulePath" -ForegroundColor Gray
        }
        
        # 7. Detectar e importar personalizações (se disponíveis)
        if (-not $SkipCustomizations) {
            # Usa a raiz do projeto detectada anteriormente
            # Se por algum motivo $projectRoot estiver vazio, usa $SourcePath como fallback
            $configProjectRoot = if ([string]::IsNullOrEmpty($projectRoot)) { $SourcePath } else { $projectRoot }
            
            # Usa a raiz do projeto para encontrar exported-config
            $exportedConfigPath = Join-Path $configProjectRoot "installation\exported-config"
            
            if (Test-CustomizationsAvailable -ImportPath $exportedConfigPath) {
                Write-Host ""
                Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
                Write-Host "  PERSONALIZAÇÕES DO WORD DETECTADAS!" -ForegroundColor White
                Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "* Personalizações exportadas foram encontradas em:" -ForegroundColor Cyan
                Write-Host "   $exportedConfigPath" -ForegroundColor Gray
                Write-Host ""
                Write-Host " Conteúdo que será importado:" -ForegroundColor White
                Write-Host "   • Faixa de Opções Personalizada (Ribbon)" -ForegroundColor Gray
                Write-Host "   • Partes Rápidas (Quick Parts)" -ForegroundColor Gray
                Write-Host "   • Blocos de Construção (Building Blocks)" -ForegroundColor Gray
                Write-Host "   • Temas de Documentos" -ForegroundColor Gray
                Write-Host "   • Template Normal.dotm" -ForegroundColor Gray
                Write-Host ""
                
                $importCustomizations = $true
                if (-not $Force) {
                    $response = Read-Host "Deseja importar estas personalizações agora? (S/N)"
                    $importCustomizations = ($response -match '^[Ss]$')
                }
                
                if ($importCustomizations) {
                    Write-Log "Iniciando importação de personalizações..." -Level INFO
                    $imported = Import-WordCustomizations -ImportPath $exportedConfigPath
                    
                    if ($imported) {
                        Write-Host ""
                        Write-Host "[OK] Personalizações importadas com sucesso!" -ForegroundColor Green
                        Write-Host ""
                        Write-Host "ℹ IMPORTANTE:" -ForegroundColor Cyan
                        Write-Host "   As personalizações serão visíveis na próxima vez" -ForegroundColor Yellow
                        Write-Host "   que você abrir o Microsoft Word." -ForegroundColor Yellow
                        Write-Host ""
                    }
                    else {
                        Write-Host ""
                        Write-Host "[AVISO] Personalizações não foram importadas completamente." -ForegroundColor Yellow
                        Write-Host "  Verifique o log para mais detalhes." -ForegroundColor Yellow
                        Write-Host ""
                    }
                }
                else {
                    Write-Host ""
                    Write-Host "ℹ Importação de personalizações ignorada." -ForegroundColor Cyan
                    Write-Host "  Para importar mais tarde, execute: .\install.ps1" -ForegroundColor Gray
                    Write-Host ""
                    Write-Log "Importação de personalizações ignorada pelo usuário" -Level INFO
                }
            }
            else {
                Write-Log "Pasta 'exported-config' não encontrada - pulando importação" -Level INFO
            }
        }
        else {
            Write-Log "Importação de personalizações desabilitada (-SkipCustomizations)" -Level INFO
        }
        
        # Sucesso!
        $endTime = Get-Date
        $duration = $endTime - $startTime
        
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "║              INSTALAÇÃO CONCLUÍDA COM SUCESSO!                 ║" -ForegroundColor Green
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host " Resumo da Instalação:" -ForegroundColor Cyan
        Write-Host "   • Operações bem-sucedidas: $script:SuccessCount" -ForegroundColor Green
        Write-Host "   • Avisos: $script:WarningCount" -ForegroundColor Yellow
        Write-Host "   • Erros: $script:ErrorCount" -ForegroundColor Red
        Write-Host "   • Tempo decorrido: $($duration.ToString('mm\:ss'))" -ForegroundColor Gray
        Write-Host ""
        
        if ($backupPath) {
            Write-Host " Backup criado em:" -ForegroundColor Cyan
            Write-Host "   $backupPath" -ForegroundColor Gray
            Write-Host ""
        }
        
        Write-Host " Log completo salvo em:" -ForegroundColor Cyan
        Write-Host "   $script:LogFile" -ForegroundColor Gray
        Write-Host ""
        
        Write-Log "=== INSTALAÇÃO CONCLUÍDA COM SUCESSO ===" -Level SUCCESS
        Write-Log "Duração: $($duration.ToString('mm\:ss'))" -Level INFO
    }
    catch {
        $endTime = Get-Date
        $duration = $endTime - $startTime
        
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Red
        Write-Host "║                  ERRO NA INSTALAÇÃO!                           ║" -ForegroundColor Red
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Red
        Write-Host ""
        Write-Host "[ERRO] Erro: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Write-Host " Verifique o arquivo de log para mais detalhes:" -ForegroundColor Yellow
        Write-Host "   $script:LogFile" -ForegroundColor Gray
        Write-Host ""
        
        Write-Log "=== INSTALAÇÃO FALHOU ===" -Level ERROR
        Write-Log "Erro: $($_.Exception.Message)" -Level ERROR
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
        Write-Log "Duração até falha: $($duration.ToString('mm\:ss'))" -Level INFO
        
        # Tenta reverter mudanças se possível
        if ($backupPath -and (Test-Path $backupPath)) {
            Write-Host " Tentando reverter mudanças..." -ForegroundColor Yellow
            try {
                $templatesPath = Join-Path $env:APPDATA "Microsoft\Templates"
                if (Test-Path $templatesPath) {
                    Remove-Item -Path $templatesPath -Recurse -Force
                }
                Rename-Item -Path $backupPath -NewName "Templates" -Force
                Write-Host "[OK] Backup restaurado com sucesso" -ForegroundColor Green
                Write-Log "Backup restaurado após falha na instalação" -Level INFO
            }
            catch {
                Write-Host "[ERRO] Não foi possível restaurar o backup automaticamente" -ForegroundColor Red
                Write-Host "  Backup disponível em: $backupPath" -ForegroundColor Yellow
                Write-Log "Falha ao restaurar backup: $_" -Level ERROR
            }
        }
        
        throw
    }
}

# =============================================================================
# EXECUÇÃO
# =============================================================================

# Verifica se o script está sendo executado como administrador
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║                      [AVISO] AVISO IMPORTANTE [AVISO]                      ║" -ForegroundColor Red
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Red
    Write-Host ""
    Write-Host "[ERRO] Este script está sendo executado com privilégios de Administrador." -ForegroundColor Red
    Write-Host ""
    Write-Host "[AVISO]  PROBLEMA:" -ForegroundColor Yellow
    Write-Host "   Executar como Administrador pode causar problemas de permissões," -ForegroundColor Yellow
    Write-Host "   pois os arquivos serão criados com o proprietário 'Administrador'" -ForegroundColor Yellow
    Write-Host "   ao invés do seu usuário normal." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "[OK]  SOLUÇÃO:" -ForegroundColor Green
    Write-Host "   1. Feche este PowerShell" -ForegroundColor White
    Write-Host "   2. Abra o PowerShell SEM privilégios de administrador:" -ForegroundColor White
    Write-Host "      - Pressione Win + X" -ForegroundColor Gray
    Write-Host "      - Selecione 'Windows PowerShell' (NÃO 'Windows PowerShell (Admin)')" -ForegroundColor Gray
    Write-Host "   3. Execute o script novamente" -ForegroundColor White
    Write-Host ""
    Write-Host "ℹ  Este script NÃO REQUER privilégios de administrador." -ForegroundColor Cyan
    Write-Host "   Todas as operações são realizadas apenas no seu perfil de usuário." -ForegroundColor Cyan
    Write-Host ""
    
    $response = Read-Host "Deseja continuar mesmo assim? (NÃO recomendado) [s/N]"
    if ($response -notmatch '^[Ss]$') {
        Write-Host ""
        Write-Host "Instalação cancelada. Execute novamente sem privilégios de administrador." -ForegroundColor Yellow
        Write-Host ""
        exit 0
    }
    
    Write-Host ""
    Write-Warning "Continuando por solicitação do usuário. Problemas de permissões podem ocorrer."
    Write-Host ""
    Start-Sleep -Seconds 2
}

# Executa instalação
try {
    Install-CHAINSAWConfig
}
catch {
    exit 1
}
