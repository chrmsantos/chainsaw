# =============================================================================
# CHAINSAW - Script de Instalação de Configurações do Word
# =============================================================================
# Versão: 1.0.0
# Licença: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
# Compatibilidade: Windows 10+, PowerShell 5.1+
# Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
# =============================================================================

<#
.SYNOPSIS
    Instala as configurações do Word do sistema Chainsaw para o usuário atual.

.DESCRIPTION
    Este script realiza as seguintes operações:
    1. Copia o arquivo stamp.png da rede para a pasta do usuário
    2. Faz backup da pasta Templates atual
    3. Copia os novos Templates da rede
    4. Registra todas as operações em arquivo de log

.PARAMETER SourcePath
    Caminho de rede base. Padrão: \\strqnapmain\Dir. Legislativa\_Christian261\chainsaw

.PARAMETER Force
    Força a instalação sem confirmação do usuário.

.PARAMETER NoBackup
    Não cria backup da pasta Templates existente (não recomendado).

.EXAMPLE
    .\install.ps1
    Executa a instalação com confirmação do usuário.

.EXAMPLE
    .\install.ps1 -Force
    Executa a instalação sem confirmação.

.NOTES
    Requer permissões de leitura no caminho de rede e escrita nas pastas do usuário.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$SourcePath = "\\strqnapmain\Dir. Legislativa\_Christian261\chainsaw",
    
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$NoBackup
)

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
        $logDir = Join-Path $env:USERPROFILE "chainsaw\logs"
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
                Write-Host "✓ $Message" -ForegroundColor $ColorSuccess
                $script:SuccessCount++
            }
            "WARNING" {
                Write-Host "⚠ $Message" -ForegroundColor $ColorWarning
                $script:WarningCount++
            }
            "ERROR" {
                Write-Host "✗ $Message" -ForegroundColor $ColorError
                $script:ErrorCount++
            }
            default {
                Write-Host "ℹ $Message" -ForegroundColor $ColorInfo
            }
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
        Write-Log "Sistema operacional: Windows $($osVersion.Major).$($osVersion.Minor) ✓" -Level SUCCESS
    }
    
    # Verifica versão do PowerShell
    $psVersion = $PSVersionTable.PSVersion
    if ($psVersion.Major -lt 5) {
        Write-Log "PowerShell 5.1 ou superior é necessário. Versão detectada: $($psVersion.ToString())" -Level ERROR
        $allOk = $false
    }
    else {
        Write-Log "PowerShell versão: $($psVersion.ToString()) ✓" -Level SUCCESS
    }
    
    # Verifica acesso ao caminho de rede
    Write-Log "Verificando acesso ao caminho de rede: $SourcePath" -Level INFO
    if (-not (Test-Path $SourcePath)) {
        Write-Log "Não foi possível acessar o caminho de rede: $SourcePath" -Level ERROR
        Write-Log "Verifique se você está conectado à rede e tem permissões de acesso." -Level ERROR
        $allOk = $false
    }
    else {
        Write-Log "Acesso ao caminho de rede confirmado ✓" -Level SUCCESS
    }
    
    # Verifica permissões de escrita no perfil do usuário
    $testFile = Join-Path $env:USERPROFILE "chainsaw_test_$(Get-Date -Format 'yyyyMMddHHmmss').tmp"
    try {
        [System.IO.File]::WriteAllText($testFile, "test")
        Remove-Item $testFile -Force -ErrorAction SilentlyContinue
        Write-Log "Permissões de escrita no perfil do usuário confirmadas ✓" -Level SUCCESS
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
        [ref]$SourceTemplatesFolder
    )
    
    Write-Log "Verificando arquivos de origem..." -Level INFO
    
    $allOk = $true
    
    # Verifica arquivo stamp.png
    $stampPath = Join-Path $SourcePath "assets\stamp.png"
    if (Test-Path $stampPath) {
        $SourceStampFile.Value = $stampPath
        Write-Log "Arquivo stamp.png encontrado ✓" -Level SUCCESS
    }
    else {
        Write-Log "Arquivo não encontrado: $stampPath" -Level ERROR
        $allOk = $false
    }
    
    # Verifica pasta Templates
    $templatesPath = Join-Path $SourcePath "configs\Templates"
    if (Test-Path $templatesPath) {
        $SourceTemplatesFolder.Value = $templatesPath
        Write-Log "Pasta Templates encontrada ✓" -Level SUCCESS
    }
    else {
        Write-Log "Pasta não encontrada: $templatesPath" -Level ERROR
        $allOk = $false
    }
    
    return $allOk
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
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupName = "Templates_backup_$timestamp"
    $backupPath = Join-Path (Split-Path $SourceFolder -Parent) $backupName
    
    Write-Log "Criando backup da pasta Templates..." -Level INFO
    Write-Log "Origem: $SourceFolder" -Level INFO
    Write-Log "Destino: $backupPath" -Level INFO
    
    try {
        # Renomeia a pasta existente
        Rename-Item -Path $SourceFolder -NewName $backupName -Force -ErrorAction Stop
        Write-Log "Backup criado com sucesso: $backupName ✓" -Level SUCCESS
        return $backupPath
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
    
    $destFolder = Join-Path $env:USERPROFILE "chainsaw\assets"
    $destFile = Join-Path $destFolder "stamp.png"
    
    Write-Log "Copiando arquivo stamp.png..." -Level INFO
    Write-Log "Origem: $SourceFile" -Level INFO
    Write-Log "Destino: $destFile" -Level INFO
    
    try {
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
                Write-Log "Arquivo stamp.png copiado com sucesso ✓" -Level SUCCESS
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
        Copia a pasta Templates da rede para o perfil do usuário.
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
        Write-Log "Pasta Templates copiada com sucesso ($copiedItems arquivos) ✓" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Erro ao copiar pasta Templates: $_" -Level ERROR
        throw
    }
}

# =============================================================================
# FUNÇÃO PRINCIPAL
# =============================================================================

function Install-ChainsawConfig {
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
        Write-Host "📝 Arquivo de log: $script:LogFile" -ForegroundColor Gray
        Write-Host ""
    }
    
    $startTime = Get-Date
    Write-Log "=== INÍCIO DA INSTALAÇÃO ===" -Level INFO
    
    try {
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
        
        if (-not (Test-SourceFiles -SourceStampFile ([ref]$sourceStampFile) -SourceTemplatesFolder ([ref]$sourceTemplatesFolder))) {
            throw "Arquivos de origem não encontrados. Verifique os erros acima."
        }
        
        # 3. Confirmação do usuário
        if (-not $Force) {
            Write-Host ""
            Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
            Write-Host "  CONFIRMAÇÃO" -ForegroundColor White
            Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
            Write-Host ""
            Write-Host "As seguintes operações serão realizadas:" -ForegroundColor Yellow
            Write-Host "  1. Copiar stamp.png para: $env:USERPROFILE\chainsaw\assets\" -ForegroundColor White
            Write-Host "  2. Fazer backup da pasta Templates atual (se existir)" -ForegroundColor White
            Write-Host "  3. Copiar nova pasta Templates da rede" -ForegroundColor White
            Write-Host ""
            
            $response = Read-Host "Deseja continuar? (S/N)"
            if ($response -notmatch '^[Ss]$') {
                Write-Log "Instalação cancelada pelo usuário." -Level WARNING
                return
            }
        }
        
        # 4. Copiar arquivo stamp.png
        Write-Host ""
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host "  ETAPA 3: Cópia do Arquivo stamp.png" -ForegroundColor White
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host ""
        
        Copy-StampFile -SourceFile $sourceStampFile | Out-Null
        
        # 5. Backup da pasta Templates
        $templatesPath = Join-Path $env:APPDATA "Microsoft\Templates"
        $backupPath = $null
        
        if (-not $NoBackup) {
            Write-Host ""
            Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
            Write-Host "  ETAPA 4: Backup da Pasta Templates" -ForegroundColor White
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
        
        # 6. Copiar pasta Templates
        Write-Host ""
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host "  ETAPA 5: Cópia da Pasta Templates" -ForegroundColor White
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host ""
        
        Copy-TemplatesFolder -SourceFolder $sourceTemplatesFolder -DestFolder $templatesPath | Out-Null
        
        # Sucesso!
        $endTime = Get-Date
        $duration = $endTime - $startTime
        
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "║              INSTALAÇÃO CONCLUÍDA COM SUCESSO!                 ║" -ForegroundColor Green
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host "📊 Resumo da Instalação:" -ForegroundColor Cyan
        Write-Host "   • Operações bem-sucedidas: $script:SuccessCount" -ForegroundColor Green
        Write-Host "   • Avisos: $script:WarningCount" -ForegroundColor Yellow
        Write-Host "   • Erros: $script:ErrorCount" -ForegroundColor Red
        Write-Host "   • Tempo decorrido: $($duration.ToString('mm\:ss'))" -ForegroundColor Gray
        Write-Host ""
        
        if ($backupPath) {
            Write-Host "💾 Backup criado em:" -ForegroundColor Cyan
            Write-Host "   $backupPath" -ForegroundColor Gray
            Write-Host ""
        }
        
        Write-Host "📝 Log completo salvo em:" -ForegroundColor Cyan
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
        Write-Host "❌ Erro: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Write-Host "📝 Verifique o arquivo de log para mais detalhes:" -ForegroundColor Yellow
        Write-Host "   $script:LogFile" -ForegroundColor Gray
        Write-Host ""
        
        Write-Log "=== INSTALAÇÃO FALHOU ===" -Level ERROR
        Write-Log "Erro: $($_.Exception.Message)" -Level ERROR
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
        Write-Log "Duração até falha: $($duration.ToString('mm\:ss'))" -Level INFO
        
        # Tenta reverter mudanças se possível
        if ($backupPath -and (Test-Path $backupPath)) {
            Write-Host "🔄 Tentando reverter mudanças..." -ForegroundColor Yellow
            try {
                $templatesPath = Join-Path $env:APPDATA "Microsoft\Templates"
                if (Test-Path $templatesPath) {
                    Remove-Item -Path $templatesPath -Recurse -Force
                }
                Rename-Item -Path $backupPath -NewName "Templates" -Force
                Write-Host "✓ Backup restaurado com sucesso" -ForegroundColor Green
                Write-Log "Backup restaurado após falha na instalação" -Level INFO
            }
            catch {
                Write-Host "✗ Não foi possível restaurar o backup automaticamente" -ForegroundColor Red
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
    Write-Host "║                      ⚠ AVISO IMPORTANTE ⚠                      ║" -ForegroundColor Red
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Red
    Write-Host ""
    Write-Host "❌ Este script está sendo executado com privilégios de Administrador." -ForegroundColor Red
    Write-Host ""
    Write-Host "⚠  PROBLEMA:" -ForegroundColor Yellow
    Write-Host "   Executar como Administrador pode causar problemas de permissões," -ForegroundColor Yellow
    Write-Host "   pois os arquivos serão criados com o proprietário 'Administrador'" -ForegroundColor Yellow
    Write-Host "   ao invés do seu usuário normal." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "✓  SOLUÇÃO:" -ForegroundColor Green
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
    Install-ChainsawConfig
}
catch {
    exit 1
}
