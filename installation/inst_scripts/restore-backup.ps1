# =============================================================================
# CHAINSAW - Script de Restauração de Backup
# =============================================================================
# Versão: 1.0.0
# Licença: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
# Compatibilidade: Windows 10+, PowerShell 5.1+
# Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
# =============================================================================

<#
.SYNOPSIS
    Restaura configurações do Word a partir de um backup.

.DESCRIPTION
    Este script permite restaurar:
    1. Pasta Templates (estrutura completa)
    2. Normal.dotm
    3. Personalizações de UI do Word
    4. Arquivo stamp.png
    5. Outras personalizações exportadas
    
    O script pode listar backups disponíveis ou restaurar um backup específico.

.PARAMETER BackupPath
    Caminho para o backup a ser restaurado. Se não especificado, lista backups disponíveis.

.PARAMETER BackupName
    Nome do backup (ex: "20251124_133000"). Se não especificado, lista backups disponíveis.

.PARAMETER Force
    Força a restauração sem confirmação do usuário.

.PARAMETER RestoreTemplates
    Restaura apenas a pasta Templates.

.PARAMETER RestoreStamp
    Restaura apenas o arquivo stamp.png.

.PARAMETER RestoreCustomizations
    Restaura apenas personalizações do Word (Normal.dotm, UI).

.PARAMETER List
    Lista todos os backups disponíveis sem restaurar.

.EXAMPLE
    .\restore-backup.ps1 -List
    Lista todos os backups disponíveis.

.EXAMPLE
    .\restore-backup.ps1 -BackupName "20251124_133000"
    Restaura o backup especificado pelo timestamp.

.EXAMPLE
    .\restore-backup.ps1 -RestoreTemplates -Force
    Restaura apenas Templates do backup mais recente sem confirmação.

.NOTES
    Requer que o Microsoft Word esteja fechado.
    Não requer privilégios de administrador.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$BackupPath = "",
    
    [Parameter()]
    [string]$BackupName = "",
    
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$RestoreTemplates,
    
    [Parameter()]
    [switch]$RestoreStamp,
    
    [Parameter()]
    [switch]$RestoreCustomizations,
    
    [Parameter()]
    [switch]$List
)

# =============================================================================
# VARIÁVEIS GLOBAIS
# =============================================================================

$script:LogFile = ""
$script:BackupLocations = @(
    (Join-Path $env:APPDATA "Microsoft"),  # Templates backups
    (Join-Path $env:USERPROFILE "CHAINSAW\backups")  # Word customizations backups
)

# =============================================================================
# FUNÇÕES DE LOG
# =============================================================================

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    if ($script:LogFile -and (Test-Path (Split-Path $script:LogFile -Parent))) {
        Add-Content -Path $script:LogFile -Value $logMessage -Encoding UTF8
    }
    
    # Console output
    switch ($Level) {
        "SUCCESS" { Write-Host "✓ $Message" -ForegroundColor Green }
        "WARNING" { Write-Host "⚠ $Message" -ForegroundColor Yellow }
        "ERROR"   { Write-Host "✗ $Message" -ForegroundColor Red }
        default   { Write-Host "ℹ $Message" -ForegroundColor Cyan }
    }
}

function Initialize-Logging {
    $logDir = Join-Path $env:USERPROFILE "CHAINSAW\logs"
    if (-not (Test-Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $script:LogFile = Join-Path $logDir "restore_$timestamp.log"
    
    Write-Log "=== CHAINSAW - Restauração de Backup ===" -Level INFO
    Write-Log "Data/Hora: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level INFO
    Write-Log "Usuário: $env:USERNAME" -Level INFO
    Write-Log "Computador: $env:COMPUTERNAME" -Level INFO
}

# =============================================================================
# FUNÇÕES DE VERIFICAÇÃO
# =============================================================================

function Test-WordRunning {
    $wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
    return ($null -ne $wordProcesses -and $wordProcesses.Count -gt 0)
}

function Test-Administrator {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# =============================================================================
# FUNÇÕES DE BUSCA DE BACKUPS
# =============================================================================

function Get-AvailableBackups {
    <#
    .SYNOPSIS
        Localiza todos os backups disponíveis no sistema.
    #>
    
    $allBackups = @()
    
    # Busca backups de Templates
    $templatesBackupPath = Join-Path $env:APPDATA "Microsoft"
    $templatesBackups = Get-ChildItem -Path $templatesBackupPath -Directory -Filter "Templates_backup_*" -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending
    
    foreach ($backup in $templatesBackups) {
        $timestamp = $backup.Name -replace "Templates_backup_", ""
        $size = (Get-ChildItem -Path $backup.FullName -Recurse -File | Measure-Object -Property Length -Sum).Sum / 1MB
        
        $allBackups += [PSCustomObject]@{
            Type = "Templates"
            Timestamp = $timestamp
            Date = [DateTime]::ParseExact($timestamp, "yyyyMMdd_HHmmss", $null)
            Path = $backup.FullName
            SizeMB = [math]::Round($size, 2)
            Name = $backup.Name
        }
    }
    
    # Busca backups de personalizações
    $customBackupPath = Join-Path $env:USERPROFILE "CHAINSAW\backups"
    if (Test-Path $customBackupPath) {
        $customBackups = Get-ChildItem -Path $customBackupPath -Directory -Filter "word-customizations_*" -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending
        
        foreach ($backup in $customBackups) {
            $timestamp = $backup.Name -replace "word-customizations_", ""
            $size = (Get-ChildItem -Path $backup.FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum / 1MB
            
            $allBackups += [PSCustomObject]@{
                Type = "Customizations"
                Timestamp = $timestamp
                Date = [DateTime]::ParseExact($timestamp, "yyyyMMdd_HHmmss", $null)
                Path = $backup.FullName
                SizeMB = [math]::Round($size, 2)
                Name = $backup.Name
            }
        }
        
        # Busca backups completos (criados por backup completo)
        $fullBackups = Get-ChildItem -Path $customBackupPath -Directory -Filter "full_backup_*" -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending
        
        foreach ($backup in $fullBackups) {
            $timestamp = $backup.Name -replace "full_backup_", ""
            $size = (Get-ChildItem -Path $backup.FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum / 1MB
            
            $allBackups += [PSCustomObject]@{
                Type = "Full"
                Timestamp = $timestamp
                Date = [DateTime]::ParseExact($timestamp, "yyyyMMdd_HHmmss", $null)
                Path = $backup.FullName
                SizeMB = [math]::Round($size, 2)
                Name = $backup.Name
            }
        }
    }
    
    return $allBackups | Sort-Object Date -Descending
}

function Show-AvailableBackups {
    <#
    .SYNOPSIS
        Exibe lista formatada de backups disponíveis.
    #>
    
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║            BACKUPS DISPONÍVEIS - CHAINSAW                              ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    
    $backups = Get-AvailableBackups
    
    if ($backups.Count -eq 0) {
        Write-Host "Nenhum backup encontrado." -ForegroundColor Yellow
        Write-Host ""
        return
    }
    
    $index = 1
    foreach ($backup in $backups) {
        $dateFormatted = $backup.Date.ToString("dd/MM/yyyy HH:mm:ss")
        $typeColor = switch ($backup.Type) {
            "Templates" { "Green" }
            "Customizations" { "Yellow" }
            "Full" { "Cyan" }
        }
        
        Write-Host "[$index] " -NoNewline -ForegroundColor White
        Write-Host "$($backup.Type.PadRight(15))" -NoNewline -ForegroundColor $typeColor
        Write-Host " | " -NoNewline -ForegroundColor DarkGray
        Write-Host "$dateFormatted" -NoNewline -ForegroundColor White
        Write-Host " | " -NoNewline -ForegroundColor DarkGray
        Write-Host "$($backup.SizeMB) MB" -ForegroundColor Gray
        Write-Host "    Timestamp: $($backup.Timestamp)" -ForegroundColor DarkGray
        Write-Host "    Caminho: $($backup.Path)" -ForegroundColor DarkGray
        Write-Host ""
        
        $index++
    }
    
    Write-Host "Total: $($backups.Count) backup(s)" -ForegroundColor Cyan
    Write-Host ""
}

# =============================================================================
# FUNÇÕES DE RESTAURAÇÃO
# =============================================================================

function Restore-TemplatesFromBackup {
    <#
    .SYNOPSIS
        Restaura a pasta Templates de um backup.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath
    )
    
    Write-Log "Iniciando restauração da pasta Templates..." -Level INFO
    
    $templatesPath = Join-Path $env:APPDATA "Microsoft\Templates"
    
    # Remove pasta Templates atual se existir
    if (Test-Path $templatesPath) {
        Write-Log "Removendo pasta Templates atual..." -Level INFO
        try {
            Remove-Item -Path $templatesPath -Recurse -Force -ErrorAction Stop
            Write-Log "Pasta Templates atual removida" -Level SUCCESS
        }
        catch {
            Write-Log "Erro ao remover pasta Templates: $_" -Level ERROR
            return $false
        }
    }
    
    # Restaura do backup
    try {
        Write-Log "Restaurando Templates de: $SourcePath" -Level INFO
        Copy-Item -Path $SourcePath -Destination $templatesPath -Recurse -Force -ErrorAction Stop
        Write-Log "Pasta Templates restaurada com sucesso" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Erro ao restaurar Templates: $_" -Level ERROR
        return $false
    }
}

function Restore-StampFromBackup {
    <#
    .SYNOPSIS
        Restaura o arquivo stamp.png de um backup.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath
    )
    
    Write-Log "Procurando stamp.png no backup..." -Level INFO
    
    $stampSource = Join-Path $SourcePath "stamp.png"
    if (-not (Test-Path $stampSource)) {
        # Tenta em subpastas comuns
        $stampSource = Get-ChildItem -Path $SourcePath -Filter "stamp.png" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
    }
    
    if (-not $stampSource) {
        Write-Log "stamp.png não encontrado no backup" -Level WARNING
        return $false
    }
    
    $destFolder = Join-Path $env:USERPROFILE "CHAINSAW\assets"
    $destPath = Join-Path $destFolder "stamp.png"
    
    try {
        if (-not (Test-Path $destFolder)) {
            New-Item -Path $destFolder -ItemType Directory -Force | Out-Null
        }
        
        Copy-Item -Path $stampSource -Destination $destPath -Force -ErrorAction Stop
        Write-Log "stamp.png restaurado com sucesso" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Erro ao restaurar stamp.png: $_" -Level ERROR
        return $false
    }
}

function Restore-CustomizationsFromBackup {
    <#
    .SYNOPSIS
        Restaura personalizações do Word de um backup.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath
    )
    
    Write-Log "Iniciando restauração de personalizações..." -Level INFO
    
    $success = $true
    
    # Restaura Normal.dotm
    $normalSource = Join-Path $SourcePath "Templates\Normal.dotm"
    if (Test-Path $normalSource) {
        $templatesPath = Join-Path $env:APPDATA "Microsoft\Templates"
        $normalDest = Join-Path $templatesPath "Normal.dotm"
        
        try {
            if (-not (Test-Path $templatesPath)) {
                New-Item -Path $templatesPath -ItemType Directory -Force | Out-Null
            }
            Copy-Item -Path $normalSource -Destination $normalDest -Force -ErrorAction Stop
            Write-Log "Normal.dotm restaurado" -Level SUCCESS
        }
        catch {
            Write-Log "Erro ao restaurar Normal.dotm: $_" -Level ERROR
            $success = $false
        }
    }
    else {
        Write-Log "Normal.dotm não encontrado no backup" -Level WARNING
    }
    
    # Restaura personalizações de UI
    $uiSource = Join-Path $SourcePath "OfficeCustomUI"
    if (Test-Path $uiSource) {
        $uiDest = Join-Path $env:LOCALAPPDATA "Microsoft\Office"
        
        try {
            $uiFiles = Get-ChildItem -Path $uiSource -Filter "*.officeUI" -ErrorAction SilentlyContinue
            foreach ($file in $uiFiles) {
                # Tenta descobrir a versão do Office procurando pastas
                $officeFolders = Get-ChildItem -Path $uiDest -Directory -Filter "16.0" -ErrorAction SilentlyContinue
                if ($officeFolders) {
                    foreach ($folder in $officeFolders) {
                        $destFile = Join-Path $folder.FullName $file.Name
                        Copy-Item -Path $file.FullName -Destination $destFile -Force -ErrorAction SilentlyContinue
                    }
                }
            }
            Write-Log "Personalizações de UI restauradas" -Level SUCCESS
        }
        catch {
            Write-Log "Erro ao restaurar personalizações UI: $_" -Level ERROR
            $success = $false
        }
    }
    
    return $success
}

function Restore-FullBackup {
    <#
    .SYNOPSIS
        Restaura um backup completo.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$BackupPath
    )
    
    Write-Log "Iniciando restauração completa..." -Level INFO
    
    $results = @{
        Templates = $false
        Stamp = $false
        Customizations = $false
    }
    
    # Restaura Templates
    $templatesBackup = Join-Path $BackupPath "Templates"
    if (Test-Path $templatesBackup) {
        $results.Templates = Restore-TemplatesFromBackup -SourcePath $templatesBackup
    }
    else {
        Write-Log "Backup de Templates não encontrado em: $templatesBackup" -Level WARNING
    }
    
    # Restaura stamp.png
    $results.Stamp = Restore-StampFromBackup -SourcePath $BackupPath
    
    # Restaura personalizações
    $customBackup = Join-Path $BackupPath "Customizations"
    if (Test-Path $customBackup) {
        $results.Customizations = Restore-CustomizationsFromBackup -SourcePath $customBackup
    }
    
    return $results
}

# =============================================================================
# FUNÇÃO PRINCIPAL
# =============================================================================

function Main {
    # Inicializa logging
    Initialize-Logging
    
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║         CHAINSAW - Restauração de Configurações do Word                ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    
    # Lista backups se solicitado
    if ($List) {
        Show-AvailableBackups
        return
    }
    
    # Verifica se Word está aberto
    if (Test-WordRunning) {
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "║          [AVISO] MICROSOFT WORD ABERTO [AVISO]                 ║" -ForegroundColor Yellow
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "O Microsoft Word deve ser fechado antes da restauração." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Por favor:" -ForegroundColor White
        Write-Host "  1. Salve todos os documentos abertos" -ForegroundColor Gray
        Write-Host "  2. Feche completamente o Microsoft Word" -ForegroundColor Gray
        Write-Host "  3. Execute este script novamente" -ForegroundColor Gray
        Write-Host ""
        Write-Log "Restauração cancelada - Word está aberto" -Level ERROR
        exit 1
    }
    
    # Busca backups disponíveis
    $availableBackups = Get-AvailableBackups
    
    if ($availableBackups.Count -eq 0) {
        Write-Host "Nenhum backup encontrado." -ForegroundColor Yellow
        Write-Host ""
        Write-Log "Nenhum backup encontrado no sistema" -Level WARNING
        return
    }
    
    # Determina qual backup usar
    $selectedBackup = $null
    
    if ($BackupPath -and (Test-Path $BackupPath)) {
        # Usa caminho especificado
        $selectedBackup = $availableBackups | Where-Object { $_.Path -eq $BackupPath } | Select-Object -First 1
        if (-not $selectedBackup) {
            # Cria objeto para backup personalizado
            $selectedBackup = [PSCustomObject]@{
                Type = "Custom"
                Path = $BackupPath
                Name = Split-Path $BackupPath -Leaf
            }
        }
    }
    elseif ($BackupName) {
        # Busca por timestamp
        $selectedBackup = $availableBackups | Where-Object { $_.Timestamp -eq $BackupName } | Select-Object -First 1
        
        if (-not $selectedBackup) {
            Write-Host "Backup não encontrado: $BackupName" -ForegroundColor Red
            Write-Host ""
            Write-Host "Use -List para ver backups disponíveis" -ForegroundColor Yellow
            return
        }
    }
    else {
        # Mostra lista e pede seleção
        Show-AvailableBackups
        
        Write-Host "Selecione o número do backup para restaurar (ou Enter para cancelar): " -NoNewline -ForegroundColor Yellow
        $selection = Read-Host
        
        if ([string]::IsNullOrWhiteSpace($selection)) {
            Write-Log "Restauração cancelada pelo usuário" -Level WARNING
            return
        }
        
        $index = 0
        if ([int]::TryParse($selection, [ref]$index) -and $index -gt 0 -and $index -le $availableBackups.Count) {
            $selectedBackup = $availableBackups[$index - 1]
        }
        else {
            Write-Host "Seleção inválida" -ForegroundColor Red
            return
        }
    }
    
    # Confirmação
    if (-not $Force) {
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
        Write-Host "║                    CONFIRMAÇÃO                                  ║" -ForegroundColor Yellow
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Backup selecionado:" -ForegroundColor White
        Write-Host "  Tipo: $($selectedBackup.Type)" -ForegroundColor Cyan
        Write-Host "  Data: $($selectedBackup.Date)" -ForegroundColor Cyan
        Write-Host "  Caminho: $($selectedBackup.Path)" -ForegroundColor Gray
        Write-Host ""
        Write-Host "AVISO: Esta operação irá sobrescrever as configurações atuais!" -ForegroundColor Red
        Write-Host ""
        
        $response = Read-Host "Deseja continuar? (S/N)"
        if ($response -notmatch '^[Ss]$') {
            Write-Log "Restauração cancelada pelo usuário" -Level WARNING
            return
        }
    }
    
    # Executa restauração
    Write-Host ""
    Write-Host "Iniciando restauração..." -ForegroundColor Cyan
    Write-Host ""
    
    $startTime = Get-Date
    $success = $false
    
    try {
        if ($selectedBackup.Type -eq "Templates" -or $RestoreTemplates) {
            $success = Restore-TemplatesFromBackup -SourcePath $selectedBackup.Path
        }
        elseif ($selectedBackup.Type -eq "Customizations" -or $RestoreCustomizations) {
            $success = Restore-CustomizationsFromBackup -SourcePath $selectedBackup.Path
        }
        elseif ($selectedBackup.Type -eq "Full") {
            $results = Restore-FullBackup -BackupPath $selectedBackup.Path
            $success = $results.Templates -or $results.Stamp -or $results.Customizations
        }
        elseif ($RestoreStamp) {
            $success = Restore-StampFromBackup -SourcePath $selectedBackup.Path
        }
        else {
            # Tenta restaurar tudo que encontrar
            $results = Restore-FullBackup -BackupPath $selectedBackup.Path
            $success = $results.Templates -or $results.Stamp -or $results.Customizations
        }
        
        $duration = (Get-Date) - $startTime
        
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor $(if($success){"Green"}else{"Red"})
        Write-Host "║                    RESULTADO                                    ║" -ForegroundColor $(if($success){"Green"}else{"Red"})
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor $(if($success){"Green"}else{"Red"})
        Write-Host ""
        
        if ($success) {
            Write-Host "✓ Restauração concluída com sucesso!" -ForegroundColor Green
            Write-Log "Restauração concluída com sucesso" -Level SUCCESS
        }
        else {
            Write-Host "✗ Restauração falhou" -ForegroundColor Red
            Write-Log "Restauração falhou" -Level ERROR
        }
        
        Write-Host "  Duração: $($duration.ToString('mm\:ss'))" -ForegroundColor Gray
        Write-Host "  Log: $script:LogFile" -ForegroundColor Gray
        Write-Host ""
    }
    catch {
        Write-Host ""
        Write-Host "✗ Erro durante restauração: $_" -ForegroundColor Red
        Write-Log "Erro durante restauração: $_" -Level ERROR
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
    }
}

# =============================================================================
# EXECUÇÃO
# =============================================================================

Main
