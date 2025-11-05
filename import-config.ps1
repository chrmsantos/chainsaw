# =============================================================================
# CHAINSAW - Script de Importação de Personalizações do Word
# =============================================================================
# Versão: 1.0.0
# Licença: GNU GPLv3
# Autor: Christian Martin dos Santos
# =============================================================================

<#
.SYNOPSIS
    Importa personalizações do Word exportadas anteriormente.

.DESCRIPTION
    Este script importa:
    1. Normal.dotm (template global)
    2. Faixa de Opções Customizada (Ribbon UI)
    3. Blocos de Construção (Building Blocks)
    4. Temas e estilos
    5. Partes rápidas (Quick Parts)
    6. Configurações do registro (opcional)

.PARAMETER ImportPath
    Caminho da pasta com as personalizações exportadas.
    Padrão: .\exported-config

.PARAMETER Force
    Não pede confirmação antes de importar.

.PARAMETER NoBackup
    Não cria backup das configurações atuais.

.EXAMPLE
    .\import-config.ps1
    Importa da pasta padrão com confirmação.

.EXAMPLE
    .\import-config.ps1 -ImportPath "C:\Backup\WordConfig" -Force
    Importa de caminho específico sem confirmação.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ImportPath = ".\exported-config",
    
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$NoBackup
)

$ErrorActionPreference = "Stop"

# =============================================================================
# CONFIGURAÇÕES
# =============================================================================

$script:LogFile = $null
$script:ImportedItems = 0

# Cores
$ColorSuccess = "Green"
$ColorWarning = "Yellow"
$ColorError = "Red"
$ColorInfo = "Cyan"

# Caminhos do Word
$AppDataPath = $env:APPDATA
$LocalAppDataPath = $env:LOCALAPPDATA
$TemplatesPath = Join-Path $AppDataPath "Microsoft\Templates"
$WordSettingsPath = Join-Path $AppDataPath "Microsoft\Word"
$UiCustomizationPath = Join-Path $LocalAppDataPath "Microsoft\Office"

# =============================================================================
# FUNÇÕES DE LOG
# =============================================================================

function Initialize-LogFile {
    try {
        $logDir = Join-Path $env:USERPROFILE "chainsaw\logs"
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $script:LogFile = Join-Path $logDir "import_$timestamp.log"
        
        $header = @"
================================================================================
CHAINSAW - Importação de Personalizações do Word
================================================================================
Data/Hora: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")
Usuário: $env:USERNAME
Computador: $env:COMPUTERNAME
Sistema: $([Environment]::OSVersion.VersionString)
PowerShell: $($PSVersionTable.PSVersion)
Caminho de Importação: $ImportPath
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
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    if ($script:LogFile) {
        try {
            Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
        }
        catch { }
    }
    
    switch ($Level) {
        "SUCCESS" { Write-Host "✓ $Message" -ForegroundColor $ColorSuccess }
        "WARNING" { Write-Host "⚠ $Message" -ForegroundColor $ColorWarning }
        "ERROR"   { Write-Host "✗ $Message" -ForegroundColor $ColorError }
        default   { Write-Host "ℹ $Message" -ForegroundColor $ColorInfo }
    }
}

# =============================================================================
# FUNÇÕES DE VERIFICAÇÃO
# =============================================================================

function Test-WordRunning {
    $wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
    return ($null -ne $wordProcesses -and $wordProcesses.Count -gt 0)
}

function Test-ImportSource {
    if (-not (Test-Path $ImportPath)) {
        Write-Log "Caminho de importação não encontrado: $ImportPath" -Level ERROR
        return $false
    }
    
    # Verifica se há um manifesto
    $manifestPath = Join-Path $ImportPath "MANIFEST.json"
    if (Test-Path $manifestPath) {
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

function Backup-CurrentSettings {
    if ($NoBackup) {
        Write-Log "Backup desabilitado (-NoBackup)" -Level WARNING
        return $true
    }
    
    Write-Log "Criando backup das configurações atuais..." -Level INFO
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = Join-Path $env:USERPROFILE "chainsaw\backups\word-config_$timestamp"
    
    try {
        if (-not (Test-Path $backupPath)) {
            New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
        }
        
        # Backup do Normal.dotm
        $normalPath = Join-Path $TemplatesPath "Normal.dotm"
        if (Test-Path $normalPath) {
            $destNormal = Join-Path $backupPath "Templates"
            New-Item -Path $destNormal -ItemType Directory -Force | Out-Null
            Copy-Item -Path $normalPath -Destination $destNormal -Force
            Write-Log "Normal.dotm backup criado" -Level INFO
        }
        
        # Backup de personalizações UI
        $uiFiles = Get-ChildItem -Path $UiCustomizationPath -Filter "*.officeUI" -Recurse -ErrorAction SilentlyContinue
        if ($uiFiles.Count -gt 0) {
            $destUI = Join-Path $backupPath "OfficeCustomUI"
            New-Item -Path $destUI -ItemType Directory -Force | Out-Null
            foreach ($file in $uiFiles) {
                Copy-Item -Path $file.FullName -Destination (Join-Path $destUI $file.Name) -Force
            }
            Write-Log "Personalizações UI backup criado: $($uiFiles.Count) arquivos" -Level INFO
        }
        
        Write-Log "Backup criado em: $backupPath ✓" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Erro ao criar backup: $_" -Level ERROR
        return $false
    }
}

# =============================================================================
# FUNÇÕES DE IMPORTAÇÃO
# =============================================================================

function Import-NormalTemplate {
    Write-Log "Importando Normal.dotm..." -Level INFO
    
    $sourcePath = Join-Path $ImportPath "Templates\Normal.dotm"
    $destPath = Join-Path $TemplatesPath "Normal.dotm"
    
    if (-not (Test-Path $sourcePath)) {
        Write-Log "Normal.dotm não encontrado no pacote de importação" -Level WARNING
        return $false
    }
    
    try {
        # Garante que a pasta Templates existe
        if (-not (Test-Path $TemplatesPath)) {
            New-Item -Path $TemplatesPath -ItemType Directory -Force | Out-Null
        }
        
        Copy-Item -Path $sourcePath -Destination $destPath -Force
        $script:ImportedItems++
        
        Write-Log "Normal.dotm importado com sucesso ✓" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Erro ao importar Normal.dotm: $_" -Level ERROR
        return $false
    }
}

function Import-BuildingBlocks {
    Write-Log "Importando Building Blocks..." -Level INFO
    
    $sourceManaged = Join-Path $ImportPath "Templates\LiveContent\16\Managed\Word Document Building Blocks"
    $sourceUser = Join-Path $ImportPath "Templates\LiveContent\16\User\Word Document Building Blocks"
    
    $destManaged = Join-Path $TemplatesPath "LiveContent\16\Managed\Word Document Building Blocks"
    $destUser = Join-Path $TemplatesPath "LiveContent\16\User\Word Document Building Blocks"
    
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
                $script:ImportedItems++
            }
            
            Write-Log "Building Blocks do usuário importados: $($files.Count) arquivos" -Level INFO
        }
        catch {
            Write-Log "Erro ao importar Building Blocks do usuário: $_" -Level WARNING
        }
    }
    
    if ($importedCount -gt 0) {
        Write-Log "Building Blocks importados: $importedCount arquivos ✓" -Level SUCCESS
        return $true
    }
    else {
        Write-Log "Nenhum Building Block para importar" -Level INFO
        return $false
    }
}

function Import-DocumentThemes {
    Write-Log "Importando temas de documentos..." -Level INFO
    
    $sourceManaged = Join-Path $ImportPath "Templates\LiveContent\16\Managed\Document Themes"
    $sourceUser = Join-Path $ImportPath "Templates\LiveContent\16\User\Document Themes"
    
    $destManaged = Join-Path $TemplatesPath "LiveContent\16\Managed\Document Themes"
    $destUser = Join-Path $TemplatesPath "LiveContent\16\User\Document Themes"
    
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
        Write-Log "Temas importados: $importedCount arquivos ✓" -Level SUCCESS
        $script:ImportedItems += $importedCount
        return $true
    }
    else {
        Write-Log "Nenhum tema para importar" -Level INFO
        return $false
    }
}

function Import-RibbonCustomization {
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
                (Join-Path $LocalAppDataPath "Microsoft\Office"),
                (Join-Path $AppDataPath "Microsoft\Office")
            )
            
            foreach ($destPath in $possibleDests) {
                if (-not (Test-Path $destPath)) {
                    New-Item -Path $destPath -ItemType Directory -Force | Out-Null
                }
                
                $destFile = Join-Path $destPath $file.Name
                Copy-Item -Path $file.FullName -Destination $destFile -Force
                Write-Log "Ribbon importado para: $destFile" -Level INFO
            }
            
            $script:ImportedItems++
        }
        
        Write-Log "Personalização do Ribbon importada: $($files.Count) arquivos ✓" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Erro ao importar Ribbon: $_" -Level ERROR
        return $false
    }
}

function Import-OfficeCustomUI {
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
        
        $destPath = Join-Path $LocalAppDataPath "Microsoft\Office"
        if (-not (Test-Path $destPath)) {
            New-Item -Path $destPath -ItemType Directory -Force | Out-Null
        }
        
        foreach ($file in $files) {
            $destFile = Join-Path $destPath $file.Name
            Copy-Item -Path $file.FullName -Destination $destFile -Force
            $script:ImportedItems++
        }
        
        Write-Log "Personalizações UI importadas: $($files.Count) arquivos ✓" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Erro ao importar personalizações UI: $_" -Level ERROR
        return $false
    }
}

function Import-RegistrySettings {
    Write-Log "Verificando configurações de registro..." -Level INFO
    
    $sourcePath = Join-Path $ImportPath "Registry"
    
    if (-not (Test-Path $sourcePath)) {
        Write-Log "Nenhuma configuração de registro para importar" -Level INFO
        return $true
    }
    
    try {
        $regFiles = Get-ChildItem -Path $sourcePath -Filter "*.reg" -ErrorAction SilentlyContinue
        
        if ($regFiles.Count -eq 0) {
            Write-Log "Nenhum arquivo de registro encontrado" -Level INFO
            return $false
        }
        
        Write-Host ""
        Write-Host "⚠ AVISO: Importar configurações do registro!" -ForegroundColor Yellow
        Write-Host "   Encontrados $($regFiles.Count) arquivo(s) de registro." -ForegroundColor Yellow
        Write-Host ""
        $response = Read-Host "Deseja importar as configurações do registro? (S/N)"
        
        if ($response -match '^[Ss]$') {
            foreach ($regFile in $regFiles) {
                Write-Log "Importando: $($regFile.Name)" -Level INFO
                $regImport = "reg import `"$($regFile.FullName)`""
                Invoke-Expression $regImport | Out-Null
                $script:ImportedItems++
            }
            Write-Log "Configurações do registro importadas ✓" -Level SUCCESS
            return $true
        }
        else {
            Write-Log "Importação de registro cancelada pelo usuário" -Level WARNING
            return $false
        }
    }
    catch {
        Write-Log "Erro ao importar configurações de registro: $_" -Level ERROR
        return $false
    }
}

# =============================================================================
# FUNÇÃO PRINCIPAL
# =============================================================================

function Import-WordCustomizations {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║        CHAINSAW - Importação de Personalizações do Word       ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    
    # Inicializa log
    Initialize-LogFile | Out-Null
    Write-Log "=== INÍCIO DA IMPORTAÇÃO ===" -Level INFO
    
    # Verifica se Word está em execução
    if (Test-WordRunning) {
        Write-Host ""
        Write-Host "❌ ERRO: O Microsoft Word está em execução!" -ForegroundColor Red
        Write-Host ""
        Write-Host "Para garantir que as personalizações sejam aplicadas corretamente," -ForegroundColor Yellow
        Write-Host "você DEVE fechar o Word antes de importar." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Feche o Word e execute este script novamente." -ForegroundColor Cyan
        Write-Host ""
        
        Write-Log "Importação abortada: Word em execução" -Level ERROR
        return
    }
    
    # Verifica fonte de importação
    if (-not (Test-ImportSource)) {
        Write-Host ""
        Write-Host "❌ Fonte de importação inválida ou não encontrada!" -ForegroundColor Red
        Write-Host ""
        return
    }
    
    # Confirmação
    if (-not $Force) {
        Write-Host ""
        Write-Host "⚠ IMPORTANTE:" -ForegroundColor Yellow
        Write-Host "   Esta operação irá substituir suas personalizações atuais do Word." -ForegroundColor Yellow
        if (-not $NoBackup) {
            Write-Host "   Um backup será criado automaticamente." -ForegroundColor Green
        }
        Write-Host ""
        $response = Read-Host "Deseja continuar? (S/N)"
        if ($response -notmatch '^[Ss]$') {
            Write-Log "Importação cancelada pelo usuário" -Level WARNING
            return
        }
    }
    
    $startTime = Get-Date
    
    try {
        # Backup
        if (-not (Backup-CurrentSettings)) {
            Write-Host ""
            $response = Read-Host "Erro ao criar backup. Continuar mesmo assim? (S/N)"
            if ($response -notmatch '^[Ss]$') {
                Write-Log "Importação cancelada: falha no backup" -Level WARNING
                return
            }
        }
        
        Write-Host ""
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host "  Importando Personalizações" -ForegroundColor White
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host ""
        
        # Importações
        Import-NormalTemplate | Out-Null
        Import-BuildingBlocks | Out-Null
        Import-DocumentThemes | Out-Null
        Import-RibbonCustomization | Out-Null
        Import-OfficeCustomUI | Out-Null
        Import-RegistrySettings | Out-Null
        
        $endTime = Get-Date
        $duration = $endTime - $startTime
        
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "║              IMPORTAÇÃO CONCLUÍDA COM SUCESSO!                 ║" -ForegroundColor Green
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host "📊 Resumo:" -ForegroundColor Cyan
        Write-Host "   • Itens importados: $script:ImportedItems" -ForegroundColor White
        Write-Host "   • Tempo decorrido: $($duration.ToString('mm\:ss'))" -ForegroundColor Gray
        Write-Host ""
        Write-Host "ℹ PRÓXIMO PASSO:" -ForegroundColor Cyan
        Write-Host "   Abra o Microsoft Word para verificar as personalizações." -ForegroundColor White
        Write-Host ""
        Write-Host "📝 Log: $script:LogFile" -ForegroundColor Gray
        Write-Host ""
        
        Write-Log "=== IMPORTAÇÃO CONCLUÍDA COM SUCESSO ===" -Level SUCCESS
        Write-Log "Total de itens importados: $script:ImportedItems" -Level INFO
    }
    catch {
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Red
        Write-Host "║                  ERRO NA IMPORTAÇÃO!                           ║" -ForegroundColor Red
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Red
        Write-Host ""
        Write-Host "❌ Erro: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        
        Write-Log "=== IMPORTAÇÃO FALHOU ===" -Level ERROR
        Write-Log "Erro: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

# =============================================================================
# EXECUÇÃO
# =============================================================================

Import-WordCustomizations
