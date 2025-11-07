# =============================================================================
# CHAINSAW - Script de Exportação de Personalizações do Word
# =============================================================================
# Versão: 1.0.0
# Licença: GNU GPLv3
# Autor: Christian Martin dos Santos
# =============================================================================

# Maximiza a janela do PowerShell
if ($Host.Name -eq "ConsoleHost") {
    $psWindow = (Get-Host).UI.RawUI
    $newSize = $psWindow.BufferSize
    $newSize.Width = 120
    $newSize.Height = 9999
    try {
        $psWindow.BufferSize = $newSize
        $psWindow.WindowSize = $psWindow.MaxPhysicalWindowSize
    } catch {
        # Ignora erros se não for possível maximizar
    }
}

<#
.SYNOPSIS
    Exporta todas as personalizações do Word do usuário atual.

.DESCRIPTION
    Este script exporta:
    1. Normal.dotm (template global com macros e personalizações)
    2. Faixa de Opções Customizada (Ribbon UI)
    3. Blocos de Construção (Building Blocks)
    4. Configurações de temas e estilos
    5. Partes rápidas (Quick Parts)
    
.PARAMETER ExportPath
    Caminho onde as personalizações serão exportadas.
    Padrão: .\exported-config

.PARAMETER IncludeRegistry
    Exporta também configurações do registro do Word.

.EXAMPLE
    .\export-config.ps1
    Exporta para a pasta padrão.

.EXAMPLE
    .\export-config.ps1 -ExportPath "C:\Backup\WordConfig"
    Exporta para caminho específico.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ExportPath = ".\exported-config",
    
    [Parameter()]
    [switch]$IncludeRegistry
)

$ErrorActionPreference = "Stop"

# =============================================================================
# CONFIGURAÇÕES
# =============================================================================

$script:LogFile = $null
$script:ExportedItems = @()

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
        $logDir = Join-Path $ExportPath "logs"
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $script:LogFile = Join-Path $logDir "export_$timestamp.log"
        
        $header = @"
================================================================================
CHAINSAW - Exportação de Personalizações do Word
================================================================================
Data/Hora: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")
Usuário: $env:USERNAME
Computador: $env:COMPUTERNAME
Sistema: $([Environment]::OSVersion.VersionString)
PowerShell: $($PSVersionTable.PSVersion)
Caminho de Exportação: $ExportPath
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
    <#
    .SYNOPSIS
        Verifica se o Word está em execução.
    #>
    $wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
    return ($null -ne $wordProcesses -and $wordProcesses.Count -gt 0)
}

function Get-WordVersion {
    <#
    .SYNOPSIS
        Obtém a versão do Word instalada.
    #>
    try {
        $wordPath = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE" -ErrorAction Stop
        if ($wordPath) {
            $versionInfo = (Get-Item $wordPath.'(default)').VersionInfo
            return $versionInfo.ProductVersion
        }
    }
    catch { }
    
    return $null
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
    Write-Host "║                          ⚠ ATENÇÃO ⚠                          ║" -ForegroundColor Yellow
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
    $response = Read-Host "Deseja FECHAR o Word e continuar a exportação? (S/N)"
    
    if ($response -notmatch '^[Ss]$') {
        Write-Host ""
        Write-Host "✓ Exportação cancelada pelo usuário" -ForegroundColor Cyan
        Write-Host "  Salve seus documentos e execute o script novamente quando estiver pronto." -ForegroundColor Gray
        Write-Host ""
        Write-Log "Exportação cancelada - usuário optou por não fechar o Word" -Level WARNING
        return $false
    }
    
    # Usuário confirmou - fecha o Word
    Write-Host ""
    Write-Host "Fechando Microsoft Word..." -ForegroundColor Cyan
    Write-Log "Usuário autorizou o fechamento do Word" -Level INFO
    
    if (Stop-WordProcesses -Force) {
        Write-Host "✓ Word fechado com sucesso" -ForegroundColor Green
        Write-Host ""
        # Aguarda um pouco para garantir que recursos foram liberados
        Start-Sleep -Seconds 2
        return $true
    }
    else {
        Write-Host "✗ Não foi possível fechar o Word completamente" -ForegroundColor Red
        Write-Host ""
        Write-Log "Falha ao fechar Word - cancelando exportação" -Level ERROR
        
        $retry = Read-Host "Deseja tentar novamente? (S/N)"
        if ($retry -match '^[Ss]$') {
            return Confirm-CloseWord # Recursão para tentar novamente
        }
        return $false
    }
}

# =============================================================================
# FUNÇÕES DE EXPORTAÇÃO
# =============================================================================

function Export-NormalTemplate {
    <#
    .SYNOPSIS
        Exporta o template Normal.dotm.
    #>
    Write-Log "Exportando Normal.dotm..." -Level INFO
    
    $normalPath = Join-Path $TemplatesPath "Normal.dotm"
    $destPath = Join-Path $ExportPath "Templates"
    
    if (-not (Test-Path $normalPath)) {
        Write-Log "Normal.dotm não encontrado em: $normalPath" -Level WARNING
        return $false
    }
    
    try {
        if (-not (Test-Path $destPath)) {
            New-Item -Path $destPath -ItemType Directory -Force | Out-Null
        }
        
        Copy-Item -Path $normalPath -Destination $destPath -Force
        
        $script:ExportedItems += [PSCustomObject]@{
            Type = "Normal Template"
            Source = $normalPath
            Destination = Join-Path $destPath "Normal.dotm"
            Size = (Get-Item $normalPath).Length
        }
        
        Write-Log "Normal.dotm exportado com sucesso ✓" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Erro ao exportar Normal.dotm: $_" -Level ERROR
        return $false
    }
}

function Export-BuildingBlocks {
    <#
    .SYNOPSIS
        Exporta os blocos de construção (Building Blocks).
    #>
    Write-Log "Exportando Building Blocks..." -Level INFO
    
    $buildingBlocksPath = Join-Path $TemplatesPath "LiveContent\16\Managed\Word Document Building Blocks"
    $userBuildingBlocksPath = Join-Path $TemplatesPath "LiveContent\16\User\Word Document Building Blocks"
    $destPath = Join-Path $ExportPath "Templates\LiveContent\16"
    
    $exportedCount = 0
    
    # Exporta Building Blocks gerenciados (sistema)
    if (Test-Path $buildingBlocksPath) {
        try {
            $destManaged = Join-Path $destPath "Managed\Word Document Building Blocks"
            if (-not (Test-Path $destManaged)) {
                New-Item -Path $destManaged -ItemType Directory -Force | Out-Null
            }
            
            $files = Get-ChildItem -Path $buildingBlocksPath -Recurse -File
            foreach ($file in $files) {
                $relativePath = $file.FullName.Substring($buildingBlocksPath.Length + 1)
                $destFile = Join-Path $destManaged $relativePath
                $destDir = Split-Path $destFile -Parent
                
                if (-not (Test-Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                }
                
                Copy-Item -Path $file.FullName -Destination $destFile -Force
                $exportedCount++
            }
            
            Write-Log "Building Blocks gerenciados: $($files.Count) arquivos" -Level INFO
        }
        catch {
            Write-Log "Erro ao exportar Building Blocks gerenciados: $_" -Level WARNING
        }
    }
    
    # Exporta Building Blocks do usuário
    if (Test-Path $userBuildingBlocksPath) {
        try {
            $destUser = Join-Path $destPath "User\Word Document Building Blocks"
            if (-not (Test-Path $destUser)) {
                New-Item -Path $destUser -ItemType Directory -Force | Out-Null
            }
            
            $files = Get-ChildItem -Path $userBuildingBlocksPath -Recurse -File
            foreach ($file in $files) {
                $relativePath = $file.FullName.Substring($userBuildingBlocksPath.Length + 1)
                $destFile = Join-Path $destUser $relativePath
                $destDir = Split-Path $destFile -Parent
                
                if (-not (Test-Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                }
                
                Copy-Item -Path $file.FullName -Destination $destFile -Force
                $exportedCount++
                
                $script:ExportedItems += [PSCustomObject]@{
                    Type = "Building Block (User)"
                    Source = $file.FullName
                    Destination = $destFile
                    Size = $file.Length
                }
            }
            
            Write-Log "Building Blocks do usuário: $($files.Count) arquivos" -Level INFO
        }
        catch {
            Write-Log "Erro ao exportar Building Blocks do usuário: $_" -Level WARNING
        }
    }
    
    if ($exportedCount -gt 0) {
        Write-Log "Building Blocks exportados: $exportedCount arquivos ✓" -Level SUCCESS
        return $true
    }
    else {
        Write-Log "Nenhum Building Block encontrado" -Level WARNING
        return $false
    }
}

function Export-DocumentThemes {
    <#
    .SYNOPSIS
        Exporta temas de documentos personalizados.
    #>
    Write-Log "Exportando temas de documentos..." -Level INFO
    
    $themesPath = Join-Path $TemplatesPath "LiveContent\16\Managed\Document Themes"
    $userThemesPath = Join-Path $TemplatesPath "LiveContent\16\User\Document Themes"
    $destPath = Join-Path $ExportPath "Templates\LiveContent\16"
    
    $exportedCount = 0
    
    # Temas gerenciados
    if (Test-Path $themesPath) {
        try {
            $destManaged = Join-Path $destPath "Managed\Document Themes"
            if (-not (Test-Path $destManaged)) {
                New-Item -Path $destManaged -ItemType Directory -Force | Out-Null
            }
            
            $files = Get-ChildItem -Path $themesPath -Recurse -File
            foreach ($file in $files) {
                $relativePath = $file.FullName.Substring($themesPath.Length + 1)
                $destFile = Join-Path $destManaged $relativePath
                $destDir = Split-Path $destFile -Parent
                
                if (-not (Test-Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                }
                
                Copy-Item -Path $file.FullName -Destination $destFile -Force
                $exportedCount++
            }
        }
        catch {
            Write-Log "Erro ao exportar temas gerenciados: $_" -Level WARNING
        }
    }
    
    # Temas do usuário
    if (Test-Path $userThemesPath) {
        try {
            $destUser = Join-Path $destPath "User\Document Themes"
            if (-not (Test-Path $destUser)) {
                New-Item -Path $destUser -ItemType Directory -Force | Out-Null
            }
            
            $files = Get-ChildItem -Path $userThemesPath -Recurse -File
            foreach ($file in $files) {
                $relativePath = $file.FullName.Substring($userThemesPath.Length + 1)
                $destFile = Join-Path $destUser $relativePath
                $destDir = Split-Path $destFile -Parent
                
                if (-not (Test-Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                }
                
                Copy-Item -Path $file.FullName -Destination $destFile -Force
                $exportedCount++
            }
        }
        catch {
            Write-Log "Erro ao exportar temas do usuário: $_" -Level WARNING
        }
    }
    
    if ($exportedCount -gt 0) {
        Write-Log "Temas exportados: $exportedCount arquivos ✓" -Level SUCCESS
        return $true
    }
    else {
        Write-Log "Nenhum tema personalizado encontrado" -Level INFO
        return $false
    }
}

function Export-RibbonCustomization {
    <#
    .SYNOPSIS
        Exporta personalizações da Faixa de Opções (Ribbon).
    #>
    Write-Log "Exportando personalização da Faixa de Opções..." -Level INFO
    
    # A personalização do Ribbon é armazenada em diferentes locais dependendo da versão
    $possiblePaths = @(
        (Join-Path $LocalAppDataPath "Microsoft\Office\Word.officeUI"),
        (Join-Path $AppDataPath "Microsoft\Office\Word.officeUI"),
        (Join-Path $LocalAppDataPath "Microsoft\Office\16.0\Word.officeUI")
    )
    
    $destPath = Join-Path $ExportPath "RibbonCustomization"
    $exportedAny = $false
    
    foreach ($uiPath in $possiblePaths) {
        if (Test-Path $uiPath) {
            try {
                if (-not (Test-Path $destPath)) {
                    New-Item -Path $destPath -ItemType Directory -Force | Out-Null
                }
                
                $fileName = Split-Path $uiPath -Leaf
                $destFile = Join-Path $destPath $fileName
                Copy-Item -Path $uiPath -Destination $destFile -Force
                
                $script:ExportedItems += [PSCustomObject]@{
                    Type = "Ribbon Customization"
                    Source = $uiPath
                    Destination = $destFile
                    Size = (Get-Item $uiPath).Length
                }
                
                Write-Log "Personalização do Ribbon exportada: $fileName ✓" -Level SUCCESS
                $exportedAny = $true
            }
            catch {
                Write-Log "Erro ao exportar $uiPath : $_" -Level WARNING
            }
        }
    }
    
    if (-not $exportedAny) {
        Write-Log "Nenhuma personalização do Ribbon encontrada" -Level INFO
    }
    
    return $exportedAny
}

function Export-OfficeCustomUI {
    <#
    .SYNOPSIS
        Exporta arquivos de personalização da interface do Office.
    #>
    Write-Log "Exportando personalizações da interface..." -Level INFO
    
    $customUIPath = Join-Path $LocalAppDataPath "Microsoft\Office"
    $destPath = Join-Path $ExportPath "OfficeCustomUI"
    
    try {
        # Procura por arquivos .officeUI
        $customFiles = Get-ChildItem -Path $customUIPath -Filter "*.officeUI" -Recurse -ErrorAction SilentlyContinue
        
        if ($customFiles.Count -gt 0) {
            if (-not (Test-Path $destPath)) {
                New-Item -Path $destPath -ItemType Directory -Force | Out-Null
            }
            
            foreach ($file in $customFiles) {
                $destFile = Join-Path $destPath $file.Name
                Copy-Item -Path $file.FullName -Destination $destFile -Force
                
                $script:ExportedItems += [PSCustomObject]@{
                    Type = "Office Custom UI"
                    Source = $file.FullName
                    Destination = $destFile
                    Size = $file.Length
                }
            }
            
            Write-Log "Personalizações UI exportadas: $($customFiles.Count) arquivos ✓" -Level SUCCESS
            return $true
        }
        else {
            Write-Log "Nenhum arquivo de personalização UI encontrado" -Level INFO
            return $false
        }
    }
    catch {
        Write-Log "Erro ao exportar personalizações UI: $_" -Level WARNING
        return $false
    }
}

function Export-QuickAccessToolbar {
    <#
    .SYNOPSIS
        Exporta configurações da Barra de Ferramentas de Acesso Rápido.
    #>
    Write-Log "Exportando Barra de Ferramentas de Acesso Rápido..." -Level INFO
    
    # A QAT é armazenada no arquivo .officeUI ou no registro
    # Já será exportada pela função Export-OfficeCustomUI
    
    Write-Log "QAT incluída nas personalizações UI" -Level INFO
    return $true
}

function Export-RegistrySettings {
    <#
    .SYNOPSIS
        Exporta configurações do Word do registro.
    #>
    if (-not $IncludeRegistry) {
        Write-Log "Exportação do registro desabilitada (use -IncludeRegistry)" -Level INFO
        return $true
    }
    
    Write-Log "Exportando configurações do registro..." -Level INFO
    
    $regPaths = @(
        "HKCU:\Software\Microsoft\Office\16.0\Word",
        "HKCU:\Software\Microsoft\Office\Common\Toolbars",
        "HKCU:\Software\Microsoft\Office\16.0\Common\Toolbars"
    )
    
    $destPath = Join-Path $ExportPath "Registry"
    $exportedAny = $false
    
    foreach ($regPath in $regPaths) {
        if (Test-Path $regPath) {
            try {
                if (-not (Test-Path $destPath)) {
                    New-Item -Path $destPath -ItemType Directory -Force | Out-Null
                }
                
                $regFileName = $regPath -replace ':', '' -replace '\\', '_'
                $destFile = Join-Path $destPath "$regFileName.reg"
                
                # Exporta a chave do registro
                $regExport = "reg export `"$regPath`" `"$destFile`" /y"
                Invoke-Expression $regExport | Out-Null
                
                if (Test-Path $destFile) {
                    Write-Log "Registro exportado: $regPath ✓" -Level SUCCESS
                    $exportedAny = $true
                }
            }
            catch {
                Write-Log "Erro ao exportar $regPath : $_" -Level WARNING
            }
        }
    }
    
    if (-not $exportedAny) {
        Write-Log "Nenhuma configuração de registro exportada" -Level INFO
    }
    
    return $exportedAny
}

function Create-ExportManifest {
    <#
    .SYNOPSIS
        Cria um manifesto com informações sobre os itens exportados.
    #>
    Write-Log "Criando manifesto de exportação..." -Level INFO
    
    $manifest = @{
        ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        UserName = $env:USERNAME
        ComputerName = $env:COMPUTERNAME
        WordVersion = Get-WordVersion
        TotalItems = $script:ExportedItems.Count
        Items = $script:ExportedItems
    }
    
    $manifestPath = Join-Path $ExportPath "MANIFEST.json"
    $manifest | ConvertTo-Json -Depth 10 | Out-File -FilePath $manifestPath -Encoding UTF8
    
    Write-Log "Manifesto criado: $manifestPath ✓" -Level SUCCESS
    
    # Cria também um README
    $readmePath = Join-Path $ExportPath "README.txt"
    $readmeContent = @"
================================================================================
CHAINSAW - Personalizações Exportadas do Word
================================================================================

Data de Exportação: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")
Usuário: $env:USERNAME
Computador: $env:COMPUTERNAME
Versão do Word: $(Get-WordVersion)

Total de itens exportados: $($script:ExportedItems.Count)

CONTEÚDO:
---------

Templates/
    - Normal.dotm: Template global do Word com macros e personalizações

RibbonCustomization/
    - Personalizações da Faixa de Opções (abas customizadas)

OfficeCustomUI/
    - Arquivos de configuração da interface do Office

Templates/LiveContent/16/
    Managed/Document Themes/
        - Temas de documentos gerenciados pelo sistema
    
    User/Document Themes/
        - Temas personalizados pelo usuário
    
    Managed/Word Document Building Blocks/
        - Blocos de construção gerenciados
    
    User/Word Document Building Blocks/
        - Blocos de construção e partes rápidas do usuário

Registry/ (se incluído)
    - Configurações do registro do Word

COMO IMPORTAR:
--------------

Para importar estas configurações em outra máquina:

1. Copie toda esta pasta para a máquina de destino

2. Execute o script de importação:
   .\import-config.ps1

Ou use o instalador principal:
   .\install.cmd

================================================================================
"@
    
    $readmeContent | Out-File -FilePath $readmePath -Encoding UTF8
    Write-Log "README criado: $readmePath ✓" -Level SUCCESS
}

# =============================================================================
# FUNÇÃO PRINCIPAL
# =============================================================================

function Export-WordCustomizations {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║        CHAINSAW - Exportação de Personalizações do Word       ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    
    # Inicializa log
    Initialize-LogFile | Out-Null
    Write-Log "=== INÍCIO DA EXPORTAÇÃO ===" -Level INFO
    
    # Verifica e fecha Word se necessário
    if (-not (Confirm-CloseWord)) {
        Write-Log "Exportação cancelada - Word não foi fechado" -Level WARNING
        return
    }
    
    # Cria pasta de exportação
    if (-not (Test-Path $ExportPath)) {
        New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
        Write-Log "Pasta de exportação criada: $ExportPath" -Level INFO
    }
    
    $startTime = Get-Date
    
    try {
        Write-Host ""
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host "  Exportando Personalizações" -ForegroundColor White
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host ""
        
        # 1. Normal.dotm
        Export-NormalTemplate | Out-Null
        
        # 2. Building Blocks
        Export-BuildingBlocks | Out-Null
        
        # 3. Temas
        Export-DocumentThemes | Out-Null
        
        # 4. Ribbon
        Export-RibbonCustomization | Out-Null
        
        # 5. Custom UI
        Export-OfficeCustomUI | Out-Null
        
        # 6. QAT
        Export-QuickAccessToolbar | Out-Null
        
        # 7. Registro (opcional)
        Export-RegistrySettings | Out-Null
        
        # 8. Manifesto
        Create-ExportManifest
        
        $endTime = Get-Date
        $duration = $endTime - $startTime
        
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "║              EXPORTAÇÃO CONCLUÍDA COM SUCESSO!                 ║" -ForegroundColor Green
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host "📊 Resumo:" -ForegroundColor Cyan
        Write-Host "   • Itens exportados: $($script:ExportedItems.Count)" -ForegroundColor White
        Write-Host "   • Caminho: $ExportPath" -ForegroundColor Gray
        Write-Host "   • Tempo decorrido: $($duration.ToString('mm\:ss'))" -ForegroundColor Gray
        Write-Host ""
        Write-Host "📝 Log: $script:LogFile" -ForegroundColor Gray
        Write-Host ""
        
        Write-Log "=== EXPORTAÇÃO CONCLUÍDA COM SUCESSO ===" -Level SUCCESS
        Write-Log "Total de itens: $($script:ExportedItems.Count)" -Level INFO
    }
    catch {
        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Red
        Write-Host "║                  ERRO NA EXPORTAÇÃO!                           ║" -ForegroundColor Red
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Red
        Write-Host ""
        Write-Host "❌ Erro: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        
        Write-Log "=== EXPORTAÇÃO FALHOU ===" -Level ERROR
        Write-Log "Erro: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

# =============================================================================
# EXECUÇÃO
# =============================================================================

Export-WordCustomizations
