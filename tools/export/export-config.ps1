# =============================================================================
# CHAINSAW - Script de Exportação de Personalizações do Word
# =============================================================================
# Versão: 1.1.1
# Licença: GNU GPLv3
# Autor: Christian Martin dos Santos
# =============================================================================

<#
.SYNOPSIS
    Exporta o módulo VBA e personalizações da interface do Word.

.DESCRIPTION
    Exporta apenas o que importa:
    1. Todo o projeto VBA (todos os módulos) a partir do Normal.dotm
    2. Personalizações do Ribbon do Word (.officeUI)

.PARAMETER ExportPath
    Caminho onde as personalizações serão exportadas.
    Padrão: .\exported-config

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
    [switch]$ForceCloseWord
)

$script:ForceCloseWord = $ForceCloseWord

# Remove argumento de maximização se presente
if ($ExportPath -eq "__MAXIMIZED__") {
    $ExportPath = ".\exported-config"
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
        Invoke-LogRetention -Directory $logDir -Pattern 'export_*.log' -KeepLatest 5
        return $true
    }
    catch {
        Write-Warning "Não foi possível criar arquivo de log: $_"
        return $false
    }
}

function Invoke-LogRetention {
    param(
        [Parameter(Mandatory)] [string]$Directory,
        [Parameter(Mandatory)] [string]$Pattern,
        [int]$KeepLatest = 5
    )

    try {
        if ($KeepLatest -lt 1) { return }
        if (-not (Test-Path $Directory)) { return }

        $logFiles = Get-ChildItem -Path $Directory -Filter $Pattern -File -ErrorAction Stop |
            Sort-Object LastWriteTime -Descending

        if ($logFiles.Count -le $KeepLatest) { return }

        $logFiles[$KeepLatest..($logFiles.Count - 1)] | ForEach-Object {
            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Verbose "Falha ao aplicar retencao de logs em ${Directory}: $_"
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
        "SUCCESS" { Write-Host "[OK] $Message" -ForegroundColor $ColorSuccess }
        "WARNING" { Write-Host "[AVISO] $Message" -ForegroundColor $ColorWarning }
        "ERROR" { Write-Host "[ERRO] $Message" -ForegroundColor $ColorError }
        default { Write-Host "[INFO] $Message" -ForegroundColor $ColorInfo }
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

function Get-WordRegistryVersion {
    <#
    .SYNOPSIS
        Detecta a versão do Word no registro (16.0, 15.0, 14.0).
    #>
    $OfficeVersions = @("16.0", "15.0", "14.0")
    foreach ($version in $OfficeVersions) {
        $regPath = "HKCU:\Software\Microsoft\Office\$version\Word"
        if (Test-Path $regPath) {
            return $version
        }
    }
    return $null
}

function Test-VBAAccessEnabled {
    <#
    .SYNOPSIS
        Verifica se o acesso programático ao VBA está habilitado.
    #>
    $wordVersion = Get-WordRegistryVersion
    if ($null -eq $wordVersion) {
        return $false
    }

    $regPath = "HKCU:\Software\Microsoft\Office\$wordVersion\Word\Security"
    if (Test-Path $regPath) {
        $accessVBOM = Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue
        return ($null -ne $accessVBOM -and $accessVBOM.AccessVBOM -eq 1)
    }
    return $false
}

function Enable-VBAAccess {
    <#
    .SYNOPSIS
        Habilita o acesso programático ao VBA.
    #>
    param(
        [switch]$Silent
    )

    $wordVersion = Get-WordRegistryVersion
    if ($null -eq $wordVersion) {
        if (-not $Silent) {
            Write-Log "Word não encontrado no registro" -Level WARNING
        }
        return $false
    }

    $regPath = "HKCU:\Software\Microsoft\Office\$wordVersion\Word\Security"

    try {
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }

        Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -Force

        $currentValue = Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue

        if ($currentValue.AccessVBOM -eq 1) {
            if (-not $Silent) {
                Write-Log "Acesso ao VBA habilitado com sucesso" -Level SUCCESS
            }
            return $true
        }
        else {
            if (-not $Silent) {
                Write-Log "Falha ao verificar habilitação do VBA" -Level ERROR
            }
            return $false
        }
    }
    catch {
        if (-not $Silent) {
            Write-Log "Erro ao habilitar acesso ao VBA: $_" -Level ERROR
        }
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

    if ($script:ForceCloseWord) {
        Write-Log "Fechamento automático do Word solicitado." -Level INFO
        if (Stop-WordProcesses -Force) {
            Write-Log "Word fechado automaticamente" -Level SUCCESS
            Start-Sleep -Seconds 2
            return $true
        }

        Write-Log "Não foi possível fechar o Word automaticamente" -Level ERROR
        return $false
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
    $response = Read-Host "Deseja FECHAR o Word e continuar a exportação? (S/N)"

    if ($response -notmatch '^[Ss]$') {
        Write-Host ""
        Write-Host "[OK] Exportação cancelada pelo usuário" -ForegroundColor Cyan
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
        Write-Host "[OK] Word fechado com sucesso" -ForegroundColor Green
        Write-Host ""
        # Aguarda um pouco para garantir que recursos foram liberados
        Start-Sleep -Seconds 2
        return $true
    }
    else {
        Write-Host "[ERRO] Não foi possível fechar o Word completamente" -ForegroundColor Red
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

function Test-VbaProjectCompilation {
    <#
    .SYNOPSIS
        Compila o módulo VBA antes da exportação para verificar erros.
    #>
    Write-Log "Compilando projeto VBA..." -Level INFO

    try {
        # Cria instância do Word
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0  # wdAlertsNone

        # Caminho do Normal.dotm
        $normalPath = Join-Path $TemplatesPath "Normal.dotm"

        if (-not (Test-Path $normalPath)) {
            Write-Log "Normal.dotm não encontrado - compilação ignorada" -Level WARNING
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            return $true
        }

        # Abre o Normal.dotm
        $template = $word.Documents.Open($normalPath, $false, $false)

        # Verifica se há projeto VBA
        if ($null -eq $template.VBProject) {
            Write-Log "Nenhum projeto VBA encontrado - compilação ignorada" -Level INFO
            $template.Close($false)
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            return $true
        }

        # Compila o projeto VBA
        try {
            $vbProject = $template.VBProject

            # Força compilação acessando os módulos
            foreach ($component in $vbProject.VBComponents) {
                $null = $component.CodeModule.CountOfLines
            }

            Write-Log "Projeto VBA compilado com sucesso [OK]" -Level SUCCESS
            $compilationSuccess = $true
        }
        catch {
            Write-Log "Erro ao compilar projeto VBA: $_" -Level ERROR
            Write-Log "ATENÇÃO: O projeto pode conter erros de compilação!" -Level WARNING
            $compilationSuccess = $false
        }

        # Fecha sem salvar
        $template.Close($false)
        $word.Quit()

        # Libera COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        return $compilationSuccess
    }
    catch {
        Write-Log "Erro ao verificar compilação: $_" -Level ERROR

        # Tenta limpar recursos
        try {
            if ($word) {
                $word.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            }
        }
        catch { }

        return $false
    }
}

function Export-VbaProject {
    <#
    .SYNOPSIS
        Exporta todo o projeto VBA (todos os componentes) do Normal.dotm.
    #>
    Write-Log "Exportando projeto VBA completo..." -Level INFO

    $normalPath = Join-Path $TemplatesPath "Normal.dotm"

    if (-not (Test-Path $normalPath)) {
        Write-Log "Normal.dotm não encontrado em: $normalPath" -Level ERROR
        return $false
    }

    $destPath = Join-Path $ExportPath "VBAProject"

    try {
        if (-not (Test-Path $destPath)) {
            New-Item -Path $destPath -ItemType Directory -Force | Out-Null
        }

        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0 # wdAlertsNone

        $template = $word.Documents.Open($normalPath, $false, $true) # ReadOnly

        if ($null -eq $template.VBProject) {
            Write-Log "Nenhum projeto VBA encontrado" -Level WARNING
            $template.Close($false)
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            return $false
        }

        $vbProject = $template.VBProject
        $exportedCount = 0

        foreach ($component in $vbProject.VBComponents) {
            $baseName = $component.Name
            $extension = switch ($component.Type) {
                1 { '.bas' }   # vbext_ct_StdModule
                2 { '.cls' }   # vbext_ct_ClassModule
                3 { '.frm' }   # vbext_ct_MSForm
                100 { '.cls' } # Document modules
                default { '.txt' }
            }

            $safeName = "{0:D2}_{1}{2}" -f $exportedCount, $baseName, $extension
            $exportFile = Join-Path $destPath $safeName

            try {
                $component.Export($exportFile)
                $script:ExportedItems += [PSCustomObject]@{
                    Type        = "VBA Component"
                    Source      = "Normal.dotm::${baseName}"
                    Destination = $exportFile
                    Size        = (Get-Item $exportFile).Length
                }
                $exportedCount++
            }
            catch {
                Write-Log "Falha ao exportar componente ${baseName}: $_" -Level WARNING
            }
        }

        $template.Close($false)
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        if ($exportedCount -gt 0) {
            Write-Log "Componentes VBA exportados: $exportedCount [OK]" -Level SUCCESS
            return $true
        }

        Write-Log "Nenhum componente VBA exportado" -Level WARNING
        return $false
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Erro ao exportar projeto VBA: $errorMsg" -Level ERROR

        if ($errorMsg -match "0x800AC35C" -or $errorMsg -match "programmatic access") {
            Write-Log "DIAGNÓSTICO: Erro 0x800AC35C - Acesso programático ao VBA bloqueado" -Level WARNING
            Write-Log "Possíveis causas:" -Level INFO
            Write-Log "  1. Word precisa ser reiniciado após habilitar AccessVBOM" -Level INFO
            Write-Log "  2. Configuração de segurança de macros muito restritiva" -Level INFO
            Write-Log "  3. Política de grupo impedindo acesso" -Level INFO
            Write-Log "ALTERNATIVA: Copiar Normal.dotm diretamente (contém o projeto VBA)" -Level INFO
        }

        try {
            if ($word) {
                $word.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            }
        }
        catch { }

        return $false
    }
}

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
            Type        = "Normal Template"
            Source      = $normalPath
            Destination = Join-Path $destPath "Normal.dotm"
            Size        = (Get-Item $normalPath).Length
        }

        Write-Log "Normal.dotm exportado com sucesso [OK]" -Level SUCCESS
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
                    Type        = "Building Block (User)"
                    Source      = $file.FullName
                    Destination = $destFile
                    Size        = $file.Length
                }
            }

            Write-Log "Building Blocks do usuário: $($files.Count) arquivos" -Level INFO
        }
        catch {
            Write-Log "Erro ao exportar Building Blocks do usuário: $_" -Level WARNING
        }
    }

    if ($exportedCount -gt 0) {
        Write-Log "Building Blocks exportados: $exportedCount arquivos [OK]" -Level SUCCESS
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
        Write-Log "Temas exportados: $exportedCount arquivos [OK]" -Level SUCCESS
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
                    Type        = "Ribbon Customization"
                    Source      = $uiPath
                    Destination = $destFile
                    Size        = (Get-Item $uiPath).Length
                }

                Write-Log "Personalização do Ribbon exportada: $fileName [OK]" -Level SUCCESS
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
        # Procura apenas personalizações do Word (.officeUI)
        $customFiles = Get-ChildItem -Path $customUIPath -Filter "*.officeUI" -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^Word.*\.officeUI$' }

        if ($customFiles.Count -gt 0) {
            if (-not (Test-Path $destPath)) {
                New-Item -Path $destPath -ItemType Directory -Force | Out-Null
            }

            foreach ($file in $customFiles) {
                $destFile = Join-Path $destPath $file.Name
                Copy-Item -Path $file.FullName -Destination $destFile -Force

                $script:ExportedItems += [PSCustomObject]@{
                    Type        = "Office Custom UI"
                    Source      = $file.FullName
                    Destination = $destFile
                    Size        = $file.Length
                }
            }

            Write-Log "Personalizações UI do Word exportadas: $($customFiles.Count) arquivos [OK]" -Level SUCCESS
            return $true
        }
        else {
            Write-Log "Nenhum arquivo de personalização UI do Word encontrado" -Level INFO
            return $false
        }
    }
    catch {
        Write-Log "Erro ao exportar personalizações UI: $_" -Level WARNING
        return $false
    }
}

function New-ExportManifest {
    <#
    .SYNOPSIS
        Cria um manifesto com informações sobre os itens exportados.
    #>
    Write-Log "Criando manifesto de exportação..." -Level INFO

    $manifest = @{
        ExportDate   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        UserName     = $env:USERNAME
        ComputerName = $env:COMPUTERNAME
        WordVersion  = Get-WordVersion
        TotalItems   = $script:ExportedItems.Count
        Items        = $script:ExportedItems
    }

    $manifestPath = Join-Path $ExportPath "MANIFEST.json"
    $manifest | ConvertTo-Json -Depth 10 | Out-File -FilePath $manifestPath -Encoding UTF8

    Write-Log "Manifesto criado: $manifestPath [OK]" -Level SUCCESS

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

VBAProject/
    - Componentes exportados do projeto VBA (módulos, classes, forms)

Templates/
    - Normal.dotm: cópia de referência contendo o projeto VBA compilado

RibbonCustomization/
    - Personalizações da Faixa de Opções do Word (Word.officeUI)

OfficeCustomUI/
    - Arquivos .officeUI do Word

COMO UTILIZAR:
--------------

- VBA: importe os arquivos da pasta VBAProject pelo editor do VBA ou substitua o Normal.dotm pela cópia exportada.
- Ribbon: copie Word.officeUI para %LOCALAPPDATA%\Microsoft\Office\.

================================================================================
"@

    $readmeContent | Out-File -FilePath $readmePath -Encoding UTF8
    Write-Log "README criado: $readmePath [OK]" -Level SUCCESS
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

    # Verifica e habilita acesso ao VBA se necessário
    $vbaAccessWasEnabled = Test-VBAAccessEnabled
    if (-not $vbaAccessWasEnabled) {
        Write-Host ""
        Write-Host "[AVISO] Acesso programático ao VBA não está habilitado" -ForegroundColor Yellow
        Write-Host "  Esta configuração é necessária para exportar módulos VBA." -ForegroundColor Gray
        Write-Host ""
        Write-Log "Acesso ao VBA não habilitado - habilitando automaticamente" -Level WARNING

        if (Enable-VBAAccess -Silent) {
            Write-Host "[OK] Acesso ao VBA habilitado automaticamente" -ForegroundColor Green
            Write-Log "Acesso ao VBA habilitado automaticamente" -Level SUCCESS
            Write-Host ""
            Write-Host "[INFO] IMPORTANTE: Se o Word foi aberto recentemente, a configuração" -ForegroundColor Cyan
            Write-Host "       só terá efeito após reiniciar o Word completamente." -ForegroundColor Cyan
            Write-Host ""
        }
        else {
            Write-Host "[ERRO] Não foi possível habilitar acesso ao VBA automaticamente" -ForegroundColor Red
            Write-Host "  Execute: .\enable-vba-access.ps1" -ForegroundColor Gray
            Write-Log "Falha ao habilitar acesso ao VBA - exportação pode falhar" -Level ERROR
        }
        Write-Host ""
    }
    else {
        Write-Log "Acesso ao VBA já está habilitado" -Level INFO
    }

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

        Write-Host "Compilando projeto VBA..." -ForegroundColor Cyan
        $compiled = Test-VbaProjectCompilation
        if (-not $compiled) {
            Write-Host "[AVISO] Projeto VBA não pôde ser compilado. Verifique o log." -ForegroundColor Yellow
        }
        Write-Host ""

        Write-Host "Exportando projeto VBA..." -ForegroundColor Cyan
        if (-not (Export-VbaProject)) {
            Write-Host "[AVISO] Nenhum componente VBA exportado" -ForegroundColor Yellow
        }
        Write-Host ""

        Write-Host "Exportando Normal.dotm (cópia de referência)..." -ForegroundColor Cyan
        if (Export-NormalTemplate) {
            Write-Host "[OK] Normal.dotm exportado com sucesso" -ForegroundColor Green
        }
        else {
            Write-Host "[AVISO] Falha ao exportar Normal.dotm" -ForegroundColor Yellow
        }
        Write-Host ""

        Export-RibbonCustomization | Out-Null
        Export-OfficeCustomUI | Out-Null

        New-ExportManifest

        $endTime = Get-Date
        $duration = $endTime - $startTime

        Write-Host ""
        Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "║              EXPORTAÇÃO CONCLUÍDA COM SUCESSO!                 ║" -ForegroundColor Green
        Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host " Resumo:" -ForegroundColor Cyan
        Write-Host "   • Itens exportados: $($script:ExportedItems.Count)" -ForegroundColor White
        Write-Host "   • Caminho: $ExportPath" -ForegroundColor Gray
        Write-Host "   • Tempo decorrido: $($duration.ToString('mm\:ss'))" -ForegroundColor Gray
        Write-Host ""
        Write-Host " Log: $script:LogFile" -ForegroundColor Gray
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
        Write-Host "[ERRO] Erro: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""

        Write-Log "=== EXPORTAÇÃO FALHOU ===" -Level ERROR
        Write-Log "Erro: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

# =============================================================================
# EXECUÇÃO
# =============================================================================

try {
    Export-WordCustomizations
}
finally {
    # Pausa ao final da execução
    Write-Host ""
    # Execucao finalizada sem pausa interativa
}
