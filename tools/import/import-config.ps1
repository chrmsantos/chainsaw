# =============================================================================
# CHAINSAW - Importador de Personalizacoes do Word
# =============================================================================
# Versao: 1.0.1
# Licenca: GNU GPLv3
# Autor: Christian Martin dos Santos
# =============================================================================

<#
.SYNOPSIS
    Importa personalizacoes do Word exportadas pelo CHAINSAW.

.DESCRIPTION
    Restaura dois artefatos principais a partir de uma pasta exportada:
    1. Projeto VBA completo (todos os componentes) para o Normal.dotm do usuario.
    2. Personalizacoes do Ribbon do Word (Word.officeUI).

.PARAMETER ImportPath
    Pasta contendo a exportacao (padrao: .\exported-config).

.PARAMETER ForceCloseWord
    Fecha o Word automaticamente antes de importar.

.EXAMPLE
    .\import-config.ps1
    Importa a partir da pasta padrao .\exported-config.

.EXAMPLE
    .\import-config.ps1 -ImportPath "D:\\backup\\chainsaw_export" -ForceCloseWord
    Importa de um caminho personalizado, fechando o Word automaticamente.
#>

[CmdletBinding()]
param(
    [Parameter()] [string]$ImportPath = '.\exported-config',
    [Parameter()] [switch]$ForceCloseWord
)

$ErrorActionPreference = 'Stop'

# =============================================================================
# CONFIG
# =============================================================================
$script:LogFile = $null
$script:LogItems = @()
$ColorSuccess = 'Green'
$ColorWarning = 'Yellow'
$ColorError = 'Red'
$ColorInfo = 'Cyan'

$AppDataPath = $env:APPDATA
$LocalAppDataPath = $env:LOCALAPPDATA
$TemplatesPath = Join-Path $AppDataPath 'Microsoft\Templates'

# =============================================================================
# LOG
# =============================================================================
function Initialize-LogFile {
    try {
        $logDir = Join-Path $ImportPath 'logs'
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }

        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $script:LogFile = Join-Path $logDir "import_${timestamp}.log"

        $header = @"
================================================================================
CHAINSAW - Importacao de Personalizacoes do Word
================================================================================
Data/Hora: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")
Usuario: $env:USERNAME
Computador: $env:COMPUTERNAME
Sistema: $([Environment]::OSVersion.VersionString)
PowerShell: $($PSVersionTable.PSVersion)
Caminho de Importacao: $ImportPath
================================================================================

"@
        Add-Content -Path $script:LogFile -Value $header
        Invoke-LogRetention -Directory $logDir -Pattern 'import_*.log' -KeepLatest 5
        return $true
    }
    catch {
        Write-Warning "Nao foi possivel criar arquivo de log: $_"
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
        [Parameter(Mandatory)] [string]$Message,
        [Parameter()] [ValidateSet('INFO','SUCCESS','WARNING','ERROR')] [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"

    if ($script:LogFile) {
        try { Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue } catch {}
    }

    switch ($Level) {
        'SUCCESS' { Write-Host "[OK] $Message" -ForegroundColor $ColorSuccess }
        'WARNING' { Write-Host "[AVISO] $Message" -ForegroundColor $ColorWarning }
        'ERROR' { Write-Host "[ERRO] $Message" -ForegroundColor $ColorError }
        default { Write-Host "[INFO] $Message" -ForegroundColor $ColorInfo }
    }
}

# =============================================================================
# WORD / VBA HELPERS
# =============================================================================
function Get-WordRegistryVersion {
    $OfficeVersions = @('16.0','15.0','14.0')
    foreach ($version in $OfficeVersions) {
        $regPath = "HKCU:\\Software\\Microsoft\\Office\\${version}\\Word"
        if (Test-Path $regPath) { return $version }
    }
    return $null
}

function Test-VBAAccessEnabled {
    $wordVersion = Get-WordRegistryVersion
    if ($null -eq $wordVersion) { return $false }

    $regPath = "HKCU:\\Software\\Microsoft\\Office\\${wordVersion}\\Word\\Security"
    if (Test-Path $regPath) {
        $accessVBOM = Get-ItemProperty -Path $regPath -Name 'AccessVBOM' -ErrorAction SilentlyContinue
        return ($null -ne $accessVBOM -and $accessVBOM.AccessVBOM -eq 1)
    }
    return $false
}

function Enable-VBAAccess {
    param([switch]$Silent)

    $wordVersion = Get-WordRegistryVersion
    if ($null -eq $wordVersion) {
        if (-not $Silent) { Write-Log 'Word nao encontrado no registro' -Level WARNING }
        return $false
    }

    $regPath = "HKCU:\\Software\\Microsoft\\Office\\${wordVersion}\\Word\\Security"

    try {
        if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
        Set-ItemProperty -Path $regPath -Name 'AccessVBOM' -Value 1 -Type DWord -Force

        $currentValue = Get-ItemProperty -Path $regPath -Name 'AccessVBOM' -ErrorAction SilentlyContinue
        if ($currentValue.AccessVBOM -eq 1) {
            if (-not $Silent) { Write-Log 'Acesso ao VBA habilitado' -Level SUCCESS }
            return $true
        }

        if (-not $Silent) { Write-Log 'Falha ao verificar habilitacao do VBA' -Level ERROR }
        return $false
    }
    catch {
        if (-not $Silent) { Write-Log "Erro ao habilitar acesso ao VBA: $_" -Level ERROR }
        return $false
    }
}

function Test-WordRunning {
    $wordProcesses = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue
    return ($null -ne $wordProcesses -and $wordProcesses.Count -gt 0)
}

function Stop-WordProcesses {
    param([switch]$Force)

    try {
        $wordProcesses = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue
        if ($null -eq $wordProcesses -or $wordProcesses.Count -eq 0) { return $true }

        foreach ($p in $wordProcesses) {
            try {
                if ($Force) {
                    $p.Kill(); $p.WaitForExit(5000)
                }
                else {
                    $p.CloseMainWindow() | Out-Null
                    Start-Sleep -Milliseconds 500
                    if (-not $p.HasExited) { $p.Kill(); $p.WaitForExit(5000) }
                }
            }
            catch { Write-Log "Erro ao encerrar WINWORD (PID $($p.Id)): $_" -Level WARNING }
        }

        Start-Sleep -Milliseconds 500
        $remaining = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue
        return ($null -eq $remaining -or $remaining.Count -eq 0)
    }
    catch {
        Write-Log "Erro ao encerrar processos do Word: $_" -Level ERROR
        return $false
    }
}

function Confirm-CloseWord {
    if (-not (Test-WordRunning)) { return $true }

    if ($ForceCloseWord) {
        Write-Log 'Fechamento automatico do Word solicitado' -Level INFO
        return (Stop-WordProcesses -Force)
    }

    Write-Host ''
    Write-Host '[AVISO] Feche o Word antes de prosseguir.' -ForegroundColor Yellow
    $answer = Read-Host 'Pressione S para fechar automaticamente ou N para cancelar [S/n]'
    if ([string]::IsNullOrWhiteSpace($answer) -or $answer.Trim().ToLowerInvariant() -in @('s','sim','y','yes')) {
        return (Stop-WordProcesses -Force)
    }

    Write-Log 'Importacao cancelada - Word permanecia aberto' -Level WARNING
    return $false
}

# =============================================================================
# IMPORT HELPERS
# =============================================================================
function Backup-NormalTemplate {
    $normalPath = Join-Path $TemplatesPath 'Normal.dotm'
    if (-not (Test-Path $normalPath)) {
        Write-Log 'Normal.dotm nao encontrado para backup' -Level WARNING
        return $false
    }

    $backupDir = Join-Path $ImportPath 'backups'
    if (-not (Test-Path $backupDir)) { New-Item -Path $backupDir -ItemType Directory -Force | Out-Null }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupFile = Join-Path $backupDir "Normal_backup_${timestamp}.dotm"

    Copy-Item -Path $normalPath -Destination $backupFile -Force
    Write-Log "Backup criado: $backupFile" -Level SUCCESS
    return $true
}

function Import-VbaProject {
    Write-Log 'Importando projeto VBA...' -Level INFO

    $projectPath = Join-Path $ImportPath 'VBAProject'
    if (-not (Test-Path $projectPath)) {
        Write-Log 'Pasta VBAProject nao encontrada na importacao' -Level ERROR
        return $false
    }

    $files = Get-ChildItem -Path $projectPath -File | Sort-Object Name
    if ($files.Count -eq 0) {
        Write-Log 'Nenhum componente VBA encontrado para importar' -Level WARNING
        return $false
    }

    $normalPath = Join-Path $TemplatesPath 'Normal.dotm'
    if (-not (Test-Path $normalPath)) {
        Write-Log 'Normal.dotm nao encontrado no perfil do usuario' -Level ERROR
        return $false
    }

    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0

        $template = $word.Documents.Open($normalPath, $false, $false)
        if ($null -eq $template.VBProject) {
            Write-Log 'VBProject nao encontrado no Normal.dotm' -Level ERROR
            $template.Close($false); $word.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            return $false
        }

        $vbProject = $template.VBProject

        # Remove componentes existentes (exceto documentos)
        foreach ($component in @($vbProject.VBComponents)) {
            if ($component.Type -eq 100) { continue } # Document modules nao podem ser removidos
            try { $vbProject.VBComponents.Remove($component) } catch { Write-Log "Falha ao remover componente existente $($component.Name): $_" -Level WARNING }
        }

        $imported = 0
        foreach ($file in $files) {
            try {
                $vbProject.VBComponents.Import($file.FullName) | Out-Null
                $imported++
                $script:LogItems += [PSCustomObject]@{
                    Type = 'VBA Component'
                    Source = $file.FullName
                    Destination = 'Normal.dotm'
                    Size = $file.Length
                }
            }
            catch {
                Write-Log "Falha ao importar componente ${file.Name}: $_" -Level WARNING
            }
        }

        $template.Save()
        $template.Close()
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        if ($imported -gt 0) {
            Write-Log "Componentes VBA importados: $imported" -Level SUCCESS
            return $true
        }

        Write-Log 'Nenhum componente VBA foi importado' -Level WARNING
        return $false
    }
    catch {
        Write-Log "Erro ao importar projeto VBA: $_" -Level ERROR
        try { if ($word) { $word.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null } } catch {}
        return $false
    }
}

function Import-RibbonCustomization {
    Write-Log 'Importando personalizacao do Ribbon...' -Level INFO

    $ribbonDir = Join-Path $ImportPath 'RibbonCustomization'
    $officeDir = Join-Path $ImportPath 'OfficeCustomUI'

    $sourceRibbon = $null
    if (Test-Path (Join-Path $ribbonDir 'Word.exportedUI')) {
        $sourceRibbon = Join-Path $ribbonDir 'Word.exportedUI'
    }
    elseif (Test-Path (Join-Path $ribbonDir 'Word.officeUI')) {
        $sourceRibbon = Join-Path $ribbonDir 'Word.officeUI'
    }
    elseif (Test-Path $officeDir) {
        $candidate = Get-ChildItem -Path $officeDir -Filter 'Word*.exportedUI' -File | Select-Object -First 1
        if (-not $candidate) {
            $candidate = Get-ChildItem -Path $officeDir -Filter 'Word*.officeUI' -File | Select-Object -First 1
        }
        if ($candidate) { $sourceRibbon = $candidate.FullName }
    }

    if (-not $sourceRibbon) {
        Write-Log 'Nenhum arquivo Word.exportedUI ou Word.officeUI encontrado para importar' -Level WARNING
        return $false
    }

    $destDir = Join-Path $LocalAppDataPath 'Microsoft\Office'
    if (-not (Test-Path $destDir)) { New-Item -Path $destDir -ItemType Directory -Force | Out-Null }

    $destFile = Join-Path $destDir 'Word.officeUI'

    if (Test-Path $destFile) {
        $backupDir = Join-Path $ImportPath 'backups'
        if (-not (Test-Path $backupDir)) { New-Item -Path $backupDir -ItemType Directory -Force | Out-Null }
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $backupFile = Join-Path $backupDir "Word.officeUI.bak_${timestamp}.bak"
        Copy-Item -Path $destFile -Destination $backupFile -Force
        Write-Log "Backup do Ribbon criado: $backupFile" -Level INFO
    }

    Copy-Item -Path $sourceRibbon -Destination $destFile -Force
    Write-Log "Ribbon importado de $sourceRibbon" -Level SUCCESS
    return $true
}

function Compile-VbaProject {
    Write-Log 'Compilando projeto VBA apos importacao...' -Level INFO

    $normalPath = Join-Path $TemplatesPath 'Normal.dotm'
    if (-not (Test-Path $normalPath)) { return $false }

    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0

        $template = $word.Documents.Open($normalPath, $false, $false)
        if ($template.VBProject) {
            foreach ($component in $template.VBProject.VBComponents) { $null = $component.CodeModule.CountOfLines }
            Write-Log 'Projeto VBA compilado com sucesso' -Level SUCCESS
        }

        $template.Close($true)
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        return $true
    }
    catch {
        Write-Log "Falha ao compilar projeto VBA: $_" -Level WARNING
        try { if ($word) { $word.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null } } catch {}
        return $false
    }
}

# =============================================================================
# EXECUCAO PRINCIPAL
# =============================================================================

try {
    Write-Host ''
    Write-Host '╔════════════════════════════════════════════════════════════════╗' -ForegroundColor Cyan
    Write-Host '║        CHAINSAW - Importacao de Personalizacoes do Word       ║' -ForegroundColor Cyan
    Write-Host '╚════════════════════════════════════════════════════════════════╝' -ForegroundColor Cyan
    Write-Host ''

    $ImportPath = Resolve-Path -Path $ImportPath -ErrorAction Stop
    Initialize-LogFile | Out-Null
    Write-Log '=== INICIO DA IMPORTACAO ===' -Level INFO

    if (-not (Test-VBAAccessEnabled)) {
        Write-Log 'Acesso ao VBA nao habilitado - tentando habilitar' -Level WARNING
        Enable-VBAAccess | Out-Null
    }

    if (-not (Confirm-CloseWord)) { throw 'Word permaneceu aberto. Importe apos fechar o aplicativo.' }

    Backup-NormalTemplate | Out-Null

    $vbaOk = Import-VbaProject
    $ribbonOk = Import-RibbonCustomization

    Compile-VbaProject | Out-Null

    Write-Log "Importacao finalizada. VBA: $vbaOk, Ribbon: $ribbonOk" -Level INFO

    Write-Host ''
    Write-Host '╔════════════════════════════════════════════════════════════════╗' -ForegroundColor Green
    Write-Host '║              IMPORTACAO CONCLUIDA COM SUCESSO!                 ║' -ForegroundColor Green
    Write-Host '╚════════════════════════════════════════════════════════════════╝' -ForegroundColor Green
    Write-Host ''
    Write-Host "  • Projeto VBA: $vbaOk" -ForegroundColor White
    Write-Host "  • Ribbon: $ribbonOk" -ForegroundColor White
    Write-Host "  • Caminho: $ImportPath" -ForegroundColor Gray
    Write-Host "  • Log: $script:LogFile" -ForegroundColor Gray
    Write-Host ''
}
catch {
    Write-Host ''
    Write-Host '╔════════════════════════════════════════════════════════════════╗' -ForegroundColor Red
    Write-Host '║                  ERRO NA IMPORTACAO!                           ║' -ForegroundColor Red
    Write-Host '╚════════════════════════════════════════════════════════════════╝' -ForegroundColor Red
    Write-Host ''
    Write-Host "[ERRO] $_" -ForegroundColor Red
    Write-Log "Importacao falhou: $_" -Level ERROR
    exit 1
}
# Execucao finalizada sem pausa interativa
