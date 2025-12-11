# =============================================================================
# CHAINSAW - Restauração de Personalizacoes do Word
# =============================================================================
# Versao: 1.0.0
# Licenca: GNU GPLv3
# =============================================================================

<#
.SYNOPSIS
    Restaura backups gerados pelo import-config.ps1.
.DESCRIPTION
    Restaura o Normal.dotm e o Word.officeUI a partir de backups em uma pasta.
.PARAMETER BackupPath
    Caminho onde os backups foram salvos (padrao: .\exported-config\backups).
.PARAMETER ForceCloseWord
    Fecha o Word automaticamente antes da restauracao.
#>

[CmdletBinding()]
param(
    [Parameter()] [string]$BackupPath = '.\exported-config\backups',
    [Parameter()] [switch]$ForceCloseWord
)

$ErrorActionPreference = 'Stop'

$AppDataPath = $env:APPDATA
$LocalAppDataPath = $env:LOCALAPPDATA
$TemplatesPath = Join-Path $AppDataPath 'Microsoft\Templates'

$ColorSuccess = 'Green'
$ColorWarning = 'Yellow'
$ColorError = 'Red'
$ColorInfo = 'Cyan'
$script:LogFile = $null

function Initialize-LogFile {
    try {
        $logDir = Join-Path $BackupPath 'logs'
        if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $script:LogFile = Join-Path $logDir "restore_${ts}.log"
        $header = @"
================================================================================
CHAINSAW - Restauracao de Personalizacoes do Word
================================================================================
Data/Hora: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")
Usuario: $env:USERNAME
Computador: $env:COMPUTERNAME
Backup: $BackupPath
================================================================================

"@
        Add-Content -Path $script:LogFile -Value $header
        return $true
    }
    catch {
        Write-Warning "Nao foi possivel criar log: $_"
        return $false
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)] [string]$Message,
        [Parameter()] [ValidateSet('INFO','SUCCESS','WARNING','ERROR')] [string]$Level = 'INFO'
    )
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts] [$Level] $Message"
    if ($script:LogFile) { try { Add-Content -Path $script:LogFile -Value $entry -ErrorAction SilentlyContinue } catch {} }
    switch ($Level) {
        'SUCCESS' { Write-Host "[OK] $Message" -ForegroundColor $ColorSuccess }
        'WARNING' { Write-Host "[AVISO] $Message" -ForegroundColor $ColorWarning }
        'ERROR' { Write-Host "[ERRO] $Message" -ForegroundColor $ColorError }
        default { Write-Host "[INFO] $Message" -ForegroundColor $ColorInfo }
    }
}

function Test-WordRunning {
    $procs = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue
    return ($procs -and $procs.Count -gt 0)
}

function Stop-WordProcesses {
    param([switch]$Force)
    try {
        $procs = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue
        if (-not $procs) { return $true }
        foreach ($p in $procs) {
            try {
                if ($Force) { $p.Kill(); $p.WaitForExit(5000) }
                else {
                    $p.CloseMainWindow() | Out-Null
                    Start-Sleep -Milliseconds 500
                    if (-not $p.HasExited) { $p.Kill(); $p.WaitForExit(5000) }
                }
            }
            catch { Write-Log "Falha ao encerrar WINWORD PID $($p.Id): $_" -Level WARNING }
        }
        Start-Sleep -Milliseconds 500
        $remain = Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue
        return (-not $remain)
    }
    catch {
        Write-Log "Erro ao encerrar Word: $_" -Level ERROR
        return $false
    }
}

function Confirm-CloseWord {
    if (-not (Test-WordRunning)) { return $true }
    if ($ForceCloseWord) { return (Stop-WordProcesses -Force) }
    Write-Host '[AVISO] Feche o Word antes de restaurar.' -ForegroundColor Yellow
    $answer = Read-Host 'Pressione S para fechar automaticamente ou N para cancelar [S/n]'
    if ([string]::IsNullOrWhiteSpace($answer) -or $answer.Trim().ToLowerInvariant() -in @('s','sim','y','yes')) {
        return (Stop-WordProcesses -Force)
    }
    Write-Log 'Restauracao cancelada - Word permaneceu aberto' -Level WARNING
    return $false
}

function Restore-NormalTemplate {
    $normalBackup = Get-ChildItem -Path $BackupPath -Filter 'Normal_backup_*.dotm' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $normalBackup) {
        Write-Log 'Nenhum backup de Normal.dotm encontrado' -Level WARNING
        return $false
    }

    $destPath = Join-Path $TemplatesPath 'Normal.dotm'
    Copy-Item -Path $normalBackup.FullName -Destination $destPath -Force
    Write-Log "Normal.dotm restaurado a partir de $($normalBackup.Name)" -Level SUCCESS
    return $true
}

function Restore-Ribbon {
    $ribbonBackup = Get-ChildItem -Path $BackupPath -Filter 'Word.officeUI.bak_*' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $ribbonBackup) {
        Write-Log 'Nenhum backup de Ribbon encontrado' -Level WARNING
        return $false
    }

    $destDir = Join-Path $LocalAppDataPath 'Microsoft\Office'
    if (-not (Test-Path $destDir)) { New-Item -Path $destDir -ItemType Directory -Force | Out-Null }
    $destFile = Join-Path $destDir 'Word.officeUI'
    Copy-Item -Path $ribbonBackup.FullName -Destination $destFile -Force
    Write-Log "Ribbon restaurado a partir de $($ribbonBackup.Name)" -Level SUCCESS
    return $true
}

try {
    $BackupPath = Resolve-Path -Path $BackupPath -ErrorAction Stop
    Initialize-LogFile | Out-Null

    if (-not (Confirm-CloseWord)) { throw 'Word permaneceu aberto. Restaure apos fechar o aplicativo.' }

    $normalOk = Restore-NormalTemplate
    $ribbonOk = Restore-Ribbon

    Write-Host ''
    Write-Host '╔════════════════════════════════════════════════════════════════╗' -ForegroundColor Green
    Write-Host '║              RESTAURACAO CONCLUIDA                             ║' -ForegroundColor Green
    Write-Host '╚════════════════════════════════════════════════════════════════╝' -ForegroundColor Green
    Write-Host ''
    Write-Host "  • Normal.dotm restaurado: $normalOk" -ForegroundColor White
    Write-Host "  • Ribbon restaurado: $ribbonOk" -ForegroundColor White
    Write-Host "  • Backup: $BackupPath" -ForegroundColor Gray
    Write-Host "  • Log: $script:LogFile" -ForegroundColor Gray
    Write-Host ''
}
catch {
    Write-Host ''
    Write-Host '╔════════════════════════════════════════════════════════════════╗' -ForegroundColor Red
    Write-Host '║                  ERRO NA RESTAURACAO!                          ║' -ForegroundColor Red
    Write-Host '╚════════════════════════════════════════════════════════════════╝' -ForegroundColor Red
    Write-Host ''
    Write-Host "[ERRO] $_" -ForegroundColor Red
    Write-Log "Restauracao falhou: $_" -Level ERROR
    exit 1
}
