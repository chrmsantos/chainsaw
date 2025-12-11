# =============================================================================
# CHAINSAW - Wrapper para restore-config.ps1
# =============================================================================

[CmdletBinding()]
param(
    [Parameter()] [string]$BackupPath = '.\exported-config\backups',
    [Parameter()] [switch]$ForceCloseWord
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent (Get-Item -LiteralPath $MyInvocation.MyCommand.Path).FullName
}
$toolsRoot = Split-Path -Parent $scriptRoot
$repoRoot = if ($toolsRoot) { Split-Path -Parent $toolsRoot } else { $scriptRoot }
$restoreScript = Join-Path $scriptRoot 'restore-config.ps1'
$resolvedBackupPath = if ([IO.Path]::IsPathRooted($BackupPath)) { $BackupPath } else { Join-Path $repoRoot $BackupPath }

function Show-Header {
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host '  CHAINSAW - Restaurador de Configuracoes' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host ''
}

function Start-Restore {
    param(
        [Parameter(Mandatory)] [string]$BackupPath,
        [Parameter()] [bool]$ForceCloseWord
    )

    $args = @(
        '-ExecutionPolicy','Bypass',
        '-NoProfile','-NoLogo',
        '-File',"`"$restoreScript`"",
        '-BackupPath',"`"$BackupPath`""
    )

    if ($ForceCloseWord) { $args += '-ForceCloseWord' }

    $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $args -Wait -PassThru
    return $proc.ExitCode
}

try {
    if (-not (Test-Path $restoreScript)) {
        Write-Host '[ERRO] restore-config.ps1 nao encontrado.' -ForegroundColor Red
        exit 1
    }

    Show-Header

    Write-Host 'Executando restaurador...' -ForegroundColor Cyan
    $exitCode = Start-Restore -BackupPath $resolvedBackupPath -ForceCloseWord:$ForceCloseWord.IsPresent

    if ($exitCode -eq 0) {
        Write-Host ''
        Write-Host 'Restauracao concluida com sucesso.' -ForegroundColor Green
        exit 0
    }

    Write-Host ''
    Write-Host 'Um erro foi identificado durante a restauracao.' -ForegroundColor Red
    exit $exitCode
}
catch {
    Write-Host ''
    Write-Host 'O restaurador encontrou uma falha inesperada.' -ForegroundColor Red
    exit 1
}
