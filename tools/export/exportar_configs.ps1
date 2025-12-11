# =============================================================================
# CHAINSAW - Wrapper para export-config.ps1
# =============================================================================
# Exporta personalizacoes do Word para uma pasta de destino (padrao: exported-config na raiz do repo).
# =============================================================================

[CmdletBinding()]
param(
    [Parameter()] [string]$ExportPath = '.\exported-config',
    [Parameter()] [switch]$ForceCloseWord
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent (Get-Item -LiteralPath $MyInvocation.MyCommand.Path).FullName
}
$toolsRoot = Split-Path -Parent $scriptRoot
$repoRoot = if ($toolsRoot) { Split-Path -Parent $toolsRoot } else { $scriptRoot }
$exportScript = Join-Path $scriptRoot 'export-config.ps1'
$resolvedExportPath = if ([IO.Path]::IsPathRooted($ExportPath)) {
    $ExportPath
} else {
    Join-Path $repoRoot $ExportPath
}

function Show-Header {
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host '  CHAINSAW - Exportador de Configuracoes' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host ''
}

function Get-LatestLogPath {
    param(
        [Parameter()] [string]$DestinationPath
    )

    $logDirectory = Join-Path $DestinationPath 'logs'
    if (-not (Test-Path $logDirectory)) {
        return $null
    }

    $latestLog = Get-ChildItem -Path $logDirectory -Filter 'export_*.log' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $latestLog) {
        return $null
    }

    return $latestLog.FullName
}

function Start-Export {
    param(
        [Parameter(Mandatory)] [string]$Destination,
        [Parameter()] [bool]$ForceCloseWord
    )

    $arguments = @(
        '-ExecutionPolicy','Bypass',
        '-NoProfile','-NoLogo',
        '-File',"`"$exportScript`"",
        '-ExportPath',"`"$Destination`""
    )
    if ($ForceCloseWord) { $arguments += '-ForceCloseWord' }

    $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -Wait -PassThru
    return $process.ExitCode
}

try {
    if (-not (Test-Path $exportScript)) {
        Write-Host '[ERRO] export-config.ps1 nao encontrado.' -ForegroundColor Red
        exit 1
    }

    Show-Header

    Write-Host ''
    Write-Host 'Executando exportador...' -ForegroundColor Cyan
    $exitCode = Start-Export -Destination $resolvedExportPath -ForceCloseWord:$ForceCloseWord.IsPresent
    $logPath = Get-LatestLogPath -DestinationPath $resolvedExportPath

    if ($exitCode -eq 0) {
        Write-Host ''
        Write-Host 'Exportacao concluida com sucesso.' -ForegroundColor Green
        Write-Host "Arquivos em: $resolvedExportPath" -ForegroundColor DarkGray
        if ($logPath) {
            Write-Host "Log: $logPath" -ForegroundColor DarkGray
        }
        exit 0
    }

    Write-Host ''
    Write-Host 'Um erro foi identificado durante a exportacao.' -ForegroundColor Red
    if ($logPath) {
        Write-Host "Consulte o log para detalhes: $logPath" -ForegroundColor Yellow
    }
    else {
        Write-Host 'Consulte os logs dentro da pasta de exportacao.' -ForegroundColor Yellow
    }
    exit $exitCode
}
catch {
    Write-Host ''
    Write-Host 'O exportador encontrou uma falha inesperada.' -ForegroundColor Red
    $logPath = Get-LatestLogPath -DestinationPath $resolvedExportPath
    if ($logPath) {
        Write-Host "Consulte o log para detalhes: $logPath" -ForegroundColor Yellow
    }
    else {
        Write-Host 'Consulte os logs dentro da pasta de exportacao.' -ForegroundColor Yellow
    }
    exit 1
}
