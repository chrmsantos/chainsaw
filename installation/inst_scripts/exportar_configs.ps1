[CmdletBinding()]
param(
    [Parameter()] [string]$ExportPath = '.\exported-config'
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent (Get-Item -LiteralPath $MyInvocation.MyCommand.Path).FullName
}
$installationRoot = Split-Path -Parent $scriptRoot
$exportScript = Join-Path $scriptRoot 'export-config.ps1'
$resolvedExportPath = if ([IO.Path]::IsPathRooted($ExportPath)) {
    $ExportPath
} else {
    Join-Path $installationRoot $ExportPath
}

function Show-Header {
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host '  CHAINSAW - Exportador de Configuracoes' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host ''
}

function Ask-YesNo {
    param(
        [Parameter(Mandatory)] [string]$Message,
        [Parameter()] [bool]$DefaultYes = $true
    )

    $suffix = if ($DefaultYes) { '[S/n]' } else { '[s/N]' }
    while ($true) {
        $answer = Read-Host "$Message $suffix"
        if ([string]::IsNullOrWhiteSpace($answer)) {
            return $DefaultYes
        }

        switch ($answer.Trim().ToLowerInvariant()) {
            's' { return $true }
            'sim' { return $true }
            'y' { return $true }
            'yes' { return $true }
            'n' { return $false }
            'nao' { return $false }
            'no' { return $false }
            default { Write-Host 'Digite apenas S ou N.' -ForegroundColor Yellow }
        }
    }
}

function Get-LatestLogPath {
    param(
        [Parameter()] [string]$DestinationPath
    )

    $logDirectory = Join-Path $DestinationPath 'logs'
    if (-not (Test-Path $logDirectory)) {
        return $null
    }

    $file = Get-ChildItem -Path $logDirectory -Filter 'export_*.log' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    return $file?.FullName
}

function Start-Export {
    param(
        [Parameter(Mandatory)] [string]$Destination,
        [Parameter()] [bool]$IncludeRegistry,
        [Parameter()] [bool]$ForceCloseWord
    )

    $arguments = @(
        '-ExecutionPolicy','Bypass',
        '-NoProfile','-NoLogo',
        '-File',"`"$exportScript`"",
        '-ExportPath',"`"$Destination`""
    )

    if ($IncludeRegistry) { $arguments += '-IncludeRegistry' }
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

    $includeRegistry = Ask-YesNo 'Deseja incluir configuracoes do registro?' $false
    $forceWordClose = Ask-YesNo 'Deseja que o Word seja fechado automaticamente?' $true
    $confirm = Ask-YesNo "Confirmar exportacao para '$resolvedExportPath'?"
    if (-not $confirm) {
        Write-Host ''
        Write-Host 'Exportacao cancelada.' -ForegroundColor Yellow
        exit 0
    }

    Write-Host ''
    Write-Host 'Executando exportador...' -ForegroundColor Cyan
    $exitCode = Start-Export -Destination $resolvedExportPath -IncludeRegistry:$includeRegistry -ForceCloseWord:$forceWordClose
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
