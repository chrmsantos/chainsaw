# =============================================================================
# CHAINSAW - Wrapper para import-config.ps1
# =============================================================================
# Importa personalizacoes do Word a partir de uma pasta exportada.
# =============================================================================

[CmdletBinding()]
param(
    [Parameter()] [string]$ImportPath = '.\\exported-config',
    [Parameter()] [switch]$ForceCloseWord
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent (Get-Item -LiteralPath $MyInvocation.MyCommand.Path).FullName
}
$toolsRoot = Split-Path -Parent $scriptRoot
$repoRoot = if ($toolsRoot) { Split-Path -Parent $toolsRoot } else { $scriptRoot }
$importScript = Join-Path $scriptRoot 'import-config.ps1'
$resolvedImportPath = if ([IO.Path]::IsPathRooted($ImportPath)) { $ImportPath } else { Join-Path $repoRoot $ImportPath }
$forceCloseResolved = $ForceCloseWord.IsPresent -or $env:CHAINSAW_FORCE_CLOSE -or $env:CHAINSAW_NO_PAUSE -or $env:CI -or $env:GITHUB_ACTIONS
$forceEnv = @{
    CHAINSAW_FORCE_CLOSE = $env:CHAINSAW_FORCE_CLOSE
    CHAINSAW_NO_PAUSE   = $env:CHAINSAW_NO_PAUSE
    CI                  = $env:CI
    GITHUB_ACTIONS      = $env:GITHUB_ACTIONS
}

function Show-Header {
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host '  CHAINSAW - Importador de Configuracoes' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host ''
}

function Start-Import {
    param(
        [Parameter(Mandatory)] [string]$SourcePath,
        [Parameter()] [bool]$ForceCloseWord
    )

    $args = @(
        '-ExecutionPolicy','Bypass',
        '-NoProfile','-NoLogo',
        '-File', $importScript,
        '-ImportPath', $SourcePath
    )

    if ($ForceCloseWord) { $args += '-ForceCloseWord' }

    try {
        & 'powershell.exe' @args | Out-Null
        $exitCode = if ([string]::IsNullOrWhiteSpace([string]$LASTEXITCODE)) { 0 } else { [int]$LASTEXITCODE }
        return $exitCode
    }
    catch {
        Write-Host "[ERRO] Falha ao iniciar import-config.ps1: $_" -ForegroundColor Red
        return 1
    }
}

try {
    if (-not (Test-Path $importScript)) {
        Write-Host '[ERRO] import-config.ps1 nao encontrado.' -ForegroundColor Red
        exit 1
    }

    Show-Header

    Write-Host ''
    Write-Host 'Executando importador...' -ForegroundColor Cyan
    Write-Host ("[INFO] flags: ForceCloseWord={0} env={1}" -f $forceCloseResolved, ($forceEnv | ConvertTo-Json -Compress)) -ForegroundColor DarkGray
    $exitCode = [int](Start-Import -SourcePath $resolvedImportPath -ForceCloseWord:$forceCloseResolved)

    Write-Host ("[INFO] exit code capturado: {0}" -f $exitCode) -ForegroundColor DarkGray

    $logDir = Join-Path $resolvedImportPath 'logs'
    $latestLog = $null
    if (Test-Path $logDir) {
        $latestLog = Get-ChildItem -Path $logDir -Filter 'import_*.log' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    }

    if ($exitCode -eq 0) {
        Write-Host ''
        Write-Host 'Importacao concluida com sucesso.' -ForegroundColor Green
        if ($latestLog) { Write-Host ('Log: ' + $latestLog.FullName) -ForegroundColor DarkGray }
        exit 0
    }

    Write-Host ''
    Write-Host 'Um erro foi identificado durante a importacao.' -ForegroundColor Red
    if ($latestLog) { Write-Host ('Consulte o log: ' + $latestLog.FullName) -ForegroundColor DarkGray }
    exit $exitCode
}
catch {
    Write-Host ''
    Write-Host 'O importador encontrou uma falha inesperada.' -ForegroundColor Red
    Write-Host ("[ERRO] Detalhes: {0}" -f $_) -ForegroundColor Red
    exit 1
}
