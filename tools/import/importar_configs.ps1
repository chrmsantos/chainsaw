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
        '-File',"`"$importScript`"",
        '-ImportPath',"`"$SourcePath`""
    )

    if ($ForceCloseWord) { $args += '-ForceCloseWord' }

    $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $args -Wait -PassThru
    return $process.ExitCode
}

try {
    if (-not (Test-Path $importScript)) {
        Write-Host '[ERRO] import-config.ps1 nao encontrado.' -ForegroundColor Red
        exit 1
    }

    Show-Header

    Write-Host ''
    Write-Host 'Executando importador...' -ForegroundColor Cyan
    $exitCode = Start-Import -SourcePath $resolvedImportPath -ForceCloseWord:$ForceCloseWord.IsPresent

    if ($exitCode -eq 0) {
        Write-Host ''
        Write-Host 'Importacao concluida com sucesso.' -ForegroundColor Green
        exit 0
    }

    Write-Host ''
    Write-Host 'Um erro foi identificado durante a importacao.' -ForegroundColor Red
    exit $exitCode
}
catch {
    Write-Host ''
    Write-Host 'O importador encontrou uma falha inesperada.' -ForegroundColor Red
    exit 1
}
