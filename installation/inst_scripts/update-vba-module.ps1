# =============================================================================
<#
CH AINS AW - Atualização do Módulo VBA
Versão: 1.0.1
Autor: chrmsantos
#>

[CmdletBinding()]
param(
    [Parameter()]
    [switch]$Force
)

# Determina o diretório do script de forma robusta
$ScriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $ScriptPath

# Caminhos importantes
$VbaModulePath = Join-Path $ProjectRoot "source\main\monolithicMod.bas"
$NormalDotmPath = Join-Path $env:APPDATA "Microsoft\Templates\Normal.dotm"

function Write-Banner {
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host " CHAINSAW - Atualização do Módulo VBA" -ForegroundColor Cyan
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
}

Write-Banner

Write-Host "Validando arquivos..." -ForegroundColor Yellow
if (-not (Test-Path $VbaModulePath)) {
    Write-Host "Erro: Módulo VBA não encontrado: $VbaModulePath" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $NormalDotmPath)) {
    Write-Host "Erro: Normal.dotm não encontrado: $NormalDotmPath" -ForegroundColor Red
    exit 1
}

# Fecha Word se o usuário permitir
$wordProcesses = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
if ($wordProcesses) {
    Write-Host "Word em execução detectado." -ForegroundColor Yellow
    if (-not $Force) {
        $resp = Read-Host "Fechar Word automaticamente e continuar? (S/N)"
        if ($resp -ne 'S' -and $resp -ne 's') { Write-Host "Operação cancelada." -ForegroundColor Yellow; exit 0 }
    }
    $wordProcesses | ForEach-Object { try { $_.CloseMainWindow() | Out-Null; Start-Sleep -Seconds 1; if (-not $_.HasExited) { $_ | Stop-Process -Force } } catch { } }
}

try {
    Write-Host "Iniciando Word em modo invisível..." -ForegroundColor Gray
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($NormalDotmPath, $false, $false)

    $vbProject = $doc.VBProject
    $moduleRemoved = $false
    $oldModuleNames = @('Módulo1','Module1','monolithicMod','Mod_Main','Chainsaw','CHAINSAW_MODX','Chainsaw_ModX','chainsawModX')
    foreach ($name in $oldModuleNames) {
        try {
            $comp = $vbProject.VBComponents.Item($name)
            if ($comp) {
                $backup = Join-Path $ScriptPath "backup_${name}_$(Get-Date -Format yyyyMMdd_HHmmss).bas"
                $comp.Export($backup)
                $vbProject.VBComponents.Remove($comp)
                Write-Host "Backup e remoção do módulo $name" -ForegroundColor DarkGreen
                $moduleRemoved = $true
            }
        } catch { }
    }

    Write-Host "Importando novo módulo: $VbaModulePath" -ForegroundColor Gray
    $vbProject.VBComponents.Import($VbaModulePath) | Out-Null

    $doc.Save()
    $doc.Close($false)
    $word.Quit()

    # Garbage collect
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host "Módulo VBA atualizado com sucesso." -ForegroundColor Green
}
catch {
    Write-Host "Erro ao atualizar módulo: $_" -ForegroundColor Red
    if ($word) { try { $word.Quit() } catch { } }
    exit 1
}
        Write-Host "Fechando Word..." -ForegroundColor Yellow

