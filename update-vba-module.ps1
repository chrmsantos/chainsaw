# =============================================================================
# CHAINSAW - Script de Atualização do Módulo VBA
# =============================================================================
# Versão: 1.0.0
# Licença: GNU GPLv3
# Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
# =============================================================================

<#
.SYNOPSIS
    Importa o módulo VBA mais recente (monolithicMod.bas) para o Normal.dotm

.DESCRIPTION
    Este script automatiza a importação do módulo VBA para o template Normal.dotm,
    garantindo que o código mais recente seja utilizado pelo Word.

.PARAMETER Force
    Força a atualização sem confirmação

.EXAMPLE
    .\update-vba-module.ps1
    Importa o módulo com confirmação do usuário

.EXAMPLE
    .\update-vba-module.ps1 -Force
    Importa o módulo sem confirmação
#>

[CmdletBinding()]
param(
    [Parameter()]
    [switch]$Force
)

# Define caminhos
$ScriptPath = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ScriptPath)) {
    $ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$VbaModulePath = Join-Path $ScriptPath "src\monolithicMod.bas"
$NormalDotmPath = "$env:APPDATA\Microsoft\Templates\Normal.dotm"

# Banner
Write-Host ""
Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║     CHAINSAW - Atualização do Módulo VBA                      ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Validações
Write-Host "🔍 Validando arquivos..." -ForegroundColor Yellow

if (-not (Test-Path $VbaModulePath)) {
    Write-Host "❌ Erro: Módulo VBA não encontrado!" -ForegroundColor Red
    Write-Host "   Esperado: $VbaModulePath" -ForegroundColor Gray
    exit 1
}
Write-Host "✓ Módulo VBA encontrado" -ForegroundColor Green

if (-not (Test-Path $NormalDotmPath)) {
    Write-Host "❌ Erro: Normal.dotm não encontrado!" -ForegroundColor Red
    Write-Host "   Esperado: $NormalDotmPath" -ForegroundColor Gray
    exit 1
}
Write-Host "✓ Normal.dotm encontrado" -ForegroundColor Green

# Verifica se o Word está aberto
$wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
if ($wordProcesses) {
    Write-Host ""
    Write-Host "⚠️  ATENÇÃO: Word está aberto!" -ForegroundColor Yellow
    Write-Host "   Por favor, feche o Word antes de continuar." -ForegroundColor Yellow
    Write-Host ""
    
    $response = Read-Host "Deseja que o script feche o Word automaticamente? (S/N)"
    if ($response -eq 'S' -or $response -eq 's') {
        Write-Host "Fechando Word..." -ForegroundColor Yellow
        $wordProcesses | ForEach-Object {
            $_.CloseMainWindow() | Out-Null
            Start-Sleep -Seconds 2
            if (-not $_.HasExited) {
                $_ | Stop-Process -Force
            }
        }
        Start-Sleep -Seconds 2
        Write-Host "✓ Word fechado" -ForegroundColor Green
    } else {
        Write-Host "Operação cancelada pelo usuário." -ForegroundColor Red
        exit 0
    }
}

# Confirmação
if (-not $Force) {
    Write-Host ""
    Write-Host "📋 Operação a ser realizada:" -ForegroundColor Cyan
    Write-Host "   • Fazer backup do módulo atual (se existir)" -ForegroundColor White
    Write-Host "   • Importar: monolithicMod.bas" -ForegroundColor White
    Write-Host "   • Destino: Normal.dotm" -ForegroundColor White
    Write-Host ""
    
    $response = Read-Host "Deseja continuar? (S/N)"
    if ($response -ne 'S' -and $response -ne 's') {
        Write-Host "Operação cancelada pelo usuário." -ForegroundColor Yellow
        exit 0
    }
}

Write-Host ""
Write-Host "🔄 Atualizando módulo VBA..." -ForegroundColor Cyan
Write-Host ""

try {
    # Cria objeto Word
    Write-Host "   [1/5] Iniciando Word..." -ForegroundColor Gray
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0  # wdAlertsNone
    
    # Abre Normal.dotm
    Write-Host "   [2/5] Abrindo Normal.dotm..." -ForegroundColor Gray
    $doc = $word.Documents.Open($NormalDotmPath, $false, $false)
    
    # Remove módulo antigo se existir
    Write-Host "   [3/5] Removendo módulo antigo (se existir)..." -ForegroundColor Gray
    $vbProject = $doc.VBProject
    $moduleRemoved = $false
    
    # Lista de nomes possíveis do módulo antigo
    $oldModuleNames = @("Módulo1", "Module1", "monolithicMod", "Mod_Main", "Chainsaw", "CHAINSAW_MODX", "Chainsaw_ModX", "chainsawModX")
    
    foreach ($moduleName in $oldModuleNames) {
        try {
            $module = $vbProject.VBComponents.Item($moduleName)
            if ($module) {
                # Faz backup do módulo antigo
                $backupPath = Join-Path $ScriptPath "src\backup_$moduleName`_$(Get-Date -Format 'yyyyMMdd_HHmmss').bas"
                $module.Export($backupPath)
                Write-Host "      ✓ Backup criado: $backupPath" -ForegroundColor DarkGreen
                
                # Remove o módulo
                $vbProject.VBComponents.Remove($module)
                Write-Host "      ✓ Módulo '$moduleName' removido" -ForegroundColor DarkGreen
                $moduleRemoved = $true
            }
        }
        catch {
            # Módulo não existe, continua
        }
    }
    
    if (-not $moduleRemoved) {
        Write-Host "      ℹ Nenhum módulo antigo encontrado" -ForegroundColor DarkGray
    }
    
    # Importa novo módulo
    Write-Host "   [4/5] Importando novo módulo..." -ForegroundColor Gray
    $vbProject.VBComponents.Import($VbaModulePath) | Out-Null
    Write-Host "      ✓ Módulo 'monolithicMod' importado" -ForegroundColor DarkGreen
    
    # Salva e fecha
    Write-Host "   [5/5] Salvando alterações..." -ForegroundColor Gray
    $doc.Save()
    $doc.Close($false)
    $word.Quit()
    
    # Libera objetos COM
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║     ✓ MÓDULO VBA ATUALIZADO COM SUCESSO!                      ║" -ForegroundColor Green
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "O módulo mais recente foi importado para o Normal.dotm." -ForegroundColor White
    Write-Host "Você já pode abrir o Word e usar o CHAINSAW v1.1" -ForegroundColor White
    Write-Host ""
    
}
catch {
    Write-Host ""
    Write-Host "❌ ERRO ao atualizar módulo:" -ForegroundColor Red
    Write-Host "   $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Possíveis causas:" -ForegroundColor Yellow
    Write-Host "   • Acesso à macro de VBA pode estar bloqueado" -ForegroundColor Yellow
    Write-Host "   • Configurações de segurança do Word" -ForegroundColor Yellow
    Write-Host "   • Word ainda está em execução em segundo plano" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Solução alternativa - Importação Manual:" -ForegroundColor Cyan
    Write-Host "   1. Abra o Word" -ForegroundColor White
    Write-Host "   2. Pressione Alt + F11 (abre o editor VBA)" -ForegroundColor White
    Write-Host "   3. Clique em 'Arquivo' > 'Importar Arquivo'" -ForegroundColor White
    Write-Host "   4. Selecione: $VbaModulePath" -ForegroundColor White
    Write-Host "   5. Feche o editor VBA e salve" -ForegroundColor White
    Write-Host ""
    
    # Cleanup
    if ($word) {
        try { $word.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
    }
    
    exit 1
}

Write-Host "Pressione qualquer tecla para sair..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
