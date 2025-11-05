# =============================================================================
# CHAINSAW - Script de Teste de Instalação
# =============================================================================
# Este script testa a instalação do Chainsaw em modo simulado (sem fazer alterações reais)
# =============================================================================

[CmdletBinding()]
param(
    [Parameter()]
    [string]$SourcePath = "\\strqnapmain\Dir. Legislativa\_Christian261\chainsaw"
)

Write-Host ""
Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║        CHAINSAW - Teste de Pré-requisitos de Instalação       ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "Este script verifica se o sistema está pronto para instalação." -ForegroundColor Gray
Write-Host "Nenhuma alteração será feita no sistema." -ForegroundColor Gray
Write-Host ""

$allChecks = $true

# Teste 1: Versão do Windows
Write-Host "1. Verificando versão do Windows..." -NoNewline
$osVersion = [Environment]::OSVersion.Version
if ($osVersion.Major -ge 10) {
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "   Windows $($osVersion.Major).$($osVersion.Minor)" -ForegroundColor Gray
} else {
    Write-Host " ✗" -ForegroundColor Red
    Write-Host "   Windows 10 ou superior necessário. Versão atual: $($osVersion.Major).$($osVersion.Minor)" -ForegroundColor Red
    $allChecks = $false
}

# Teste 2: Versão do PowerShell
Write-Host ""
Write-Host "2. Verificando versão do PowerShell..." -NoNewline
$psVersion = $PSVersionTable.PSVersion
if ($psVersion.Major -ge 5) {
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "   PowerShell $($psVersion.ToString())" -ForegroundColor Gray
} else {
    Write-Host " ✗" -ForegroundColor Red
    Write-Host "   PowerShell 5.1 ou superior necessário. Versão atual: $($psVersion.ToString())" -ForegroundColor Red
    $allChecks = $false
}

# Teste 3: Acesso ao caminho de rede
Write-Host ""
Write-Host "3. Verificando acesso ao caminho de rede..." -NoNewline
if (Test-Path $SourcePath) {
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "   $SourcePath" -ForegroundColor Gray
} else {
    Write-Host " ✗" -ForegroundColor Red
    Write-Host "   Não foi possível acessar: $SourcePath" -ForegroundColor Red
    $allChecks = $false
}

# Teste 4: Arquivos necessários
Write-Host ""
Write-Host "4. Verificando arquivos necessários..." -NoNewline
$stampFile = Join-Path $SourcePath "assets\stamp.png"
$templatesFolder = Join-Path $SourcePath "configs\Templates"
$installScript = Join-Path $SourcePath "install.ps1"

$missingFiles = @()
if (-not (Test-Path $stampFile)) { $missingFiles += "assets\stamp.png" }
if (-not (Test-Path $templatesFolder)) { $missingFiles += "configs\Templates" }
if (-not (Test-Path $installScript)) { $missingFiles += "install.ps1" }

if ($missingFiles.Count -eq 0) {
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "   Todos os arquivos encontrados" -ForegroundColor Gray
} else {
    Write-Host " ✗" -ForegroundColor Red
    Write-Host "   Arquivos ausentes:" -ForegroundColor Red
    foreach ($file in $missingFiles) {
        Write-Host "   - $file" -ForegroundColor Red
    }
    $allChecks = $false
}

# Teste 5: Permissões de escrita
Write-Host ""
Write-Host "5. Verificando permissões de escrita..." -NoNewline
$testFile = Join-Path $env:USERPROFILE "chainsaw_test_$(Get-Date -Format 'yyyyMMddHHmmss').tmp"
try {
    [System.IO.File]::WriteAllText($testFile, "test")
    Remove-Item $testFile -Force -ErrorAction SilentlyContinue
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "   Permissões OK em: $env:USERPROFILE" -ForegroundColor Gray
} catch {
    Write-Host " ✗" -ForegroundColor Red
    Write-Host "   Sem permissões de escrita em: $env:USERPROFILE" -ForegroundColor Red
    $allChecks = $false
}

# Teste 6: Word em execução
Write-Host ""
Write-Host "6. Verificando se o Word está em execução..." -NoNewline
$wordProcess = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
if ($wordProcess) {
    Write-Host " ⚠" -ForegroundColor Yellow
    Write-Host "   Word está em execução. Feche-o antes de instalar." -ForegroundColor Yellow
} else {
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "   Word não está em execução" -ForegroundColor Gray
}

# Teste 7: Pasta Templates existente
Write-Host ""
Write-Host "7. Verificando pasta Templates atual..." -NoNewline
$templatesPath = Join-Path $env:APPDATA "Microsoft\Templates"
if (Test-Path $templatesPath) {
    $itemCount = (Get-ChildItem -Path $templatesPath -Recurse -File | Measure-Object).Count
    Write-Host " ⚠" -ForegroundColor Yellow
    Write-Host "   Pasta existe ($itemCount arquivos) - será feito backup" -ForegroundColor Yellow
} else {
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "   Pasta não existe - instalação limpa" -ForegroundColor Gray
}

# Resultado final
Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray

if ($allChecks) {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║         ✓ SISTEMA PRONTO PARA INSTALAÇÃO!                     ║" -ForegroundColor Green
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "Execute o comando abaixo para instalar:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  .\install.ps1" -ForegroundColor White
    Write-Host ""
} else {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║         ✗ SISTEMA NÃO ESTÁ PRONTO PARA INSTALAÇÃO             ║" -ForegroundColor Red
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Red
    Write-Host ""
    Write-Host "Corrija os problemas acima antes de prosseguir." -ForegroundColor Yellow
    Write-Host ""
}
