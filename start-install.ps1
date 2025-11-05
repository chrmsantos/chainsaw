<#
.SYNOPSIS
    Launcher seguro para install.ps1 com bypass automático.

.DESCRIPTION
    Este script garante que install.ps1 seja executado mesmo com políticas
    restritivas de execução, usando bypass temporário seguro.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$SourcePath = "\\strqnapmain\Dir. Legislativa\_Christian261\chainsaw",
    
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$NoBackup
)

Write-Host "🔒 Chainsaw - Launcher Seguro" -ForegroundColor Cyan
Write-Host ""

# Determina o caminho do script de instalação
$scriptPath = Join-Path $PSScriptRoot "install.ps1"

if (-not (Test-Path $scriptPath)) {
    Write-Host "✗ Erro: install.ps1 não encontrado em: $scriptPath" -ForegroundColor Red
    exit 1
}

Write-Host "ℹ  Executando install.ps1 com bypass temporário seguro..." -ForegroundColor Cyan
Write-Host ""
Write-Host "🔐 GARANTIAS DE SEGURANÇA:" -ForegroundColor Green
Write-Host "   • Apenas o install.ps1 será executado" -ForegroundColor Gray
Write-Host "   • A política do sistema NÃO será alterada" -ForegroundColor Gray
Write-Host "   • O bypass expira quando o script terminar" -ForegroundColor Gray
Write-Host "   • Nenhum privilégio de administrador é usado" -ForegroundColor Gray
Write-Host ""

# Constrói argumentos
$arguments = @(
    "-ExecutionPolicy", "Bypass",
    "-NoProfile",
    "-File", "`"$scriptPath`""
)

if ($SourcePath -ne "\\strqnapmain\Dir. Legislativa\_Christian261\chainsaw") {
    $arguments += @("-SourcePath", "`"$SourcePath`"")
}
if ($Force) {
    $arguments += "-Force"
}
if ($NoBackup) {
    $arguments += "-NoBackup"
}

# Executa install.ps1
try {
    $processInfo = Start-Process -FilePath "powershell.exe" `
                                 -ArgumentList $arguments `
                                 -Wait `
                                 -NoNewWindow `
                                 -PassThru
    
    exit $processInfo.ExitCode
}
catch {
    Write-Host ""
    Write-Host "✗ Erro ao executar install.ps1: $_" -ForegroundColor Red
    exit 1
}
