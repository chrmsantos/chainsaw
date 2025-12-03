# ============================================================================
# CHAINSAW - Atualizacao Automatica via GitHub
# ============================================================================
# Versao: 1.0.0
# Licenca: GNU GPLv3
# Descricao: Script para atualizar o CHAINSAW baixando a versao mais recente
#            do repositorio GitHub
# ============================================================================

<#
.SYNOPSIS
    Atualiza o CHAINSAW para a versao mais recente do GitHub.

.DESCRIPTION
    Este script:
    1. Cria backup da instalacao atual (chainsaw -> chainsaw_old)
    2. Baixa a versao mais recente do GitHub (main.zip)
    3. Descompacta e instala a nova versao
    4. Remove o backup se bem-sucedido (ou mantem se especificado)
    
.PARAMETER KeepBackup
    Mantem a pasta chainsaw_old apos atualizacao bem-sucedida.
    
.EXAMPLE
    .\update-from-github.ps1
    Atualiza o CHAINSAW e remove o backup apos sucesso.
    
.EXAMPLE
    .\update-from-github.ps1 -KeepBackup
    Atualiza o CHAINSAW e mantem o backup para seguranca.

.NOTES
    - Requer conexao com a Internet
    - Requer permissoes de escrita em $env:USERPROFILE
    - Em caso de falha, o backup sera restaurado automaticamente
#>

param(
    [Parameter()]
    [switch]$KeepBackup
)

# Configuracao de encoding
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Carregar funcoes de backup
$backupFunctionsPath = Join-Path $PSScriptRoot "backup-functions.ps1"

if (Test-Path $backupFunctionsPath) {
    . $backupFunctionsPath
    Write-Host "[OK] Funcoes de backup carregadas" -ForegroundColor Green
}
else {
    Write-Host "[ERRO] Arquivo backup-functions.ps1 nao encontrado!" -ForegroundColor Red
    Write-Host "Caminho esperado: $backupFunctionsPath" -ForegroundColor Red
    exit 1
}

# Verificar se as funcoes foram carregadas
if (-not (Get-Command Update-ChainsawFromGitHub -ErrorAction SilentlyContinue)) {
    Write-Host "[ERRO] Funcao Update-ChainsawFromGitHub nao foi carregada!" -ForegroundColor Red
    exit 1
}

# Executar atualizacao
Write-Host ""
Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                 CHAINSAW - ATUALIZACAO AUTOMATICA              ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "Este script ira:" -ForegroundColor White
Write-Host "  1. Criar backup da instalacao atual (chainsaw_old)" -ForegroundColor Gray
Write-Host "  2. Baixar versao mais recente do GitHub" -ForegroundColor Gray
Write-Host "  3. Instalar a nova versao" -ForegroundColor Gray

if ($KeepBackup) {
    Write-Host "  4. Manter o backup para seguranca" -ForegroundColor Gray
}
else {
    Write-Host "  4. Remover o backup se instalacao bem-sucedida" -ForegroundColor Gray
}

Write-Host ""
$confirm = Read-Host "Deseja continuar? (S/N)"

if ($confirm -notmatch '^[Ss]$') {
    Write-Host ""
    Write-Host "[INFO] Atualizacao cancelada pelo usuario" -ForegroundColor Cyan
    exit 0
}

# Executar atualizacao
$result = Update-ChainsawFromGitHub -KeepBackup:$KeepBackup

if ($result) {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║              ATUALIZACAO CONCLUIDA COM SUCESSO!                ║" -ForegroundColor Green
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    
    if ($KeepBackup) {
        Write-Host "[INFO] Backup mantido em: $env:USERPROFILE\chainsaw_old" -ForegroundColor Cyan
        Write-Host "[INFO] Voce pode remove-lo manualmente quando desejar" -ForegroundColor Cyan
    }
    
    Write-Host ""
    Write-Host "Proximos passos:" -ForegroundColor White
    Write-Host "  1. Abra o Microsoft Word" -ForegroundColor Gray
    Write-Host "  2. As personalizacoes do CHAINSAW devem estar disponiveis" -ForegroundColor Gray
    Write-Host "  3. Se houver problemas, execute: .\install.ps1" -ForegroundColor Gray
    Write-Host ""
    
    exit 0
}
else {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║                    ATUALIZACAO FALHOU                          ║" -ForegroundColor Red
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Red
    Write-Host ""
    Write-Host "[INFO] O backup foi mantido em: $env:USERPROFILE\chainsaw_old" -ForegroundColor Cyan
    Write-Host "[INFO] Sua instalacao anterior foi restaurada" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Para reportar o problema:" -ForegroundColor White
    Write-Host "  GitHub Issues: https://github.com/chrmsantos/chainsaw/issues" -ForegroundColor Gray
    Write-Host "  Email: chrmsantos@protonmail.com" -ForegroundColor Gray
    Write-Host ""
    
    exit 1
}
