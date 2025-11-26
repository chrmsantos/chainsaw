# CHAINSAW - Monitor de Integridade do Projeto
# Verifica se o projeto esta intacto e alerta sobre problemas

param(
    [string]$ProjectRoot = "C:\Users\csantos\chainsaw"
)

function Test-ProjectIntegrity {
    param([string]$Path)
    
    $issues = @()
    
    # Verifica se diretorio existe
    if (-not (Test-Path $Path)) {
        $issues += "CRITICO: Diretorio do projeto nao existe!"
        return $issues
    }
    
    # Verifica .git
    if (-not (Test-Path (Join-Path $Path ".git"))) {
        $issues += "CRITICO: .git nao encontrado!"
    }
    
    # Verifica diretorios essenciais
    $essential = @("installation", "source", "tests", "docs")
    foreach ($dir in $essential) {
        if (-not (Test-Path (Join-Path $Path $dir))) {
            $issues += "ERRO: Diretorio essencial ausente: $dir"
        }
    }
    
    # Verifica arquivos essenciais
    $files = @("chainsaw_installer.cmd", "README.md", "LICENSE")
    foreach ($file in $files) {
        if (-not (Test-Path (Join-Path $Path $file))) {
            $issues += "AVISO: Arquivo importante ausente: $file"
        }
    }
    
    # Conta arquivos totais
    $fileCount = (Get-ChildItem $Path -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count
    if ($fileCount -lt 50) {
        $issues += "AVISO: Projeto parece incompleto ($fileCount arquivos)"
    }
    
    return $issues
}

Write-Host "`n=== CHAINSAW - Monitor de Integridade ===" -ForegroundColor Cyan
Write-Host "Verificando: $ProjectRoot`n" -ForegroundColor Gray

$problems = Test-ProjectIntegrity -Path $ProjectRoot

if ($problems.Count -eq 0) {
    Write-Host "OK Projeto esta integro!" -ForegroundColor Green
    
    $fileCount = (Get-ChildItem $ProjectRoot -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count
    Write-Host "  Total de arquivos: $fileCount" -ForegroundColor Gray
    Write-Host "  Git: OK" -ForegroundColor Green
    Write-Host "  Diretorios essenciais: OK" -ForegroundColor Green
}
else {
    Write-Host "X PROBLEMAS DETECTADOS!" -ForegroundColor Red
    foreach ($problem in $problems) {
        Write-Host "  - $problem" -ForegroundColor Yellow
    }
    
    Write-Host "`nRECUPERACAO RECOMENDADA:" -ForegroundColor Cyan
    Write-Host "  cd C:\Users\csantos" -ForegroundColor White
    Write-Host "  Remove-Item chainsaw -Recurse -Force -ErrorAction SilentlyContinue" -ForegroundColor White
    Write-Host "  git clone https://github.com/chrmsantos/chainsaw.git chainsaw" -ForegroundColor White
}

Write-Host ""
