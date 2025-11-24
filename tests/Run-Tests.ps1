<# Runner para executar testes Pester do projeto CHAINSAW
   - Garante que Pester v5 está disponível
   - Executa `All.Tests.ps1`
#>
param(
    [switch]$InstallPester
)

if ($InstallPester) {
    Write-Host 'Instalando Pester (se necessário) via PowerShellGallery...'
    if (-not (Get-Module -ListAvailable -Name Pester)) {
        Install-Module -Name Pester -Scope CurrentUser -Force -AllowClobber
    }
}

# Importa o módulo
Import-Module Pester -ErrorAction Stop

Push-Location $PSScriptRoot
try {
    # Executa todos os testes
    $allTests = @(
        '.\All.Tests.ps1'
        '.\Export-Install.Tests.ps1'
    )
    
    $totalFailed = 0
    $totalPassed = 0
    
    foreach ($testFile in $allTests) {
        if (Test-Path $testFile) {
            Write-Host "`nExecutando: $testFile" -ForegroundColor Cyan
            $result = Invoke-Pester -Script $testFile -PassThru
            $totalFailed += $result.FailedCount
            $totalPassed += $result.PassedCount
        }
    }
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "RESUMO FINAL:" -ForegroundColor White
    Write-Host "  Passou: $totalPassed" -ForegroundColor Green
    Write-Host "  Falhou: $totalFailed" -ForegroundColor $(if ($totalFailed -gt 0) { 'Red' } else { 'Green' })
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    if ($totalFailed -gt 0) {
        Write-Host "Alguns testes falharam: $totalFailed" -ForegroundColor Red
        exit 1
    }
    Write-Host 'Todos os testes passaram.' -ForegroundColor Green
    exit 0
}
finally {
    Pop-Location
}
