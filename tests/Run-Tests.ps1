<# Runner para executar testes Pester do projeto CHAINSAW
   - Garante que Pester v5 está disponível
   - Executa todos os testes
#>
param(
    [switch]$InstallPester,
    [string]$TestSuite = "All"  # All, Export, VBA
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
    $testScript = switch ($TestSuite) {
        "Export" { "./Export.Tests.ps1" }
        "VBA" { "./VBA.Tests.ps1" }
        default { "./All.Tests.ps1" }
    }

    Write-Host "Executando: $testScript" -ForegroundColor Cyan
    $result = Invoke-Pester -Script $testScript -EnableExit -PassThru

    if ($result.FailedCount -gt 0) {
        Write-Host "Alguns testes falharam: $($result.FailedCount)" -ForegroundColor Red
        exit 1
    }
    Write-Host 'Todos os testes passaram.' -ForegroundColor Green
    exit 0
}
finally {
    Pop-Location
}
