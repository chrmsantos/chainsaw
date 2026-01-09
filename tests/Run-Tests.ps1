<# Runner para executar testes Pester do projeto CHAINSAW
   - Importa Pester
   - Executa suites de testes
#>
param(
    [switch]$InstallPester,
    [string]$TestSuite = "All"  # All, VBA, Encoding
)

if ($InstallPester) {
    Write-Host 'Instalando Pester (se necessario) via PowerShellGallery...'
    if (-not (Get-Module -ListAvailable -Name Pester)) {
        Install-Module -Name Pester -Scope CurrentUser -Force -AllowClobber
    }
}

# Importa o modulo
Import-Module Pester -ErrorAction Stop

$pester = Get-Module -ListAvailable -Name Pester | Sort-Object Version -Descending | Select-Object -First 1
$pesterMajor = 0
if ($pester -and $pester.Version) { $pesterMajor = $pester.Version.Major }

Push-Location $PSScriptRoot
try {
    $testScripts = switch ($TestSuite) {
        "VBA" { @("./VBA.Tests.ps1") }
        "Encoding" { @("./Encoding.Tests.ps1") }
        default {
            @(Get-ChildItem -Path . -Filter "*.Tests.ps1" -File | Sort-Object Name | Select-Object -ExpandProperty FullName)
        }
    }

    if (-not $testScripts -or $testScripts.Count -eq 0) {
        throw "Nenhum arquivo '*.Tests.ps1' encontrado em: $PSScriptRoot"
    }

    Write-Host "Executando suite '$TestSuite'" -ForegroundColor Cyan

    if ($pesterMajor -ge 5) {
        $result = Invoke-Pester -Path $testScripts -PassThru
    }
    else {
        $result = Invoke-Pester -Script $testScripts -EnableExit -PassThru
    }

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
