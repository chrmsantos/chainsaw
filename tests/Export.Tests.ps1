#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1

Describe 'CHAINSAW - Testes de Exportacao' {

    BeforeAll {
        $repoRoot = Get-RepoRoot
        $scriptsPath = Join-Path $repoRoot "tools\export"
        $exportScript = Join-Path $scriptsPath "export-config.ps1"
        $exportContent = Get-Content $exportScript -Raw
    }

    Context 'export-config.ps1 - Estrutura' {

        It 'Script existe e eh legivel' {
            Test-Path $exportScript | Should Be $true
            $exportContent.Length | Should BeGreaterThan 0
        }

        It 'Usa CmdletBinding' {
            $exportContent -match '\\[CmdletBinding\\(\\)\\]' | Should Be $true
        }

        It 'Contem funcoes principais de exportacao' {
            $exportContent -match 'function Test-VbaModuleCompilation' | Should Be $true
            $exportContent -match 'function Export-VbaModule' | Should Be $true
            $exportContent -match 'function Export-RibbonCustomization' | Should Be $true
            $exportContent -match 'function Export-OfficeCustomUI' | Should Be $true
            $exportContent -match 'function New-ExportManifest' | Should Be $true
        }

        It 'Sintaxe PowerShell valida' {
            $errors = $null
            $null = [System.Management.Automation.PSParser]::Tokenize($exportContent, [ref]$errors)
            $errors.Count | Should Be 0
        }
    }

    Context 'export-config.ps1 - Parametros' {

        It 'Aceita parametro ExportPath' {
            $exportContent -match '\\[string\\]\$ExportPath' | Should Be $true
        }

        It 'Define valor padrao exported-config' {
            $exportContent -match 'ExportPath\s*=\s*"\.\\exported-config"' | Should Be $true
        }

        It 'Aceita IncludeRegistry e ForceCloseWord' {
            $exportContent -match '\\[switch\\]\$IncludeRegistry' | Should Be $true
            $exportContent -match '\\[switch\\]\$ForceCloseWord' | Should Be $true
        }
    }

    Context 'Exportacao - Estrutura de diretorios' {

        It 'Prevê pastas de saida esperadas' {
            $exportContent -match 'VBAModule' | Should Be $true
            $exportContent -match 'RibbonCustomization' | Should Be $true
            $exportContent -match 'OfficeCustomUI' | Should Be $true
        }

        It 'Referencia Normal.dotm e Office UI' {
            $exportContent -match 'Normal\.dotm' | Should Be $true
            $exportContent -match '\\.officeUI' | Should Be $true
        }
    }

    Context 'Exportacao - Manifesto' {

        It 'Cria manifesto de exportacao' {
            $exportContent -match 'New-ExportManifest' | Should Be $true
            $exportContent -match 'ExportDate' | Should Be $true
            $exportContent -match 'USERNAME|ComputerName' | Should Be $true
        }
    }

    Context 'Feedback e logs' {

        It 'Usa Write-Host com cores para feedback' {
            $exportContent -match '-ForegroundColor' | Should Be $true
        }

        It 'Aplica retencao de logs' {
            $exportContent -match 'Invoke-LogRetention' | Should Be $true
        }
    }
}
