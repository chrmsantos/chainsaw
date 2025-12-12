#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1

Describe 'CHAINSAW - Import/Export Logging' {
    BeforeAll {
        $repoRoot = Get-RepoRoot
        $importPs1 = Join-Path $repoRoot 'tools\import\import-config.ps1'
        $importWrapper = Join-Path $repoRoot 'tools\import\importar_configs.ps1'
        $importCmd = Join-Path $repoRoot 'tools\import\importar_configs.cmd'
        $exportPs1 = Join-Path $repoRoot 'tools\export\export-config.ps1'

        $importContent = Get-Content $importPs1 -Raw -Encoding UTF8
        $importWrapperContent = Get-Content $importWrapper -Raw -Encoding UTF8
        $importCmdContent = Get-Content $importCmd -Raw -Encoding ASCII
        $exportContent = Get-Content $exportPs1 -Raw -Encoding UTF8
    }

    Context 'Logging helpers present' {
        It 'import-config.ps1 define Initialize-LogFile/Invoke-LogRetention/Write-Log' {
            $importContent -match 'Initialize-LogFile' | Should -BeTrue
            $importContent -match 'Invoke-LogRetention' | Should -BeTrue
            $importContent -match 'Write-Log' | Should -BeTrue
        }

        It 'export-config.ps1 define Initialize-LogFile/Invoke-LogRetention/Write-Log' {
            $exportContent -match 'Initialize-LogFile' | Should -BeTrue
            $exportContent -match 'Invoke-LogRetention' | Should -BeTrue
            $exportContent -match 'Write-Log' | Should -BeTrue
        }
    }

    Context 'Log naming and retention' {
        It 'import-config.ps1 cria logs com prefixo import_' {
            $importContent -match 'import_\$\(Get-Date' | Should -BeTrue
        }

        It 'export-config.ps1 cria logs com prefixo export_' {
            $exportContent -match 'export_\$\(Get-Date' | Should -BeTrue
        }

        It 'scripts mantêm apenas 5 logs (Invoke-LogRetention KeepLatest 5)' {
            $importContent -match 'KeepLatest 5' | Should -BeTrue
            $exportContent -match 'KeepLatest 5' | Should -BeTrue
        }
    }

    Context 'Wrapper behavior' {
        It 'importar_configs.ps1 propaga exit code via LASTEXITCODE' {
            $importWrapperContent -match '\$exitCode = if \(\$null -ne \$LASTEXITCODE\)' | Should -BeTrue
        }

        It 'importar_configs.ps1 mostra log mais recente se existir' {
            $importWrapperContent -match 'import_\*\.log' | Should -BeTrue
        }

        It 'importar_configs.cmd permite pular pausa com CHAINSAW_NO_PAUSE' {
            $importCmdContent -match 'CHAINSAW_NO_PAUSE' | Should -BeTrue
        }
    }
}
