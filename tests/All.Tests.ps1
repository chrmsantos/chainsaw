#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1

# Sobrepor Get-RepoRoot caso Helpers.ps1 contenha uma implementação falha
function Get-RepoRoot {
    # Always use the parent of the tests dir where this file lives
    $testsDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $repoRoot = Split-Path -Parent $testsDir
    return $repoRoot
}

Describe 'CHAINSAW - Testes de Integridade' {

    Context 'PowerShell scripts syntax' {
        $scripts = Get-PowerShellScripts
        It 'Encontrar scripts de instalação' {
            $scripts | Should Not BeNullOrEmpty
        }

        foreach ($s in $scripts) {
            It "Checar parse: $($s.Name)" {
                $errors = $null
                $tokens = $null
                { [void] [System.Management.Automation.Language.Parser]::ParseFile($s.FullName, [ref]$tokens, [ref]$errors) } | Should Not Throw
                (($errors -eq $null) -or ($errors.Count -eq 0)) | Should Be $true
            }
        }
    }

    Context 'VBA / BAS files' {
        $basFiles = Get-VbaFiles
        It 'Existe ao menos um módulo monolítico' {
            $basFiles | Where-Object Name -Match 'Módulo1\.bas' | Should Not BeNullOrEmpty
        }

        It 'Não existam backups duplicados com mesmo tamanho' {
            $grouped = $basFiles | Group-Object Length | Where-Object { $_.Count -gt 1 }
            $grouped | Should BeNullOrEmpty
        }
    }

    Context 'Documentação' {
        It 'Existem docs essenciais mínimos' {
            $expected = @('IDENTIFICACAO_ELEMENTOS.md','SEM_PRIVILEGIOS_ADMIN.md','SUBSTITUICOES_CONDICIONAIS.md','VALIDACAO_TIPO_DOCUMENTO.md')
            foreach ($e in $expected) { 
                (Test-Path (Join-Path (Get-RepoRoot) "docs\$e")) | Should Be $true
            }
        }

        It 'Não existam duplicatas markdown na raiz' {
            $rootMd = Get-ChildItem -Path (Get-RepoRoot) -Filter *.md -File
            ($rootMd | Where-Object { $_.Name -in @('INSTALACAO_LOCAL.md','GUIA_INSTALACAO_UNIFICADA.md','INSTALL.md','GUIA_RAPIDO_IDENTIFICACAO.md','GUIA_RAPIDO_EXPORT_IMPORT.md','IMPLEMENTACAO_COMPLETA.md') }).Count | Should Be 0
        }
    }

    Context 'Versão' {
        $versionFile = Get-ProjectFile 'installation\inst_configs\version.json'
        It 'Arquivo de versão existe' {
            Test-Path $versionFile | Should Be $true
        }
        
        It 'Arquivo de versão contém versão 2.0.3' {
            $versionContent = Get-Content $versionFile -Raw | ConvertFrom-Json
            $versionContent.version | Should Be '2.0.3'
        }
    }

    Context 'Test Suites' {
        It 'Existe suite de testes de exportação/instalação' {
            Test-Path (Join-Path (Get-RepoRoot) "tests\Export-Install.Tests.ps1") | Should Be $true
        }

        It 'Existe suite de testes de instalação' {
            Test-Path (Join-Path (Get-RepoRoot) "tests\Installation.Tests.ps1") | Should Be $true
        }

        It 'Existe suite de testes VBA' {
            Test-Path (Join-Path (Get-RepoRoot) "tests\VBA.Tests.ps1") | Should Be $true
        }
        
        It 'Existe suite de testes de backup automático' {
            Test-Path (Join-Path (Get-RepoRoot) "tests\Backup.Tests.ps1") | Should Be $true
        }
    }

}
