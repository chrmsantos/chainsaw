#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1

function Get-RepoRoot {
    $testsDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $repoRoot = Split-Path -Parent $testsDir
    return $repoRoot
}

Describe 'CHAINSAW - Testes de Exportação e Instalação' {

    BeforeAll {
        $repoRoot = Get-RepoRoot
        $scriptsPath = Join-Path $repoRoot "installation\inst_scripts"
        $installScript = Join-Path $scriptsPath "install.ps1"
        $exportScript = Join-Path $scriptsPath "export-config.ps1"
        
        # Caminhos de teste
        $testExportPath = Join-Path $env:TEMP "chainsaw_export_test_$(Get-Date -Format 'yyyyMMddHHmmss')"
        $testBackupPath = Join-Path $env:TEMP "chainsaw_backup_test_$(Get-Date -Format 'yyyyMMddHHmmss')"
    }

    AfterAll {
        # Limpeza
        if (Test-Path $testExportPath) {
            Remove-Item $testExportPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path $testBackupPath) {
            Remove-Item $testBackupPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    Context 'export-config.ps1 - Validação de Estrutura' {
        
        BeforeAll {
            $exportContent = Get-Content $exportScript -Raw
        }

        It 'Script existe e é legível' {
            Test-Path $exportScript | Should Be $true
            $exportContent.Length | Should BeGreaterThan 0
        }

        It 'Contém função Test-VbaModuleCompilation' {
            $exportContent -match 'function Test-VbaModuleCompilation' | Should Be $true
        }

        It 'Contém função Export-VbaModule' {
            $exportContent -match 'function Export-VbaModule' | Should Be $true
        }

        It 'Contém função Export-RibbonCustomization' {
            $exportContent -match 'function Export-RibbonCustomization' | Should Be $true
        }

        It 'Contém função Export-OfficeCustomUI' {
            $exportContent -match 'function Export-OfficeCustomUI' | Should Be $true
        }

        It 'Contém função New-ExportManifest' {
            $exportContent -match 'function New-ExportManifest' | Should Be $true
        }

        It 'Não usa verbos não aprovados (Compile-VbaModule)' {
            $exportContent -match 'function Compile-VbaModule' | Should Be $false
        }

        It 'Não usa verbos não aprovados (Create-ExportManifest)' {
            $exportContent -match 'function Create-ExportManifest' | Should Be $false
        }

        It 'Usa comparação de null correta' {
            # Deve usar $null -eq, não -eq $null
            $badPattern = '\$\w+\s+-eq\s+\$null'
            $matches = [regex]::Matches($exportContent, $badPattern)
            $matches.Count | Should Be 0
        }

        It 'Sintaxe PowerShell válida' {
            $errors = $null
            $null = [System.Management.Automation.PSParser]::Tokenize($exportContent, [ref]$errors)
            $errors.Count | Should Be 0
        }
    }

    Context 'install.ps1 - Validação de Estrutura' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
        }

        It 'Script existe e é legível' {
            Test-Path $installScript | Should Be $true
            $installContent.Length | Should BeGreaterThan 0
        }

        It 'Contém função Import-VbaModule' {
            $installContent -match 'function Import-VbaModule' | Should Be $true
        }

        It 'Contém função Import-RibbonCustomization' {
            $installContent -match 'function Import-RibbonCustomization' | Should Be $true
        }

        It 'Contém função Import-OfficeCustomUI' {
            $installContent -match 'function Import-OfficeCustomUI' | Should Be $true
        }

        It 'Contém função Backup-CompleteConfiguration' {
            $installContent -match 'function Backup-CompleteConfiguration' | Should Be $true
        }

        It 'Contém função Copy-TemplatesFolder' {
            $installContent -match 'function Copy-TemplatesFolder' | Should Be $true
        }

        It 'Contém função Update-VbaModule' {
            $installContent -match 'function Update-VbaModule' | Should Be $true
        }

        It 'Usa comparação de null correta' {
            $badPattern = '\$\w+\s+-eq\s+\$null'
            $matches = [regex]::Matches($installContent, $badPattern)
            $matches.Count | Should Be 0
        }

        It 'Sintaxe PowerShell válida' {
            $errors = $null
            $null = [System.Management.Automation.PSParser]::Tokenize($installContent, [ref]$errors)
            $errors.Count | Should Be 0
        }
    }

    Context 'export-config.ps1 - Validação de Parâmetros' {
        
        It 'Aceita parâmetro ExportPath' {
            $exportContent = Get-Content $exportScript -Raw
            $exportContent -match '\[Parameter.*\]\s*\[string\]\$ExportPath' | Should Be $true
        }

        It 'Tem valor padrão para ExportPath' {
            $exportContent = Get-Content $exportScript -Raw
            $exportContent -match 'ExportPath.*=.*exported-config' | Should Be $true
        }
    }

    Context 'install.ps1 - Validação de Parâmetros' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
        }

        It 'Aceita parâmetro Force' {
            $installContent -match '\[switch\]\$Force' | Should Be $true
        }

        It 'Aceita parâmetro SkipBackup' {
            $installContent -match '\[switch\]\$SkipBackup' | Should Be $true
        }

        It 'Aceita parâmetro SkipCustomizations' {
            $installContent -match '\[switch\]\$SkipCustomizations' | Should Be $true
        }
    }

    Context 'Exportação - Estrutura de Diretórios' {
        
        It 'Deve criar estrutura de pastas para exportação' {
            $expectedFolders = @(
                'VBAModule',
                'RibbonCustomization',
                'OfficeCustomUI'
            )
            
            # Simulação: verifica que as funções tentam criar essas pastas
            $exportContent = Get-Content $exportScript -Raw
            foreach ($folder in $expectedFolders) {
                $exportContent -match $folder | Should Be $true
            }
        }

        It 'Deve exportar módulo VBA para VBAModule/monolithicMod.bas' {
            $exportContent = Get-Content $exportScript -Raw
            $exportContent -match 'monolithicMod\.bas' | Should Be $true
        }

        It 'Deve procurar arquivos .officeUI' {
            $exportContent = Get-Content $exportScript -Raw
            $exportContent -match '\.officeUI' | Should Be $true
        }
    }

    Context 'Instalação - Backup Completo' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
        }

        It 'Deve criar backup antes da instalação' {
            $installContent -match 'Backup-CompleteConfiguration' | Should Be $true
        }

        It 'Backup deve incluir Templates' {
            $installContent -match 'Templates.*backup|backup.*Templates' | Should Be $true
        }

        It 'Backup deve criar manifesto' {
            $installContent -match 'manifest\.json|backup_manifest' | Should Be $true
        }

        It 'Backup deve incluir Normal.dotm' {
            $installContent -match 'Normal\.dotm' | Should Be $true
        }
    }

    Context 'Exportação - Manifesto' {
        
        BeforeAll {
            $exportContent = Get-Content $exportScript -Raw
        }

        It 'Deve criar manifesto de exportação' {
            $exportContent -match 'New-ExportManifest|manifesto.*exporta' | Should Be $true
        }

        It 'Manifesto deve incluir data de exportação' {
            $exportContent -match 'ExportDate|Data.*exporta' | Should Be $true
        }

        It 'Manifesto deve incluir nome do usuário' {
            $exportContent -match 'UserName|USERNAME' | Should Be $true
        }

        It 'Manifesto deve incluir computador' {
            $exportContent -match 'ComputerName|COMPUTERNAME' | Should Be $true
        }
    }

    Context 'Instalação - Validação de Pré-requisitos' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
        }

        It 'Deve verificar se Word está em execução' {
            $installContent -match 'Test-WordRunning|Word.*execu' | Should Be $true
        }

        It 'Deve verificar versão do PowerShell' {
            $installContent -match 'PSVersion|PowerShell.*vers' | Should Be $true
        }

        It 'Deve verificar sistema operacional' {
            $installContent -match 'Windows|Sistema.*operacional' | Should Be $true
        }

        It 'Deve verificar permissões de escrita' {
            $installContent -match 'permiss.*escrita|write.*permission' | Should Be $true
        }
    }

    Context 'Exportação - Validação VBA' {
        
        BeforeAll {
            $exportContent = Get-Content $exportScript -Raw
        }

        It 'Deve compilar módulo VBA antes de exportar' {
            $exportContent -match 'Test-VbaModuleCompilation' | Should Be $true
        }

        It 'Deve permitir continuar mesmo com erros de compilação' {
            $exportContent -match 'continuar.*exporta.*mesmo assim|continue.*export.*anyway' | Should Be $true
        }

        It 'Deve acessar VBProject do template' {
            $exportContent -match 'VBProject|\.VBComponents' | Should Be $true
        }
    }

    Context 'Instalação - Importação de Customizações' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
        }

        It 'Deve importar módulo VBA' {
            $installContent -match 'Import-VbaModule' | Should Be $true
        }

        It 'Deve importar Ribbon customization' {
            $installContent -match 'Import-RibbonCustomization' | Should Be $true
        }

        It 'Deve importar Office Custom UI' {
            $installContent -match 'Import-OfficeCustomUI' | Should Be $true
        }

        It 'Deve copiar para LOCALAPPDATA' {
            $installContent -match 'LOCALAPPDATA|LocalAppData' | Should Be $true
        }

        It 'Deve copiar para APPDATA' {
            $installContent -match 'env:APPDATA|APPDATA' | Should Be $true
        }
    }

    Context 'Exportação - Gerenciamento de Erros' {
        
        BeforeAll {
            $exportContent = Get-Content $exportScript -Raw
        }

        It 'Usa try-catch em Export-VbaModule' {
            $pattern = 'function Export-VbaModule.*?try.*?catch'
            $exportContent -match $pattern | Should Be $true
        }

        It 'Libera objetos COM após uso' {
            $exportContent -match 'ReleaseComObject|Marshal.*Release' | Should Be $true
        }

        It 'Executa garbage collection após COM' {
            $exportContent -match 'GC.*Collect|garbage.*collection' | Should Be $true
        }

        It 'Fecha Word corretamente' {
            $exportContent -match 'word\.Quit\(\)|word\.Close' | Should Be $true
        }
    }

    Context 'Instalação - Gerenciamento de Erros' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
        }

        It 'Usa try-catch em Import-VbaModule' {
            $pattern = 'function Import-VbaModule.*?try.*?catch'
            $installContent -match $pattern | Should Be $true
        }

        It 'Libera objetos COM após uso' {
            $installContent -match 'ReleaseComObject|Marshal.*Release' | Should Be $true
        }

        It 'Fecha Word corretamente em caso de erro' {
            $installContent -match 'word\.Quit\(\)|word\.Close' | Should Be $true
        }

        It 'Registra erros no log' {
            $installContent -match 'Write-Log.*ERROR|Log.*erro' | Should Be $true
        }
    }

    Context 'Exportação - Logging' {
        
        BeforeAll {
            $exportContent = Get-Content $exportScript -Raw
        }

        It 'Deve ter função Write-Log' {
            $exportContent -match 'function Write-Log' | Should Be $true
        }

        It 'Deve registrar início da exportação' {
            $exportContent -match 'Exportando|Export.*start' | Should Be $true
        }

        It 'Deve registrar sucesso' {
            $exportContent -match 'SUCCESS|sucesso' | Should Be $true
        }

        It 'Deve registrar avisos' {
            $exportContent -match 'WARNING|aviso' | Should Be $true
        }

        It 'Deve registrar erros' {
            $exportContent -match 'ERROR|erro' | Should Be $true
        }
    }

    Context 'Instalação - Logging' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
        }

        It 'Deve ter função Write-Log' {
            $installContent -match 'function Write-Log' | Should Be $true
        }

        It 'Deve criar arquivo de log' {
            $installContent -match 'LogFile|log.*file|arquivo.*log' | Should Be $true
        }

        It 'Deve registrar cada etapa' {
            $installContent -match 'ETAPA|STEP|Stage' | Should Be $true
        }

        It 'Deve registrar conclusão' {
            $installContent -match 'CONCLU[IÍ]DA|COMPLETED|SUCCESS' | Should Be $true
        }
    }

    Context 'Consistência entre Export e Import' {
        
        BeforeAll {
            $exportContent = Get-Content $exportScript -Raw
            $installContent = Get-Content $installScript -Raw
        }

        It 'Pasta VBAModule é exportada e importada' {
            ($exportContent -match 'VBAModule') | Should Be $true
            ($installContent -match 'VBAModule') | Should Be $true
        }

        It 'Pasta RibbonCustomization é exportada e importada' {
            ($exportContent -match 'RibbonCustomization') | Should Be $true
            ($installContent -match 'RibbonCustomization') | Should Be $true
        }

        It 'Pasta OfficeCustomUI é exportada e importada' {
            ($exportContent -match 'OfficeCustomUI') | Should Be $true
            ($installContent -match 'OfficeCustomUI') | Should Be $true
        }

        It 'Arquivo monolithicMod.bas é exportado e importado' {
            ($exportContent -match 'monolithicMod') | Should Be $true
            ($installContent -match 'monolithicMod') | Should Be $true
        }
    }

    Context 'Segurança e Permissões' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
        }

        It 'Verifica permissões antes de escrever' {
            $installContent -match 'Test-Path.*Write|permiss.*escrita' | Should Be $true
        }

        It 'Não requer privilégios de administrador' {
            # Não deve usar comandos que requerem admin
            $installContent -match 'RunAs|AsAdministrator|Elevate' | Should Be $false
        }

        It 'Usa caminhos do perfil do usuário' {
            $installContent -match 'APPDATA|USERPROFILE|LOCALAPPDATA' | Should Be $true
        }
    }

    Context 'Interatividade e Confirmações' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
            $exportContent = Get-Content $exportScript -Raw
        }

        It 'Install permite modo Force (não interativo)' {
            $installContent -match '\$Force' | Should Be $true
        }

        It 'Install pede confirmação sem Force' {
            $installContent -match 'Read-Host|confirma' | Should Be $true
        }

        It 'Export pede confirmação para erros de compilação' {
            $exportContent -match 'Read-Host.*continuar|continue.*export' | Should Be $true
        }
    }
}
