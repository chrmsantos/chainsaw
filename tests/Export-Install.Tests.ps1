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

    Context 'Criação e Limpeza de Backups' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
        }

        It 'Cria backup completo antes da instalação' {
            $installContent -match 'Backup-CompleteConfiguration|full_backup' | Should Be $true
        }

        It 'Cria backup da pasta Templates' {
            $installContent -match 'Templates_backup' | Should Be $true
        }

        It 'Remove backups antigos mantendo os 5 mais recentes' {
            $installContent -match 'mantendo.*5|Select.*-First 5' | Should Be $true
        }

        It 'Backup contém manifesto JSON' {
            $installContent -match 'backup_manifest\.json' | Should Be $true
        }
    }

    Context 'Validação de VBA Module Export' {
        
        BeforeAll {
            $exportContent = Get-Content $exportScript -Raw
        }

        It 'Exporta especificamente o módulo monolithicMod' {
            $exportContent -match 'monolithicMod' | Should Be $true
        }

        It 'Salva módulo como .bas' {
            $exportContent -match '\.bas' | Should Be $true
        }

        It 'Verifica existência do projeto VBA antes de exportar' {
            $exportContent -match 'VBProject' | Should Be $true
        }

        It 'Usa COM automation para acessar Word' {
            $exportContent -match 'New-Object.*Word\.Application|ComObject' | Should Be $true
        }

        It 'Fecha Word corretamente após exportação' {
            $exportContent -match '\.Quit\(\)|ReleaseComObject' | Should Be $true
        }
    }

    Context 'Validação de VBA Module Import' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
        }

        It 'Função Import-VbaModule existe' {
            $installContent -match 'function Import-VbaModule' | Should Be $true
        }

        It 'Importa módulo do caminho VBAModule' {
            $installContent -match 'VBAModule' | Should Be $true
        }

        It 'Remove módulo existente antes de importar' {
            $installContent -match 'Remove.*component|VBComponents\.Remove' | Should Be $true
        }

        It 'Salva Normal.dotm após importação' {
            $installContent -match '\.Save\(\)|template\.Save' | Should Be $true
        }

        It 'Libera recursos COM após importação' {
            $installContent -match 'ReleaseComObject|GC\.Collect' | Should Be $true
        }
    }

    Context 'Validação de UI Customizations Export' {
        
        BeforeAll {
            $exportContent = Get-Content $exportScript -Raw
        }

        It 'Exporta arquivos .officeUI' {
            $exportContent -match '\.officeUI' | Should Be $true
        }

        It 'Procura em múltiplos locais possíveis' {
            $exportContent -match 'possiblePaths|@\(' | Should Be $true
        }

        It 'Exporta RibbonCustomization' {
            $exportContent -match 'Export-RibbonCustomization|RibbonCustomization' | Should Be $true
        }

        It 'Exporta OfficeCustomUI' {
            $exportContent -match 'Export-OfficeCustomUI|OfficeCustomUI' | Should Be $true
        }

        It 'Registra itens exportados em variável' {
            $exportContent -match '\$script:ExportedItems' | Should Be $true
        }
    }

    Context 'Validação de UI Customizations Import' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
        }

        It 'Importa RibbonCustomization' {
            $installContent -match 'Import-RibbonCustomization' | Should Be $true
        }

        It 'Importa OfficeCustomUI' {
            $installContent -match 'Import-OfficeCustomUI' | Should Be $true
        }

        It 'Copia para LOCALAPPDATA e APPDATA' {
            $installContent -match 'LOCALAPPDATA.*Office|APPDATA.*Office' | Should Be $true
        }

        It 'Cria diretórios de destino se não existirem' {
            $installContent -match 'New-Item.*Directory|mkdir' | Should Be $true
        }
    }

    Context 'Validação de Manifesto de Exportação' {
        
        BeforeAll {
            $exportContent = Get-Content $exportScript -Raw
        }

        It 'Cria manifesto com New-ExportManifest' {
            $exportContent -match 'New-ExportManifest' | Should Be $true
        }

        It 'Manifesto contém data de exportação' {
            $exportContent -match 'ExportDate|Get-Date' | Should Be $true
        }

        It 'Manifesto contém nome do usuário' {
            $exportContent -match 'UserName|env:USERNAME' | Should Be $true
        }

        It 'Manifesto contém nome do computador' {
            $exportContent -match 'ComputerName|env:COMPUTERNAME' | Should Be $true
        }

        It 'Manifesto é salvo como JSON' {
            $exportContent -match 'ConvertTo-Json|\.json' | Should Be $true
        }
    }

    Context 'Tratamento de Erros' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
            $exportContent = Get-Content $exportScript -Raw
        }

        It 'Install usa try-catch em operações críticas' {
            ($installContent -match 'try\s*\{').Count | Should BeGreaterThan 3
        }

        It 'Export usa try-catch em operações críticas' {
            ($exportContent -match 'try\s*\{').Count | Should BeGreaterThan 3
        }

        It 'Install registra erros em log' {
            $installContent -match 'catch.*Write-Log.*ERROR|LOG_LEVEL_ERROR' | Should Be $true
        }

        It 'Export registra erros em log' {
            $exportContent -match 'catch.*Write-Log.*ERROR|LOG_LEVEL_ERROR' | Should Be $true
        }

        It 'Install continua após falhas não críticas' {
            $installContent -match 'ErrorAction.*Continue|SilentlyContinue' | Should Be $true
        }
    }

    Context 'Logging e Auditoria' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
            $exportContent = Get-Content $exportScript -Raw
        }

        It 'Install cria arquivo de log' {
            $installContent -match 'LogFile|log.*\.log' | Should Be $true
        }

        It 'Export cria arquivo de log' {
            $exportContent -match 'LogFile|log.*\.log' | Should Be $true
        }

        It 'Logs contêm timestamps' {
            $installContent -match 'Get-Date.*Format|timestamp' | Should Be $true
            $exportContent -match 'Get-Date.*Format|timestamp' | Should Be $true
        }

        It 'Diferentes níveis de log são usados' {
            $installContent -match 'LOG_LEVEL_INFO|LOG_LEVEL_SUCCESS|LOG_LEVEL_WARNING|LOG_LEVEL_ERROR' | Should Be $true
        }
    }

    Context 'Verificação de Pré-requisitos' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
        }

        It 'Verifica se Word está em execução' {
            $installContent -match 'Test-WordRunning|Word.*running|Get-Process.*WINWORD' | Should Be $true
        }

        It 'Verifica versão do PowerShell' {
            $installContent -match 'PSVersion|PowerShell.*vers' | Should Be $true
        }

        It 'Verifica sistema operacional' {
            $installContent -match 'Windows|OS|Sistema operacional' | Should Be $true
        }

        It 'Verifica existência de arquivos de origem' {
            $installContent -match 'Test-Path.*stamp\.png|Test-Path.*Templates' | Should Be $true
        }
    }

    Context 'Detecção de Raiz do Projeto' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
        }

        It 'Detecta raiz do projeto automaticamente' {
            $installContent -match 'projectRoot|raiz.*projeto|Find.*Root' | Should Be $true
        }

        It 'Procura por marcadores de raiz (CHANGELOG, README)' {
            $installContent -match 'CHANGELOG|README|\.git' | Should Be $true
        }

        It 'Usa caminho relativo se raiz não for encontrada' {
            $installContent -match 'SourcePath|fallback' | Should Be $true
        }
    }

    Context 'Compilação VBA' {
        
        BeforeAll {
            $exportContent = Get-Content $exportScript -Raw
        }

        It 'Verifica compilação antes de exportar' {
            $exportContent -match 'Test-VbaModuleCompilation|Compila.*VBA' | Should Be $true
        }

        It 'Acessa módulos VBA para forçar compilação' {
            $exportContent -match 'VBComponents|CodeModule' | Should Be $true
        }

        It 'Reporta erros de compilação ao usuário' {
            $exportContent -match 'erro.*compila|compilation.*error' | Should Be $true
        }

        It 'Permite continuar mesmo com erros de compilação' {
            $exportContent -match 'Read-Host.*continuar|continue.*anyway' | Should Be $true
        }
    }

    Context 'Integração ETAPA 6' {
        
        BeforeAll {
            $installContent = Get-Content $installScript -Raw
        }

        It 'ETAPA 6 importa personalizações' {
            $installContent -match 'ETAPA 6.*Importa.*Personaliza|Personaliza.*Word' | Should Be $true
        }

        It 'Importa módulo VBA na ETAPA 6' {
            $installContent -match 'Import-VbaModule' | Should Be $true
        }

        It 'Importa Ribbon na ETAPA 6' {
            $installContent -match 'Import-RibbonCustomization' | Should Be $true
        }

        It 'Conta itens importados' {
            $installContent -match 'importedCount|\$importedCount' | Should Be $true
        }
    }
}
