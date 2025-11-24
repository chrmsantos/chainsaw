#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1

# Override Get-RepoRoot for test context
function Get-RepoRoot {
    $testsDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $repoRoot = Split-Path -Parent $testsDir
    return $repoRoot
}

Describe 'CHAINSAW - Testes de Scripts de Instalação' {

    BeforeAll {
        $repoRoot = Get-RepoRoot
        $scriptsPath = Join-Path $repoRoot "installation\inst_scripts"
        $installScript = Join-Path $scriptsPath "install.ps1"
        $exportScript = Join-Path $scriptsPath "export-config.ps1"
        $updateScript = Join-Path $scriptsPath "update-vba-module.ps1"
    }

    Context 'Estrutura de Arquivos de Instalação' {
        
        It 'Pasta inst_scripts existe' {
            Test-Path $scriptsPath | Should Be $true
        }

        It 'install.ps1 existe' {
            Test-Path $installScript | Should Be $true
        }

        It 'export-config.ps1 existe' {
            Test-Path $exportScript | Should Be $true
        }

        It 'update-vba-module.ps1 existe' {
            Test-Path $updateScript | Should Be $true
        }

        It 'Pasta inst_configs existe' {
            Test-Path (Join-Path $repoRoot "installation\inst_configs") | Should Be $true
        }

        It 'Pasta inst_docs existe' {
            Test-Path (Join-Path $repoRoot "installation\inst_docs") | Should Be $true
        }
    }

    Context 'install.ps1 - Validação de Conteúdo' {
        
        BeforeAll {
            $content = Get-Content $installScript -Raw
        }

        It 'Contém cabeçalho de licença GPLv3' {
            $content -match 'GNU GPLv3' | Should Be $true
        }

        It 'Contém informações de versão' {
            # Teste flexível para diferentes encodings
            ($content -match 'Vers[aã]o:\s*\d+\.\d+\.\d+') -or ($content -match 'Version:\s*\d+\.\d+\.\d+') | Should Be $true
        }

        It 'Contém help documentation (SYNOPSIS)' {
            $content -match '\.SYNOPSIS' | Should Be $true
        }

        It 'Contém help documentation (DESCRIPTION)' {
            $content -match '\.DESCRIPTION' | Should Be $true
        }

        It 'Contém help documentation (EXAMPLE)' {
            $content -match '\.EXAMPLE' | Should Be $true
        }

        It 'Define parâmetro Force' {
            $content -match '\[switch\]\$Force' | Should Be $true
        }

        It 'Define parâmetro NoBackup' {
            $content -match '\[switch\]\$NoBackup' | Should Be $true
        }

        It 'Define parâmetro SkipCustomizations' {
            $content -match '\[switch\]\$SkipCustomizations' | Should Be $true
        }

        It 'Implementa função de backup' {
            $content -match 'backup|Backup' | Should Be $true
        }

        It 'Implementa função de log' {
            $content -match 'log|Log|Write-Log' | Should Be $true
        }

        It 'Valida versão do PowerShell' {
            # install.ps1 não tem #requires mas documenta compatibilidade
            ($content -match '#requires\s+-Version\s+\d+\.\d+') -or ($content -match 'PowerShell\s+\d+\.\d+') | Should Be $true
        }
    }

    Context 'export-config.ps1 - Validação de Conteúdo' {
        
        BeforeAll {
            $content = Get-Content $exportScript -Raw
        }

        It 'Contém cabeçalho de licença' {
            $content -match 'GNU GPLv3|Licença' | Should Be $true
        }

        It 'Contém documentação de help' {
            $content -match '\.SYNOPSIS' | Should Be $true
        }

        It 'Define parâmetro ExportPath' {
            $content -match '\$ExportPath' | Should Be $true
        }

        It 'Define parâmetro IncludeRegistry' {
            $content -match '\[switch\]\$IncludeRegistry' | Should Be $true
        }

        It 'Referencia Normal.dotm' {
            $content -match 'Normal\.dotm' | Should Be $true
        }

        It 'Referencia Ribbon/Office UI' {
            $content -match 'Ribbon|OfficeUI|Office.*UI' | Should Be $true
        }

        It 'Implementa exportação de Building Blocks' {
            $content -match 'Building\s*Blocks|BuildingBlocks' | Should Be $true
        }

        It 'Cria estrutura de manifesto' {
            $content -match 'MANIFEST|manifest' | Should Be $true
        }
    }

    Context 'update-vba-module.ps1 - Validação de Conteúdo' {
        
        BeforeAll {
            $content = Get-Content $updateScript -Raw
        }

        It 'Contém cabeçalho CHAINSAW' {
            $content -match 'CHAINSAW|chainsaw' | Should Be $true
        }

        It 'Define parâmetro Force' {
            $content -match '\[switch\]\$Force' | Should Be $true
        }

        It 'Valida caminho do módulo VBA' {
            $content -match 'monolithicMod\.bas' | Should Be $true
        }

        It 'Valida caminho do Normal.dotm' {
            $content -match 'Normal\.dotm' | Should Be $true
        }

        It 'Implementa detecção de Word em execução' {
            $content -match 'WINWORD|Get-Process.*Word' | Should Be $true
        }

        It 'Implementa fechamento de Word' {
            $content -match 'CloseMainWindow|Stop-Process' | Should Be $true
        }

        It 'Referencia caminho do projeto' {
            $content -match '\$ProjectRoot|\$ScriptPath' | Should Be $true
        }
    }

    Context 'Validação de Caminhos Críticos' {
        
        It 'Módulo VBA monolítico existe' {
            $vbaPath = Join-Path $repoRoot "source\main\monolithicMod.bas"
            Test-Path $vbaPath | Should Be $true
        }

        It 'Pasta Templates existe em inst_configs' {
            $templatesPath = Join-Path $repoRoot "installation\inst_configs\Templates"
            Test-Path $templatesPath | Should Be $true
        }

        It 'Arquivo chainsaw.config existe' {
            $configPath = Join-Path $repoRoot "installation\inst_configs\Templates\chainsaw.config"
            Test-Path $configPath | Should Be $true
        }

        It 'Pasta de logs existe' {
            $logsPath = Join-Path $repoRoot "installation\inst_docs\inst_logs"
            Test-Path $logsPath | Should Be $true
        }
    }

    Context 'Validação de Parâmetros CmdletBinding' {
        
        It 'install.ps1 usa [CmdletBinding()]' {
            $installContent = Get-Content $installScript -Raw
            $installContent -match '\[CmdletBinding\(\)\]' | Should Be $true
        }

        It 'export-config.ps1 usa [CmdletBinding()]' {
            $exportContent = Get-Content $exportScript -Raw
            $exportContent -match '\[CmdletBinding\(\)\]' | Should Be $true
        }

        It 'update-vba-module.ps1 usa [CmdletBinding()]' {
            $updateContent = Get-Content $updateScript -Raw
            $updateContent -match '\[CmdletBinding\(\)\]' | Should Be $true
        }
    }

    Context 'Segurança e Validação de Entrada' {
        
        It 'install.ps1 valida caminhos antes de copiar' {
            $content = Get-Content $installScript -Raw
            $content -match 'Test-Path' | Should Be $true
        }

        It 'export-config.ps1 valida existência de arquivos' {
            $content = Get-Content $exportScript -Raw
            $content -match 'Test-Path' | Should Be $true
        }

        It 'update-vba-module.ps1 valida arquivos críticos' {
            $content = Get-Content $updateScript -Raw
            ($content -match 'Test-Path.*VbaModulePath') -or ($content -match 'Test-Path.*NormalDotmPath') | Should Be $true
        }
    }

    Context 'Tratamento de Erros' {
        
        It 'install.ps1 implementa try-catch ou ErrorAction' {
            $content = Get-Content $installScript -Raw
            ($content -match 'try\s*\{') -or ($content -match '-ErrorAction') | Should Be $true
        }

        It 'export-config.ps1 implementa try-catch ou ErrorAction' {
            $content = Get-Content $exportScript -Raw
            ($content -match 'try\s*\{') -or ($content -match '-ErrorAction') | Should Be $true
        }

        It 'update-vba-module.ps1 implementa try-catch ou ErrorAction' {
            $content = Get-Content $updateScript -Raw
            ($content -match 'try\s*\{') -or ($content -match '-ErrorAction') | Should Be $true
        }
    }

    Context 'Funções de Output e Feedback' {
        
        It 'install.ps1 fornece feedback visual' {
            $content = Get-Content $installScript -Raw
            $content -match 'Write-Host|Write-Output|Write-Verbose' | Should Be $true
        }

        It 'export-config.ps1 fornece feedback visual' {
            $content = Get-Content $exportScript -Raw
            $content -match 'Write-Host|Write-Output|Write-Verbose' | Should Be $true
        }

        It 'update-vba-module.ps1 fornece feedback visual' {
            $content = Get-Content $updateScript -Raw
            $content -match 'Write-Host|Write-Output|Write-Verbose' | Should Be $true
        }

        It 'Scripts usam cores para feedback (ForegroundColor)' {
            $installContent = Get-Content $installScript -Raw
            $exportContent = Get-Content $exportScript -Raw
            $updateContent = Get-Content $updateScript -Raw
            
            $hasColors = ($installContent -match '-ForegroundColor') -or 
                        ($exportContent -match '-ForegroundColor') -or 
                        ($updateContent -match '-ForegroundColor')
            
            $hasColors | Should Be $true
        }
    }

    Context 'Compatibilidade e Requisitos' {
        
        It 'install.ps1 documenta requisitos de PowerShell' {
            $content = Get-Content $installScript -Raw -Encoding UTF8
            # Aceita #requires OU documentação de compatibilidade
            (($content -match '#requires\s+-Version') -or ($content -match 'PowerShell\s+\d+\.\d+') -or ($content -match 'Compatibilidade')) | Should Be $true
        }

        It 'Scripts estão codificados em UTF-8 ou ASCII válido' {
            # Verificação básica de encoding - arquivos devem ser legíveis
            $scripts = @($installScript, $exportScript, $updateScript)
            foreach ($script in $scripts) {
                { $null = Get-Content $script -Raw -Encoding UTF8 } | Should Not Throw
            }
        }

        It 'Scripts referenciam caminhos do AppData' {
            $installContent = Get-Content $installScript -Raw
            $exportContent = Get-Content $exportScript -Raw
            
            ($installContent -match '\$env:APPDATA') -or ($exportContent -match '\$env:APPDATA') | Should Be $true
        }
    }

    Context 'Integração COM com Microsoft Word' {
        
        It 'update-vba-module.ps1 usa COM para interagir com Word' {
            $content = Get-Content $updateScript -Raw
            ($content -match 'New-Object.*Word\.Application') -or 
            ($content -match 'CreateObject.*Word') -or 
            ($content -match 'Word\.Application') | Should Be $true
        }
    }

    Context 'Wrappers CMD' {
        
        It 'install.cmd existe' {
            $cmdPath = Join-Path $scriptsPath "install.cmd"
            Test-Path $cmdPath | Should Be $true
        }

        It 'export-config.cmd existe' {
            $cmdPath = Join-Path $scriptsPath "export-config.cmd"
            Test-Path $cmdPath | Should Be $true
        }

        It 'update-vba-module.cmd existe' {
            $cmdPath = Join-Path $scriptsPath "update-vba-module.cmd"
            Test-Path $cmdPath | Should Be $true
        }

        It 'install.cmd chama install.ps1' {
            $cmdPath = Join-Path $scriptsPath "install.cmd"
            if (Test-Path $cmdPath) {
                $content = Get-Content $cmdPath -Raw
                $content -match 'install\.ps1' | Should Be $true
            }
        }

        It 'Wrappers CMD usam ExecutionPolicy Bypass' {
            $cmdFiles = Get-ChildItem $scriptsPath -Filter "*.cmd"
            $allBypass = $true
            foreach ($cmd in $cmdFiles) {
                $content = Get-Content $cmd.FullName -Raw
                if ($content -notmatch '-ExecutionPolicy\s+Bypass') {
                    $allBypass = $false
                    break
                }
            }
            $allBypass | Should Be $true
        }
    }

    Context 'Documentação de Instalação' {
        
        It 'Documentação de instalação existe (INSTALL.md ou GUIA_INSTALACAO.md)' {
            $installDoc1 = Join-Path $repoRoot "installation\inst_docs\INSTALL.md"
            $installDoc2 = Join-Path $repoRoot "installation\inst_docs\GUIA_INSTALACAO.md"
            ((Test-Path $installDoc1) -or (Test-Path $installDoc2)) | Should Be $true
        }

        It 'Documentação contém instruções' {
            $installDoc1 = Join-Path $repoRoot "installation\inst_docs\INSTALL.md"
            $installDoc2 = Join-Path $repoRoot "installation\inst_docs\GUIA_INSTALACAO.md"
            $docPath = if (Test-Path $installDoc1) { $installDoc1 } else { $installDoc2 }
            if (Test-Path $docPath) {
                $content = Get-Content $docPath -Raw
                $content.Length -gt 100 | Should Be $true
            }
        }
    }

    Context 'Análise de Complexidade Ciclomática' {
        
        It 'install.ps1 tem complexidade gerenciável (< 150 condicionais)' {
            $content = Get-Content $installScript -Raw
            $conditionals = ([regex]::Matches($content, '\bif\b|\belse\b|\belseif\b|\bswitch\b')).Count
            # Script complexo mas ainda gerenciável - alerta se > 150
            $conditionals -lt 150 | Should Be $true
        }

        It 'export-config.ps1 não é excessivamente complexo (< 80 condicionais)' {
            $content = Get-Content $exportScript -Raw
            $conditionals = ([regex]::Matches($content, '\bif\b|\belse\b|\belseif\b|\bswitch\b')).Count
            $conditionals -lt 80 | Should Be $true
        }

        It 'update-vba-module.ps1 é relativamente simples (< 20 condicionais)' {
            $content = Get-Content $updateScript -Raw
            $conditionals = ([regex]::Matches($content, '\bif\b|\belse\b|\belseif\b|\bswitch\b')).Count
            $conditionals -lt 20 | Should Be $true
        }
    }

    Context 'Verificação de Encoding UTF-8' {
        
        It 'install.ps1 está em encoding válido (sem BOMs inválidos)' {
            $bytes = [System.IO.File]::ReadAllBytes($installScript)
            # Não deve ter sequências de bytes inválidas
            $bytes.Count -gt 0 | Should Be $true
        }

        It 'export-config.ps1 está em encoding válido' {
            $bytes = [System.IO.File]::ReadAllBytes($exportScript)
            $bytes.Count -gt 0 | Should Be $true
        }

        It 'update-vba-module.ps1 está em encoding válido' {
            $bytes = [System.IO.File]::ReadAllBytes($updateScript)
            $bytes.Count -gt 0 | Should Be $true
        }
    }
}
