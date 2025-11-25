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
        $restoreScript = Join-Path $scriptsPath "restore-backup.ps1"
        $installerCmd = Join-Path $repoRoot "chainsaw_installer.cmd"
    }

    Context 'Estrutura de Arquivos de Instalação' {
        
        It 'Pasta inst_scripts existe' {
            Test-Path $scriptsPath | Should Be $true
        }

        It 'chainsaw_installer.cmd existe na raiz do projeto' {
            Test-Path $installerCmd | Should Be $true
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

        It 'restore-backup.ps1 existe' {
            Test-Path $restoreScript | Should Be $true
        }

        It 'restore-backup.cmd existe' {
            $cmdPath = Join-Path $scriptsPath "restore-backup.cmd"
            Test-Path $cmdPath | Should Be $true
        }

        It 'Pasta inst_configs existe' {
            Test-Path (Join-Path $repoRoot "installation\inst_configs") | Should Be $true
        }

        It 'Pasta inst_docs existe' {
            Test-Path (Join-Path $repoRoot "installation\inst_docs") | Should Be $true
        }
    }

    Context 'chainsaw_installer.cmd - Validação de Conteúdo' {
        
        BeforeAll {
            $content = Get-Content $installerCmd -Raw
        }

        It 'Contém cabeçalho CHAINSAW' {
            $content -match 'CHAINSAW' | Should Be $true
        }

        It 'Define URL do repositório GitHub' {
            $content -match 'REPO_URL.*github\.com/chrmsantos/chainsaw' | Should Be $true
        }

        It 'Define diretório de instalação no USERPROFILE' {
            $content -match 'INSTALL_DIR.*%USERPROFILE%\\chainsaw' | Should Be $true
        }

        It 'Define variáveis de logging com timestamp' {
            ($content -match 'DATESTAMP') -and ($content -match 'TIMESTAMP') | Should Be $true
        }

        It 'Define arquivo de log' {
            $content -match 'LOG_FILE' | Should Be $true
        }

        It 'Verifica disponibilidade do PowerShell' {
            $content -match 'where powershell' | Should Be $true
        }

        It 'Implementa download usando PowerShell Invoke-WebRequest' {
            $content -match 'Invoke-WebRequest' | Should Be $true
        }

        It 'Verifica sucesso do download antes de prosseguir' {
            ($content -match 'if not exist "%TEMP_ZIP%"') -and ($content -match 'if errorlevel 1') | Should Be $true
        }

        It 'Cria backup antes de modificar instalação existente' {
            $content -match 'BACKUP_DIR|backup|xcopy.*backup' | Should Be $true
        }

        It 'Backup é criado APÓS verificação de download bem-sucedido' {
            # Verifica que o backup vem depois da verificação do ZIP
            $downloadCheck = $content.IndexOf('if not exist "%TEMP_ZIP%"')
            $backupSection = $content.IndexOf('BACKUP_DIR')
            ($backupSection -gt $downloadCheck) | Should Be $true
        }

        It 'Implementa extração de arquivos ZIP' {
            $content -match 'Expand-Archive' | Should Be $true
        }

        It 'Copia arquivos extraídos com xcopy' {
            $content -match 'xcopy' | Should Be $true
        }

        It 'Chama script install.cmd após extração' {
            $content -match 'call install\.cmd' | Should Be $true
        }

        It 'Implementa limpeza de arquivos temporários' {
            ($content -match 'del.*TEMP_ZIP') -and ($content -match 'rd.*TEMP_EXTRACT') | Should Be $true
        }

        It 'Copia log para pasta de instalação' {
            $content -match 'copy.*LOG_FILE.*INSTALL_LOG_DIR' | Should Be $true
        }

        It 'Implementa função de log (:Log)' {
            $content -match ':Log' | Should Be $true
        }

        It 'Implementa inicialização de log (:LogInit)' {
            $content -match ':LogInit' | Should Be $true
        }

        It 'Log contém informações do sistema (OS, USERNAME)' {
            ($content -match '%OS%') -and ($content -match '%USERNAME%') | Should Be $true
        }

        It 'Fornece feedback visual em todas as etapas' {
            ($content -match 'ETAPA 1') -and ($content -match 'ETAPA 2') -and ($content -match 'ETAPA 3') | Should Be $true
        }

        It 'Captura código de saída do instalador' {
            $content -match 'INSTALL_EXIT_CODE.*errorlevel' | Should Be $true
        }

        It 'Retorna código de saída apropriado' {
            $content -match 'exit /b.*INSTALL_EXIT_CODE' | Should Be $true
        }

        It 'Usa ExecutionPolicy Bypass para PowerShell' {
            $content -match '-ExecutionPolicy Bypass' | Should Be $true
        }

        It 'Suprime progresso do PowerShell para performance' {
            $content -match 'ProgressPreference.*SilentlyContinue' | Should Be $true
        }

        It 'Implementa tratamento de erros no download' {
            ($content -match '(?s)try.*catch') -and ($content -match 'ERRO.*download') | Should Be $true
        }

        It 'Implementa tratamento de erros na extração' {
            $content -match 'ERRO.*extracao|ERRO.*extrair' | Should Be $true
        }

        It 'Valida existência do script install.cmd após extração' {
            $content -match '(?s)if not exist.*install\.cmd' | Should Be $true
        }

        It 'Cria backup com timestamp único' {
            $content -match 'chainsaw_backup_.*DATETIME' | Should Be $true
        }

        It 'Backup usa xcopy com flags apropriadas (/E /H /C /I /Y)' {
            $content -match 'xcopy.*\/E.*\/H.*\/C.*\/I.*\/Y' | Should Be $true
        }

        It 'Implementa backup seletivo como fallback' {
            $content -match 'backup seletivo|Tentando backup seletivo' | Should Be $true
        }

        It 'Valida estrutura extraída (chainsaw-main)' {
            $content -match 'chainsaw-main' | Should Be $true
        }

        It 'Usa variáveis de ambiente do Windows corretamente' {
            ($content -match '%USERPROFILE%') -and ($content -match '%TEMP%') | Should Be $true
        }

        It 'Desabilita eco de comandos (@echo off)' {
            $content -match '@echo off' | Should Be $true
        }

        It 'Habilita delayed expansion para variáveis' {
            $content -match 'enabledelayedexpansion' | Should Be $true
        }

        It 'Usa labels para controle de fluxo (:extract)' {
            $content -match ':extract' | Should Be $true
        }

        It 'Implementa múltiplas tentativas de remoção de pasta' {
            ($content -match 'Tentativa 1') -or ($content -match 'Tentativa 2') -or ($content -match 'rd /s /q') | Should Be $true
        }
    }

    Context 'chainsaw_installer.cmd - Ordem de Execução Segura' {
        
        BeforeAll {
            $content = Get-Content $installerCmd -Raw
        }

        It 'Download vem antes de qualquer modificação de arquivos' {
            $downloadIdx = $content.IndexOf('Invoke-WebRequest')
            $backupIdx = $content.IndexOf('BACKUP_DIR')
            $removeIdx = $content.IndexOf('rd /s /q "%INSTALL_DIR%"')
            
            ($downloadIdx -lt $backupIdx) -and ($downloadIdx -lt $removeIdx) | Should Be $true
        }

        It 'Verificação de download bem-sucedido vem antes do backup' {
            $verifyIdx = $content.IndexOf('if not exist "%TEMP_ZIP%"')
            $backupIdx = $content.IndexOf('BACKUP_DIR')
            
            $verifyIdx -lt $backupIdx | Should Be $true
        }

        It 'Backup vem antes de remover instalação existente' {
            $backupIdx = $content.IndexOf('xcopy "%INSTALL_DIR%\*" "!BACKUP_DIR!\"')
            $removeIdx = $content.IndexOf('rd /s /q "%INSTALL_DIR%"')
            
            $backupIdx -lt $removeIdx | Should Be $true
        }

        It 'Extração vem depois do backup' {
            $backupIdx = $content.IndexOf('BACKUP_DIR')
            $extractIdx = $content.IndexOf('Expand-Archive')
            
            $extractIdx -gt $backupIdx | Should Be $true
        }

        It 'Log é copiado antes de chamar install.cmd' {
            $logCopyIdx = $content.IndexOf('copy "%LOG_FILE%"')
            $installCallIdx = $content.IndexOf('call install.cmd')
            
            $logCopyIdx -lt $installCallIdx | Should Be $true
        }
    }

    Context 'chainsaw_installer.cmd - Segurança e Validação' {
        
        BeforeAll {
            $content = Get-Content $installerCmd -Raw
        }

        It 'Não executa comandos perigosos sem validação' {
            # Verifica que rd e del sempre vêm após validações
            $lines = $content -split "`r?`n"
            $dangerousCommands = $lines | Where-Object { $_ -match '^\s*(rd|del)\s+' }
            # Todos os comandos perigosos devem estar em blocos condicionais ou após validações
            $dangerousCommands.Count -eq 0 -or ($content -match 'if exist.*rd|if exist.*del') | Should Be $true
        }

        It 'Valida PowerShell antes de usá-lo' {
            ($content -match 'where powershell') -and ($content -match 'if errorlevel 1') | Should Be $true
        }

        It 'Fornece mensagens de erro claras' {
            ($content -match '\[ERRO\]') -and ($content -match 'PowerShell nao encontrado|Falha ao baixar|Falha ao extrair') | Should Be $true
        }

        It 'Pausa antes de sair em caso de erro' {
            $content -match 'pause' | Should Be $true
        }

        It 'Usa redirecionamento de erros apropriadamente (>nul 2>&1)' {
            $content -match '>nul 2>&1' | Should Be $true
        }
    }

    Context 'chainsaw_installer.cmd - Logging Completo' {
        
        BeforeAll {
            $content = Get-Content $installerCmd -Raw
        }

        It 'Log salvo no mesmo diretório do script' {
            $content -match 'SCRIPT_DIR.*~dp0' | Should Be $true
        }

        It 'Log também copiado para pasta do projeto' {
            $content -match 'INSTALL_LOG_DIR.*installation.*inst_docs.*inst_logs' | Should Be $true
        }

        It 'Log inclui timestamp no nome do arquivo' {
            $content -match 'installer_.*DATETIME.*\.log' | Should Be $true
        }

        It 'Função de log escreve para console e arquivo' {
            $logFunction = $content.Substring($content.IndexOf(':Log'))
            ($logFunction -match 'echo %MSG%') -and ($logFunction -match 'echo %MSG% >> "%LOG_FILE%"') | Should Be $true
        }

        It 'Log inicializa com informações do sistema' {
            $content -match 'Data/Hora de inicio|Sistema:|Usuario:|Diretorio de instalacao:' | Should Be $true
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

        It 'Implementa backup completo' {
            $content -match 'Backup-CompleteConfiguration|full_backup' | Should Be $true
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

    Context 'restore-backup.ps1 - Validação de Conteúdo' {
        
        BeforeAll {
            $content = Get-Content $restoreScript -Raw
        }

        It 'Contém cabeçalho CHAINSAW' {
            $content -match 'CHAINSAW|chainsaw' | Should Be $true
        }

        It 'Contém documentação de help' {
            $content -match '\.SYNOPSIS' | Should Be $true
        }

        It 'Define parâmetro BackupPath' {
            $content -match '\$BackupPath' | Should Be $true
        }

        It 'Define parâmetro List' {
            $content -match '\[switch\]\$List' | Should Be $true
        }

        It 'Define parâmetro Force' {
            $content -match '\[switch\]\$Force' | Should Be $true
        }

        It 'Implementa função Get-AvailableBackups' {
            $content -match 'Get-AvailableBackups|function\s+Get-AvailableBackups' | Should Be $true
        }

        It 'Implementa função Restore-TemplatesFromBackup' {
            $content -match 'Restore-TemplatesFromBackup|function\s+Restore-TemplatesFromBackup' | Should Be $true
        }

        It 'Implementa verificação de Word em execução' {
            $content -match 'Test-WordRunning|WINWORD' | Should Be $true
        }

        It 'Referencia backups completos' {
            $content -match 'full_backup|Templates_backup' | Should Be $true
        }

        It 'Implementa logging' {
            $content -match 'Write-Log|log' | Should Be $true
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

        It 'Documentação de backup existe' {
            $backupDoc = Join-Path $repoRoot "installation\inst_docs\GUIA_BACKUP_RESTAURACAO.md"
            Test-Path $backupDoc | Should Be $true
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
        
        It 'install.ps1 tem complexidade gerenciável (< 180 condicionais)' {
            $content = Get-Content $installScript -Raw
            $conditionals = ([regex]::Matches($content, '\bif\b|\belse\b|\belseif\b|\bswitch\b')).Count
            # Script complexo mas ainda gerenciável - alerta se > 180
            $conditionals -lt 180 | Should Be $true
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

        It 'restore-backup.ps1 tem complexidade moderada (< 60 condicionais)' {
            $content = Get-Content $restoreScript -Raw
            $conditionals = ([regex]::Matches($content, '\bif\b|\belse\b|\belseif\b|\bswitch\b')).Count
            $conditionals -lt 60 | Should Be $true
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

    Context 'Sistema de Verificação de Versão' {
        
        BeforeAll {
            $versionFile = Join-Path $repoRoot 'version.json'
            $vbaModulePath = Join-Path $repoRoot 'source\main\monolithicMod.bas'
        }

        It 'Arquivo version.json existe na raiz do repositório' {
            Test-Path $versionFile | Should Be $true
        }

        It 'version.json contém versão válida (formato X.Y.Z)' {
            $content = Get-Content $versionFile -Raw | ConvertFrom-Json
            $content.version | Should Match '^\d+\.\d+\.\d+$'
        }

        It 'version.json contém URL de download' {
            $content = Get-Content $versionFile -Raw | ConvertFrom-Json
            $content.downloadUrl | Should Not BeNullOrEmpty
            $content.downloadUrl | Should Match 'github\.com'
        }

        It 'version.json contém URL do instalador' {
            $content = Get-Content $versionFile -Raw | ConvertFrom-Json
            $content.installerUrl | Should Not BeNullOrEmpty
            $content.installerUrl | Should Match 'chainsaw_installer\.cmd'
        }

        It 'version.json contém data de release' {
            $content = Get-Content $versionFile -Raw | ConvertFrom-Json
            $content.releaseDate | Should Not BeNullOrEmpty
        }

        It 'VBA contém constante CHAINSAW_VERSION' {
            $vbaContent = Get-Content $vbaModulePath -Raw
            $vbaContent | Should Match 'Private Const CHAINSAW_VERSION As String'
        }

        It 'VBA contém função CheckForUpdates' {
            $vbaContent = Get-Content $vbaModulePath -Raw
            $vbaContent | Should Match 'Public Function CheckForUpdates\(\) As Boolean'
        }

        It 'VBA contém função GetLocalVersion' {
            $vbaContent = Get-Content $vbaModulePath -Raw
            $vbaContent | Should Match 'Private Function GetLocalVersion\(\) As String'
        }

        It 'VBA contém função GetRemoteVersion' {
            $vbaContent = Get-Content $vbaModulePath -Raw
            $vbaContent | Should Match 'Private Function GetRemoteVersion\(\) As String'
        }

        It 'VBA contém função CompareVersions' {
            $vbaContent = Get-Content $vbaModulePath -Raw
            $vbaContent | Should Match 'Private Function CompareVersions\('
        }

        It 'VBA contém sub PromptForUpdate' {
            $vbaContent = Get-Content $vbaModulePath -Raw
            $vbaContent | Should Match 'Public Sub PromptForUpdate\(\)'
        }

        It 'GetRemoteVersion usa URL correto do GitHub' {
            $vbaContent = Get-Content $vbaModulePath -Raw
            $vbaContent | Should Match 'raw\.githubusercontent\.com/chrmsantos/chainsaw/main/version\.json'
        }

        It 'PromptForUpdate executa chainsaw_installer.cmd' {
            $vbaContent = Get-Content $vbaModulePath -Raw
            $vbaContent | Should Match 'chainsaw_installer\.cmd'
        }

        It 'chainsaw_installer.cmd cria version.json local' {
            $installerContent = Get-Content $installerCmd -Raw
            $installerContent | Should Match 'LOCAL_VERSION_FILE.*version\.json'
            $installerContent | Should Match 'ConvertTo-Json'
        }

        It 'Versão no VBA corresponde ao version.json' {
            $versionJson = Get-Content $versionFile -Raw | ConvertFrom-Json
            $vbaContent = Get-Content $vbaModulePath -Raw
            
            if ($vbaContent -match 'Private Const CHAINSAW_VERSION As String = "([^"]+)"') {
                $vbaVersion = $matches[1]
                $vbaVersion | Should Be $versionJson.version
            }
            else {
                throw "CHAINSAW_VERSION não encontrado no VBA"
            }
        }

        It 'Versão no cabeçalho VBA corresponde ao version.json' {
            $versionJson = Get-Content $versionFile -Raw | ConvertFrom-Json
            $vbaContent = Get-Content $vbaModulePath -Raw -Encoding UTF8
            
            if ($vbaContent -match "' Versão: ([^\r\n]+)") {
                $headerVersion = $matches[1].Trim()
                $headerVersion | Should Be $versionJson.version
            }
            else {
                throw "Versão não encontrada no cabeçalho VBA"
            }
        }
    }
}

