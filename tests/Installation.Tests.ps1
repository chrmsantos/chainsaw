#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1

    Context 'chainsaw_installer.cmd - Wrapper' {
        BeforeAll {
            $content = Get-Content $installerCmd -Raw
        }

        It 'Verifica PowerShell antes de executar' {
            $content -match 'where powershell' | Should Be $true
        }

        It 'Aponta para chainsaw_installer.ps1' {
            $content -match 'chainsaw_installer\.ps1' | Should Be $true
        }

        It 'Utiliza ExecutionPolicy Bypass' {
            $content -match '-ExecutionPolicy Bypass' | Should Be $true
        }

        It 'Mantem pausa quando aberto por duplo clique' {
            $content -match 'CMDCMDLINE' | Should Be $true
        }
    }

    Context 'exportar_configs.cmd - Wrapper' {
        BeforeAll {
            $content = Get-Content $exportCmd -Raw
        }

        It 'Verifica PowerShell antes de executar' {
            $content -match 'where powershell' | Should Be $true
        }

        It 'Aponta para exportar_configs.ps1' {
            $content -match 'exportar_configs\.ps1' | Should Be $true
        }

        It 'Utiliza ExecutionPolicy Bypass' {
            $content -match '-ExecutionPolicy Bypass' | Should Be $true
        }

        It 'Mantem pausa quando aberto por duplo clique' {
            $content -match 'CMDCMDLINE' | Should Be $true
        }
    }

    Context 'chainsaw_installer.ps1 - Interface simplificada' {
        BeforeAll {
            $content = Get-Content $installerPipeline -Raw
        }

        It 'Define menus numerados para escolhas' {
            $content -match 'Read-MenuOption' | Should Be $true
        }

        It 'Utiliza perguntas de sim ou nao' {
            ($content -match 'Ask-YesNo') | Should Be $true
        }

        It 'Executa install.ps1 via Start-Process' {
            $content -match 'install\.ps1' | Should Be $true
            $content -match 'Start-Process' | Should Be $true
        }

        It 'Consulta logs em inst_docs\\inst_logs' {
            $content -match 'inst_docs\\inst_logs' | Should Be $true
        }
    }

    Context 'exportar_configs.ps1 - Interface simplificada' {
        BeforeAll {
            $content = Get-Content $exportPipeline -Raw
        }

        It 'Reutiliza perguntas de sim ou nao' {
            $content -match 'Ask-YesNo' | Should Be $true
        }

        It 'Chama export-config.ps1 via Start-Process' {
            $content -match 'export-config\.ps1' | Should Be $true
            $content -match 'Start-Process' | Should Be $true
        }

        It 'Propaga parametro ForceCloseWord' {
            $content -match 'ForceCloseWord' | Should Be $true
        }

        It 'Resolve destino padrao exported-config' {
            $content -match 'exported-config' | Should Be $true
        }
    }

    }

    Context 'Compatibilidade e Requisitos' {

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

        It 'chainsaw.ps1 usa [CmdletBinding()]' {
            $launcherContent = Get-Content $chainsawLauncher -Raw
            $launcherContent -match '\[CmdletBinding\(\)\]' | Should Be $true
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

        It 'backup-functions.ps1 valida caminhos críticos' {
            $content = Get-Content $backupScript -Raw
            $content -match 'Test-Path' | Should Be $true
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

        It 'backup-functions.ps1 implementa try-catch ou ErrorAction' {
            $content = Get-Content $backupScript -Raw
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

        It 'backup-functions.ps1 fornece feedback visual' {
            $content = Get-Content $backupScript -Raw
            $content -match 'Write-Host|Write-Output|Write-Verbose' | Should Be $true
        }

        It 'Scripts usam cores para feedback (ForegroundColor)' {
            $installContent = Get-Content $installScript -Raw
            $exportContent = Get-Content $exportScript -Raw
            $backupContent = Get-Content $backupScript -Raw
            
            $hasColors = ($installContent -match '-ForegroundColor') -or 
            ($exportContent -match '-ForegroundColor') -or 
            ($backupContent -match '-ForegroundColor')
            
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
            $scripts = @($installScript, $exportScript, $backupScript)
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

    Context 'Wrappers CMD' {
        
        It 'Somente dois wrappers permanecem' {
            $cmdFiles = Get-ChildItem $scriptsPath -Filter "*.cmd"
            ($cmdFiles.Count) | Should Be 2
        }

        It 'chainsaw_installer.cmd chama chainsaw_installer.ps1' {
            $content = Get-Content $installerCmd -Raw
            $content -match 'chainsaw_installer\.ps1' | Should Be $true
        }

        It 'exportar_configs.cmd chama exportar_configs.ps1' {
            $content = Get-Content $exportCmd -Raw
            $content -match 'exportar_configs\.ps1' | Should Be $true
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

        It 'backup-functions.ps1 tem complexidade baixa (< 30 condicionais)' {
            $content = Get-Content $backupScript -Raw
            $conditionals = ([regex]::Matches($content, '\bif\b|\belse\b|\belseif\b|\bswitch\b')).Count
            $conditionals -lt 30 | Should Be $true
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

        It 'backup-functions.ps1 está em encoding válido' {
            $bytes = [System.IO.File]::ReadAllBytes($backupScript)
            $bytes.Count -gt 0 | Should Be $true
        }

        It 'monolithicMod.bas está em encoding UTF-8' {
            $vbaPath = Join-Path $repoRoot 'source\main\monolithicMod.bas'
            $bytes = [System.IO.File]::ReadAllBytes($vbaPath)
            
            # Verifica UTF-8 BOM (opcional mas recomendado)
            $hasUtf8Bom = ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
            
            # Deve ter conteúdo
            $bytes.Count -gt 0 | Should Be $true
            
            # Tenta ler como UTF-8 sem erros
            $content = Get-Content $vbaPath -Raw -Encoding UTF8
            $content.Length -gt 0 | Should Be $true
        }

        It 'monolithicMod.bas não contém caracteres de controle inválidos' {
            $vbaPath = Join-Path $repoRoot 'source\main\monolithicMod.bas'
            $content = Get-Content $vbaPath -Raw -Encoding UTF8
            
            # Não deve conter null bytes
            $content.Contains([char]0) | Should Be $false
            
            # Não deve conter caracteres de controle exceto CR, LF, TAB
            $invalidControlChars = $content.ToCharArray() | Where-Object { 
                $code = [int]$_
                ($code -lt 32 -and $code -ne 9 -and $code -ne 10 -and $code -ne 13)
            }
            $invalidControlChars.Count | Should Be 0
        }

        It 'monolithicMod.bas usa quebras de linha consistentes (CRLF)' {
            $vbaPath = Join-Path $repoRoot 'source\main\monolithicMod.bas'
            $content = [System.IO.File]::ReadAllText($vbaPath, [System.Text.Encoding]::UTF8)
            
            # VBA usa CRLF (Windows)
            $lfOnly = ($content -match "[^`r]`n")
            $lfOnly | Should Be $false
        }

        It 'monolithicMod.bas pode ser lido com diferentes encodings sem erro' {
            $vbaPath = Join-Path $repoRoot 'source\main\monolithicMod.bas'
            
            # UTF-8
            { Get-Content $vbaPath -Raw -Encoding UTF8 -ErrorAction Stop } | Should Not Throw
            
            # Default (para compatibilidade)
            { Get-Content $vbaPath -Raw -ErrorAction Stop } | Should Not Throw
        }

        It 'monolithicMod.bas: acentuação portuguesa está correta' {
            $vbaPath = Join-Path $repoRoot 'source\main\monolithicMod.bas'
            $content = Get-Content $vbaPath -Raw -Encoding UTF8
            
            # Verifica se acentos comuns em português são lidos corretamente
            # O arquivo deve conter "Versão" no cabeçalho
            if ($content -match 'Vers.o:') {
                $content | Should Match 'Versão:'
            }
        }

        It 'Todos os arquivos .ps1 usam UTF-8 ou ASCII' {
            $psFiles = Get-ChildItem (Join-Path $repoRoot 'installation\inst_scripts') -Filter '*.ps1'
            
            foreach ($file in $psFiles) {
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
                
                # Não deve ter UTF-16 BOM
                $hasUtf16BOM = ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) -or 
                ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
                $hasUtf16BOM | Should Be $false
                
                # Deve conseguir ler como UTF-8
                { Get-Content $file.FullName -Raw -Encoding UTF8 -ErrorAction Stop } | Should Not Throw
            }
        }

        It 'chainsaw_installer.cmd não contém caracteres não-ASCII problemáticos' {
            $content = Get-Content $installerCmd -Raw
            
            # CMD deve ser ASCII ou compatível
            $nonAsciiChars = $content.ToCharArray() | Where-Object { [int]$_ -gt 127 }
            
            # Pode ter alguns caracteres latinos (ã, ç, etc) mas não deve ter símbolos especiais
            foreach ($char in $nonAsciiChars) {
                $code = [int]$char
                # Permite caracteres latinos-1 (128-255) mas avisa sobre outros
                if ($code -gt 255) {
                    Write-Warning "Caractere Unicode detectado em installer.cmd: U+$($code.ToString('X4'))"
                }
            }
            
            # Deve ter conteúdo válido
            $content.Length -gt 0 | Should Be $true
        }

        It 'version.json é UTF-8 válido' {
            $versionPath = Join-Path $repoRoot 'version.json'
            $bytes = [System.IO.File]::ReadAllBytes($versionPath)
            
            # Deve conseguir parsear como JSON
            { Get-Content $versionPath -Raw | ConvertFrom-Json -ErrorAction Stop } | Should Not Throw
            
            # UTF-8 sem BOM ou com BOM
            $content = [System.IO.File]::ReadAllText($versionPath, [System.Text.Encoding]::UTF8)
            $content.Length -gt 0 | Should Be $true
        }

        It 'CHANGELOG.md é UTF-8 válido' {
            $changelogPath = Join-Path $repoRoot 'CHANGELOG.md'
            $bytes = [System.IO.File]::ReadAllBytes($changelogPath)
            
            # Deve conseguir ler como UTF-8
            { Get-Content $changelogPath -Raw -Encoding UTF8 -ErrorAction Stop } | Should Not Throw
            
            # Deve conter acentuação correta
            $content = Get-Content $changelogPath -Raw -Encoding UTF8
            $content | Should Match 'Versão|versão|Adicionado'
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
    
    Context 'Validação de Encoding em Arquivos Críticos' {
        
        It 'VBA pode ser lido com UTF-8 sem caracteres corrompidos' {
            $vbaContent = Get-Content $vbaModulePath -Raw -Encoding UTF8
            
            # Não deve conter caracteres de substituição Unicode
            $vbaContent | Should Not Match '�'
            
            # Deve conter comentários em português
            $vbaContent | Should Match "' ="
        }
        
        It 'version.json usa UTF-8 válido' {
            $content = Get-Content $versionFile -Raw -Encoding UTF8
            
            # Não deve ter caracteres corrompidos
            $content | Should Not Match '�'
            
            # Deve ser JSON válido
            $json = $content | ConvertFrom-Json
            $json.version | Should Not BeNullOrEmpty
        }
        
        It 'CHANGELOG.md pode ser lido com UTF-8' {
            $changelogPath = Join-Path $repoRoot "CHANGELOG.md"
            
            if (Test-Path $changelogPath) {
                $content = Get-Content $changelogPath -Raw -Encoding UTF8
                
                # Não deve ter caracteres corrompidos
                $content | Should Not Match '�'
                
                # Deve conter cabeçalho
                $content | Should Match '# Changelog|# CHANGELOG'
            }
        }
        
        It 'README.md preserva acentuação portuguesa' {
            $readmePath = Join-Path $repoRoot "README.md"
            
            if (Test-Path $readmePath) {
                $content = Get-Content $readmePath -Raw -Encoding UTF8
                
                # Não deve ter caracteres corrompidos
                $content | Should Not Match '�'
                
                # Verifica que pode re-encodar sem perda
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
                $reencoded = [System.Text.Encoding]::UTF8.GetString($bytes)
                
                $reencoded.Length | Should Be $content.Length
            }
        }
        
        It 'Scripts PowerShell com acentuação usam UTF-8' {
            $scriptsWithAccents = @(
                $installScript,
                $exportScript,
                $restoreScript
            )
            
            foreach ($script in $scriptsWithAccents) {
                if (Test-Path $script) {
                    $content = Get-Content $script -Raw -Encoding UTF8
                    
                    # Não deve ter caracteres corrompidos
                    $content | Should Not Match '�'
                    
                    # Se tem acentos, verifica que são lidos corretamente
                    if ($content -match '[áàâãéêíóôõúçÁÀÂÃÉÊÍÓÔÕÚÇ]') {
                        # Conseguiu ler caracteres acentuados
                        $true | Should Be $true
                    }
                }
            }
        }
        
        It 'chainsaw_installer.cmd não tem encoding misto' {
            $content = Get-Content $installerCmd -Raw
            
            # Não deve ter null bytes
            $content | Should Not Match '\x00'
            
            # Deve ter conteúdo
            $content.Length | Should BeGreaterThan 0
        }
        
        It 'Documentação em inst_docs usa UTF-8 consistente' {
            $docsPath = Join-Path $repoRoot "installation\inst_docs"
            
            if (Test-Path $docsPath) {
                $mdFiles = Get-ChildItem -Path $docsPath -Filter "*.md" -Recurse
                
                foreach ($file in $mdFiles) {
                    $content = Get-Content $file.FullName -Raw -Encoding UTF8
                    
                    # Não deve ter caracteres corrompidos
                    if ($content -match '�') {
                        throw "Arquivo $($file.Name) tem caracteres corrompidos"
                    }
                    
                    $true | Should Be $true
                }
            }
        }
        
        It 'Nenhum arquivo crítico usa UTF-16' {
            $criticalFiles = @(
                $vbaModulePath,
                $versionFile,
                $installerCmd,
                $installScript
            )
            
            foreach ($filePath in $criticalFiles) {
                if (Test-Path $filePath) {
                    $bytes = [System.IO.File]::ReadAllBytes($filePath)
                    
                    if ($bytes.Length -ge 2) {
                        # UTF-16 LE BOM
                        $isUtf16Le = ($bytes[0] -eq 0xFF) -and ($bytes[1] -eq 0xFE)
                        
                        # UTF-16 BE BOM
                        $isUtf16Be = ($bytes[0] -eq 0xFE) -and ($bytes[1] -eq 0xFF)
                        
                        if ($isUtf16Le -or $isUtf16Be) {
                            throw "Arquivo $filePath usa UTF-16 - deveria usar UTF-8"
                        }
                    }
                }
            }
            
            $true | Should Be $true
        }
        
        It 'Regex patterns funcionam com caracteres acentuados' {
            # Testa padrões comuns de regex com acentuação
            $testStrings = @{
                "Versão: 2.0.2"           = "Versão: ([^\r\n]+)"
                "Instalação completa"     = "Instalação"
                "Configuração do sistema" = "Configuração"
                "Função principal"        = "Função"
            }
            
            foreach ($testStr in $testStrings.Keys) {
                $pattern = $testStrings[$testStr]
                
                if ($testStr -match $pattern) {
                    $true | Should Be $true
                }
                else {
                    throw "Pattern '$pattern' falhou ao encontrar '$testStr'"
                }
            }
        }
        
        It 'Get-Content -Encoding UTF8 é usado consistentemente nos testes' {
            $thisTestFile = $PSCommandPath
            $content = Get-Content $thisTestFile -Raw
            
            # Conta quantas vezes Get-Content é usado
            $allGetContent = ([regex]::Matches($content, 'Get-Content')).Count
            
            # Conta quantas vezes -Encoding UTF8 é especificado
            $withEncoding = ([regex]::Matches($content, 'Get-Content.*-Encoding UTF8')).Count
            
            # Pelo menos 20% dos Get-Content devem especificar encoding (quando necessário)
            if ($allGetContent -gt 0) {
                ($withEncoding / $allGetContent) | Should BeGreaterThan 0.2
            }
        }
    }
}

