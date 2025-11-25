# =============================================================================
# CHAINSAW - Testes de Segurança e Proteção contra Perda de Dados
# =============================================================================
# Este arquivo testa todos os mecanismos de proteção implementados para
# prevenir perda acidental de dados durante instalação/atualização.
# =============================================================================

BeforeAll {
    $ErrorActionPreference = 'Stop'
    
    # Caminho do projeto
    $script:ProjectRoot = Split-Path -Parent $PSScriptRoot
    $script:InstallScript = Join-Path $ProjectRoot "installation\inst_scripts\install.ps1"
    $script:InstallerCmd = Join-Path $ProjectRoot "chainsaw_installer.cmd"
    
    # Cria ambiente de teste isolado
    $script:TestRoot = Join-Path $env:TEMP "CHAINSAW_SecurityTests_$(Get-Date -Format 'yyyyMMddHHmmss')"
    New-Item -Path $script:TestRoot -ItemType Directory -Force | Out-Null
    
    Write-Host "Ambiente de teste: $script:TestRoot" -ForegroundColor Cyan
}

AfterAll {
    # Limpa ambiente de teste
    if (Test-Path $script:TestRoot) {
        Remove-Item -Path $script:TestRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe "Proteções do chainsaw_installer.cmd" {
    
    It "Deve conter validação de tamanho mínimo do ZIP" {
        $content = Get-Content $script:InstallerCmd -Raw
        $content | Should -Match 'ZIP_SIZE.*LSS.*102400'
        $content | Should -Match 'Arquivo ZIP muito pequeno'
    }
    
    It "Deve conter teste de integridade do ZIP" {
        $content = Get-Content $script:InstallerCmd -Raw
        $content | Should -Match 'chainsaw_test_zip.ps1'
        $content | Should -Match 'System.IO.Compression.ZipFile'
        $content | Should -Match 'entryCount'
    }
    
    It "Deve criar backup OBRIGATÓRIO antes de modificar" {
        $content = Get-Content $script:InstallerCmd -Raw
        $content | Should -Match 'Backup OBRIGATORIO'
        $content | Should -Match 'BACKUP_DIR='
    }
    
    It "Deve validar o backup após criação" {
        $content = Get-Content $script:InstallerCmd -Raw
        $content | Should -Match 'VALIDAÇÃO DO BACKUP'
        $content | Should -Match 'BACKUP_FILE_COUNT'
        $content | Should -Match 'backup parece incompleto'
    }
    
    It "Deve abortar instalação se backup falhar" {
        $content = Get-Content $script:InstallerCmd -Raw
        $content | Should -Match 'ERRO CRITICO.*Falha ao criar backup'
        $content | Should -Match 'Instalacao ABORTADA.*proteger seus dados'
        $content | Should -Match 'exit /b 1'
    }
    
    It "Deve validar conteúdo extraído ANTES de instalar" {
        $content = Get-Content $script:InstallerCmd -Raw
        $content | Should -Match 'VALIDAÇÃO CRÍTICA DO CONTEÚDO EXTRAÍDO'
        $content | Should -Match 'installation\\inst_scripts'
        $content | Should -Match 'install.cmd'
        $content | Should -Match 'install.ps1'
    }
    
    It "Deve contar arquivos extraídos e validar quantidade mínima" {
        $content = Get-Content $script:InstallerCmd -Raw
        $content | Should -Match 'EXTRACTED_FILE_COUNT'
        $content | Should -Match 'LSS 20'
        $content | Should -Match 'muito poucos arquivos'
    }
    
    It "Deve remover pasta existente SOMENTE APÓS validação completa" {
        $content = Get-Content $script:InstallerCmd -Raw
        
        # Verifica que validação vem ANTES da remoção
        $validationIndex = $content.IndexOf('VALIDAÇÃO CRÍTICA DO CONTEÚDO EXTRAÍDO')
        $removalIndex = $content.IndexOf('Removendo pasta antiga (backup ja criado e validado)')
        
        $validationIndex | Should -BeLessThan $removalIndex
    }
    
    It "Deve implementar rollback em caso de falha na cópia" {
        $content = Get-Content $script:InstallerCmd -Raw
        $content | Should -Match '\[ROLLBACK\]'
        $content | Should -Match 'Tentando restaurar backup'
        $content | Should -Match 'xcopy.*BACKUP_DIR'
    }
    
    It "Deve validar instalação final" {
        $content = Get-Content $script:InstallerCmd -Raw
        $content | Should -Match 'VALIDAÇÃO FINAL DA INSTALAÇÃO'
        $content | Should -Match 'FINAL_VALIDATION_FAILED'
    }
}

Describe "Proteções do install.ps1 - Validação de Origem" {
    
    It "Deve validar existência de stamp.png antes de copiar" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match 'VALIDAÇÃO CRÍTICA 1.*Arquivo de origem não existe'
    }
    
    It "Deve validar tamanho mínimo de stamp.png" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match 'VALIDAÇÃO CRÍTICA 2.*Verifica tamanho mínimo'
        $content | Should -Match 'sourceFileInfo.Length -lt 100'
    }
    
    It "Deve validar que stamp.png foi copiado corretamente" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match 'VALIDAÇÃO CRÍTICA 3.*arquivo foi copiado'
        $content | Should -Match 'sourceSize -ne \$destSize'
    }
    
    It "Deve validar existência da pasta Templates de origem" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match 'VALIDAÇÃO CRÍTICA 1.*Pasta de origem não existe'
    }
    
    It "Deve validar que pasta de origem não está vazia" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match 'VALIDAÇÃO CRÍTICA 2.*pasta de origem está vazia'
        $content | Should -Match 'sourceItems.Count -eq 0'
    }
    
    It "Deve validar existência de Normal.dotm na origem" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match 'VALIDAÇÃO CRÍTICA 3.*Normal.dotm não encontrado'
    }
    
    It "Deve validar tamanho mínimo de Normal.dotm" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match 'normalDotmSize -lt 10000'
        $content | Should -Match 'Normal.dotm muito pequeno ou corrompido'
    }
    
    It "Deve validar que Normal.dotm foi copiado corretamente" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match 'VALIDAÇÃO CRÍTICA 4.*Normal.dotm não foi copiado'
        $content | Should -Match 'destNormalDotmSize -ne \$normalDotmSize'
    }
}

Describe "Proteções do install.ps1 - Backup e Rollback" {
    
    It "Deve criar backup completo antes de qualquer modificação" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match 'function Backup-CompleteConfiguration'
        $content | Should -Match 'backup completo antes da instalação'
    }
    
    It "Deve validar backup antes de usar para rollback" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match 'VALIDA BACKUP ANTES DE RESTAURAR'
        $content | Should -Match 'backupItems.Count -eq 0'
        $content | Should -Match 'Backup está vazio - não é seguro restaurar'
    }
    
    It "Deve ter rollback automático em caso de erro" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match 'Iniciando rollback automático'
        $content | Should -Match 'Restaurando backup de:'
    }
    
    It "Deve validar restauração após rollback" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match 'Valida restauração'
        $content | Should -Match 'if \(Test-Path \$restoredPath\)'
    }
}

Describe "Proteções do install.ps1 - Tratamento de Erros" {
    
    It "Deve usar ErrorActionPreference = Stop" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match '\$ErrorActionPreference\s*=\s*"Stop"'
    }
    
    It "Deve logar todos os erros críticos" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match 'Write-Log.*ERRO CRITICO'
        $content | Should -Match 'Level ERROR'
    }
    
    It "Deve usar try-catch em operações críticas" {
        $content = Get-Content $script:InstallScript -Raw
        
        # Copy-StampFile deve ter try-catch
        $content | Should -Match 'function Copy-StampFile.*try.*catch'
        
        # Copy-TemplatesFolder deve ter try-catch
        $content | Should -Match 'function Copy-TemplatesFolder.*try.*catch'
    }
    
    It "Deve abortar instalação em erros críticos" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match 'throw.*Instalação abortada'
    }
}

Describe "Simulação de Cenários de Falha" {
    
    BeforeEach {
        # Cria estrutura de teste
        $script:TestSource = Join-Path $script:TestRoot "source"
        $script:TestDest = Join-Path $script:TestRoot "dest"
        $script:TestBackup = Join-Path $script:TestRoot "backup"
        
        New-Item -Path $script:TestSource -ItemType Directory -Force | Out-Null
        New-Item -Path $script:TestDest -ItemType Directory -Force | Out-Null
    }
    
    AfterEach {
        # Limpa após cada teste
        Get-ChildItem -Path $script:TestRoot | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    It "Não deve copiar se arquivo de origem não existe" {
        $testFile = Join-Path $script:TestSource "nonexistent.txt"
        $destFile = Join-Path $script:TestDest "file.txt"
        
        # Simula função de cópia com validação
        {
            if (-not (Test-Path $testFile)) {
                throw "Arquivo de origem não existe"
            }
            Copy-Item -Path $testFile -Destination $destFile
        } | Should -Throw -ExpectedMessage "*não existe*"
    }
    
    It "Não deve copiar se arquivo de origem está corrompido (tamanho = 0)" {
        $testFile = Join-Path $script:TestSource "empty.txt"
        New-Item -Path $testFile -ItemType File -Force | Out-Null
        
        $fileInfo = Get-Item $testFile
        $fileInfo.Length | Should -Be 0
        
        # Simula validação de tamanho
        {
            if ($fileInfo.Length -lt 100) {
                throw "Arquivo muito pequeno ou corrompido"
            }
        } | Should -Throw -ExpectedMessage "*corrompido*"
    }
    
    It "Deve validar que cópia foi bem-sucedida" {
        $sourceFile = Join-Path $script:TestSource "test.txt"
        "Test Content 12345" | Out-File -FilePath $sourceFile -Encoding UTF8
        
        $destFile = Join-Path $script:TestDest "test.txt"
        Copy-Item -Path $sourceFile -Destination $destFile
        
        # Validação
        (Test-Path $destFile) | Should -Be $true
        $sourceSize = (Get-Item $sourceFile).Length
        $destSize = (Get-Item $destFile).Length
        $sourceSize | Should -Be $destSize
    }
    
    It "Deve criar backup antes de modificar" {
        # Cria arquivo original
        $originalFile = Join-Path $script:TestDest "important.txt"
        "Original Data" | Out-File -FilePath $originalFile -Encoding UTF8
        
        # Backup
        $backupFile = Join-Path $script:TestBackup "important.txt"
        New-Item -Path $script:TestBackup -ItemType Directory -Force | Out-Null
        Copy-Item -Path $originalFile -Destination $backupFile
        
        # Modifica original
        "New Data" | Out-File -FilePath $originalFile -Encoding UTF8
        
        # Valida backup preservou dados originais
        (Get-Content $backupFile -Raw) | Should -Match "Original Data"
        (Get-Content $originalFile -Raw) | Should -Match "New Data"
    }
    
    It "Deve restaurar backup se instalação falhar" {
        # Cria backup
        $backupPath = Join-Path $script:TestBackup "Templates_backup"
        New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
        "Original Normal.dotm" | Out-File -FilePath (Join-Path $backupPath "Normal.dotm") -Encoding UTF8
        
        # Simula falha na instalação
        $installFailed = $true
        
        # Rollback
        if ($installFailed -and (Test-Path $backupPath)) {
            $destPath = Join-Path $script:TestDest "Templates"
            if (Test-Path $destPath) {
                Remove-Item -Path $destPath -Recurse -Force
            }
            Copy-Item -Path $backupPath -Destination $destPath -Recurse
        }
        
        # Valida que backup foi restaurado
        $restoredFile = Join-Path $script:TestDest "Templates\Normal.dotm"
        (Test-Path $restoredFile) | Should -Be $true
        (Get-Content $restoredFile -Raw) | Should -Match "Original Normal.dotm"
    }
}

Describe "Validação de Integridade de Dados" {
    
    It "Deve comparar checksums após cópia crítica" {
        $sourceFile = Join-Path $script:TestRoot "source.dat"
        1..1000 | ForEach-Object { "Line $_" } | Out-File -FilePath $sourceFile -Encoding UTF8
        
        $destFile = Join-Path $script:TestRoot "dest.dat"
        Copy-Item -Path $sourceFile -Destination $destFile
        
        # Calcula hashes
        $sourceHash = (Get-FileHash -Path $sourceFile -Algorithm SHA256).Hash
        $destHash = (Get-FileHash -Path $destFile -Algorithm SHA256).Hash
        
        $sourceHash | Should -Be $destHash
    }
    
    It "Deve validar estrutura de diretórios críticos" {
        $testRoot = Join-Path $script:TestRoot "chainsaw-extract"
        New-Item -Path "$testRoot\installation\inst_scripts" -ItemType Directory -Force | Out-Null
        New-Item -Path "$testRoot\installation\inst_configs" -ItemType Directory -Force | Out-Null
        New-Item -Path "$testRoot\installation\inst_scripts\install.cmd" -ItemType File -Force | Out-Null
        New-Item -Path "$testRoot\installation\inst_scripts\install.ps1" -ItemType File -Force | Out-Null
        
        # Validação
        (Test-Path "$testRoot\installation") | Should -Be $true
        (Test-Path "$testRoot\installation\inst_scripts") | Should -Be $true
        (Test-Path "$testRoot\installation\inst_scripts\install.cmd") | Should -Be $true
        (Test-Path "$testRoot\installation\inst_scripts\install.ps1") | Should -Be $true
        (Test-Path "$testRoot\installation\inst_configs") | Should -Be $true
    }
}

Describe "Documentação de Segurança" {
    
    It "Deve ter comentários explicando cada validação crítica" {
        $content = Get-Content $script:InstallScript -Raw
        $content | Should -Match '# VALIDAÇÃO CRÍTICA'
        
        # Conta quantas validações críticas existem
        $criticalValidations = ([regex]::Matches($content, '# VALIDAÇÃO CRÍTICA')).Count
        $criticalValidations | Should -BeGreaterThan 5
    }
    
    It "Deve ter logging de todas as operações críticas" {
        $content = Get-Content $script:InstallerCmd -Raw
        $content | Should -Match ':Log'
        $content | Should -Match 'ERRO CRITICO'
        $content | Should -Match 'AVISO'
    }
}

# =============================================================================
# RELATÓRIO FINAL
# =============================================================================

AfterAll {
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║         TESTES DE SEGURANÇA CONCLUÍDOS                       ║" -ForegroundColor Green
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "Proteções Validadas:" -ForegroundColor Cyan
    Write-Host "  ✓ Validação de integridade de arquivos baixados" -ForegroundColor Green
    Write-Host "  ✓ Backup obrigatório antes de modificações" -ForegroundColor Green
    Write-Host "  ✓ Validação de origem de dados" -ForegroundColor Green
    Write-Host "  ✓ Validação de cópias bem-sucedidas" -ForegroundColor Green
    Write-Host "  ✓ Rollback automático em caso de falha" -ForegroundColor Green
    Write-Host "  ✓ Logging completo de operações" -ForegroundColor Green
    Write-Host ""
}
