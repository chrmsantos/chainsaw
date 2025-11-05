# =============================================================================
# CHAINSAW - Teste de Permissões (Sem Privilégios de Administrador)
# =============================================================================
# Este script verifica se o sistema pode executar todas as operações necessárias
# sem privilégios de administrador
# =============================================================================

Write-Host ""
Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║      CHAINSAW - Teste de Permissões (Usuário Normal)          ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

$allTests = $true

# Teste 1: Verifica se NÃO está executando como administrador
Write-Host "1. Verificando modo de execução..." -NoNewline
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "   Executando como usuário normal (correto)" -ForegroundColor Gray
} else {
    Write-Host " ⚠" -ForegroundColor Yellow
    Write-Host "   Executando como Administrador (não recomendado)" -ForegroundColor Yellow
}

# Teste 2: Permissões de escrita em %USERPROFILE%
Write-Host ""
Write-Host "2. Testando escrita em %USERPROFILE%..." -NoNewline
try {
    $testFile = Join-Path $env:USERPROFILE "chainsaw_permtest_$(Get-Date -Format 'yyyyMMddHHmmss').tmp"
    "test" | Out-File $testFile -Force
    Remove-Item $testFile -Force -ErrorAction SilentlyContinue
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "   Permissões OK: $env:USERPROFILE" -ForegroundColor Gray
} catch {
    Write-Host " ✗" -ForegroundColor Red
    Write-Host "   Sem permissões de escrita: $env:USERPROFILE" -ForegroundColor Red
    $allTests = $false
}

# Teste 3: Criação de diretórios em %USERPROFILE%
Write-Host ""
Write-Host "3. Testando criação de diretórios em %USERPROFILE%..." -NoNewline
try {
    $testDir = Join-Path $env:USERPROFILE "chainsaw\test_$(Get-Date -Format 'yyyyMMddHHmmss')"
    New-Item -Path $testDir -ItemType Directory -Force | Out-Null
    Remove-Item $testDir -Recurse -Force -ErrorAction SilentlyContinue
    
    # Remove pasta pai se estiver vazia
    $parentDir = Join-Path $env:USERPROFILE "chainsaw"
    if (Test-Path $parentDir) {
        $items = Get-ChildItem $parentDir
        if ($items.Count -eq 0) {
            Remove-Item $parentDir -Force -ErrorAction SilentlyContinue
        }
    }
    
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "   Criação de diretórios OK" -ForegroundColor Gray
} catch {
    Write-Host " ✗" -ForegroundColor Red
    Write-Host "   Falha ao criar diretórios" -ForegroundColor Red
    $allTests = $false
}

# Teste 4: Permissões de escrita em %APPDATA%
Write-Host ""
Write-Host "4. Testando escrita em %APPDATA%..." -NoNewline
try {
    $testFile = Join-Path $env:APPDATA "chainsaw_permtest_$(Get-Date -Format 'yyyyMMddHHmmss').tmp"
    "test" | Out-File $testFile -Force
    Remove-Item $testFile -Force -ErrorAction SilentlyContinue
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "   Permissões OK: $env:APPDATA" -ForegroundColor Gray
} catch {
    Write-Host " ✗" -ForegroundColor Red
    Write-Host "   Sem permissões de escrita: $env:APPDATA" -ForegroundColor Red
    $allTests = $false
}

# Teste 5: Renomeação de pastas em %APPDATA%
Write-Host ""
Write-Host "5. Testando renomeação de pastas em %APPDATA%..." -NoNewline
try {
    $testDir = Join-Path $env:APPDATA "chainsaw_test_$(Get-Date -Format 'yyyyMMddHHmmss')"
    $testDirRenamed = "${testDir}_renamed"
    
    New-Item -Path $testDir -ItemType Directory -Force | Out-Null
    Rename-Item -Path $testDir -NewName (Split-Path $testDirRenamed -Leaf) -Force
    Remove-Item $testDirRenamed -Recurse -Force -ErrorAction SilentlyContinue
    
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "   Renomeação de pastas OK" -ForegroundColor Gray
} catch {
    Write-Host " ✗" -ForegroundColor Red
    Write-Host "   Falha ao renomear pastas" -ForegroundColor Red
    $allTests = $false
}

# Teste 6: Cópia de arquivos
Write-Host ""
Write-Host "6. Testando cópia de arquivos..." -NoNewline
try {
    $sourceFile = Join-Path $env:TEMP "chainsaw_source_$(Get-Date -Format 'yyyyMMddHHmmss').tmp"
    $destFile = Join-Path $env:USERPROFILE "chainsaw_dest_$(Get-Date -Format 'yyyyMMddHHmmss').tmp"
    
    "test content" | Out-File $sourceFile -Force
    Copy-Item -Path $sourceFile -Destination $destFile -Force
    
    Remove-Item $sourceFile -Force -ErrorAction SilentlyContinue
    Remove-Item $destFile -Force -ErrorAction SilentlyContinue
    
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "   Cópia de arquivos OK" -ForegroundColor Gray
} catch {
    Write-Host " ✗" -ForegroundColor Red
    Write-Host "   Falha ao copiar arquivos" -ForegroundColor Red
    $allTests = $false
}

# Teste 7: Cópia recursiva de diretórios
Write-Host ""
Write-Host "7. Testando cópia recursiva de diretórios..." -NoNewline
try {
    $sourceDir = Join-Path $env:TEMP "chainsaw_source_dir_$(Get-Date -Format 'yyyyMMddHHmmss')"
    $destDir = Join-Path $env:USERPROFILE "chainsaw_dest_dir_$(Get-Date -Format 'yyyyMMddHHmmss')"
    
    # Cria estrutura de teste
    New-Item -Path "$sourceDir\subdir1" -ItemType Directory -Force | Out-Null
    New-Item -Path "$sourceDir\subdir2" -ItemType Directory -Force | Out-Null
    "test1" | Out-File "$sourceDir\file1.txt" -Force
    "test2" | Out-File "$sourceDir\subdir1\file2.txt" -Force
    
    # Copia recursivamente
    Copy-Item -Path $sourceDir -Destination $destDir -Recurse -Force
    
    # Verifica se copiou
    $success = (Test-Path "$destDir\file1.txt") -and (Test-Path "$destDir\subdir1\file2.txt")
    
    # Limpa
    Remove-Item $sourceDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $destDir -Recurse -Force -ErrorAction SilentlyContinue
    
    if ($success) {
        Write-Host " ✓" -ForegroundColor Green
        Write-Host "   Cópia recursiva OK" -ForegroundColor Gray
    } else {
        throw "Arquivos não foram copiados corretamente"
    }
} catch {
    Write-Host " ✗" -ForegroundColor Red
    Write-Host "   Falha na cópia recursiva" -ForegroundColor Red
    $allTests = $false
}

# Teste 8: Acesso a informações do sistema
Write-Host ""
Write-Host "8. Testando acesso a informações do sistema..." -NoNewline
try {
    $osVersion = [Environment]::OSVersion.Version
    $psVersion = $PSVersionTable.PSVersion
    $userName = $env:USERNAME
    $computerName = $env:COMPUTERNAME
    
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "   Acesso a informações OK" -ForegroundColor Gray
    Write-Host "   - OS: Windows $($osVersion.Major).$($osVersion.Minor)" -ForegroundColor DarkGray
    Write-Host "   - PowerShell: $($psVersion.ToString())" -ForegroundColor DarkGray
    Write-Host "   - Usuário: $userName" -ForegroundColor DarkGray
    Write-Host "   - Computador: $computerName" -ForegroundColor DarkGray
} catch {
    Write-Host " ✗" -ForegroundColor Red
    Write-Host "   Falha ao acessar informações do sistema" -ForegroundColor Red
    $allTests = $false
}

# Resultado final
Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray

if ($allTests) {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║     ✓ TODOS OS TESTES PASSARAM!                                ║" -ForegroundColor Green
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "✓ O sistema está pronto para executar a instalação" -ForegroundColor Green
    Write-Host "✓ Nenhum privilégio de administrador é necessário" -ForegroundColor Green
    Write-Host ""
    
    if ($isAdmin) {
        Write-Host "⚠ IMPORTANTE: Execute o install.ps1 SEM privilégios de administrador" -ForegroundColor Yellow
        Write-Host ""
    }
} else {
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║     ✗ ALGUNS TESTES FALHARAM                                   ║" -ForegroundColor Red
    Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Red
    Write-Host ""
    Write-Host "⚠ Verifique as permissões do seu usuário." -ForegroundColor Yellow
    Write-Host ""
}

# Informações sobre permissões
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host ""
Write-Host "ℹ SOBRE PERMISSÕES:" -ForegroundColor Cyan
Write-Host ""
Write-Host "O script de instalação do Chainsaw opera APENAS nas seguintes áreas:" -ForegroundColor White
Write-Host "  • %USERPROFILE%\chainsaw\          (seus arquivos)" -ForegroundColor Gray
Write-Host "  • %APPDATA%\Microsoft\Templates\   (configurações do Word)" -ForegroundColor Gray
Write-Host ""
Write-Host "Estas pastas fazem parte do SEU perfil de usuário e NÃO requerem" -ForegroundColor White
Write-Host "privilégios de administrador para serem modificadas." -ForegroundColor White
Write-Host ""
Write-Host "⚠ Se você executar como Administrador:" -ForegroundColor Yellow
Write-Host "  • Os arquivos serão criados com proprietário 'Administrador'" -ForegroundColor Gray
Write-Host "  • Você pode ter problemas de acesso depois" -ForegroundColor Gray
Write-Host "  • O Word pode não conseguir acessar os templates" -ForegroundColor Gray
Write-Host ""
