# =============================================================================
# CHAINSAW - Validador de Encoding e Emojis
# =============================================================================
# Script para verificar encoding UTF-8 e ausencia de emojis
# =============================================================================

param(
    [switch]$Verbose = $false
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "CHAINSAW - Validacao de Encoding" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Funcao para detectar emojis via bytes UTF-8
function Test-ContainsEmoji {
    param([byte[]]$Bytes)
    
    for ($i = 0; $i -lt $Bytes.Length - 3; $i++) {
        # TODOS os emojis começam com F0 9F (UTF-8 encoding de U+1F000 em diante)
        if ($Bytes[$i] -eq 0xF0 -and $Bytes[$i+1] -eq 0x9F) {
            return $true
        }
        
        # Simbolos diversos comecam com E2 e segundo byte especifico
        # E2 9C (Check marks, crosses), E2 9D (Crosses, hearts),
        # E2 AD (Warnings), E2 9E (Arrows), E2 AD (Star), E2 AC (Geometric)
        if ($Bytes[$i] -eq 0xE2) {
            $secondByte = $Bytes[$i+1]
            # Ranges comuns de simbolos decorativos
            if ($secondByte -ge 0x9C -and $secondByte -le 0xAD) {
                return $true
            }
        }
    }
    
    return $false
}

$totalErrors = 0
$totalWarnings = 0
$filesChecked = 0

# ==================================================================
# Funcao: Verificar arquivo por emojis
# ==================================================================
function Test-FileForEmojis {
    param([string]$FilePath, [string]$FileType)
    
    $script:filesChecked++
    
    try {
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        
        if (Test-ContainsEmoji -Bytes $bytes) {
            Write-Host "  [ERRO] $FileType contem emoji: $($FilePath | Split-Path -Leaf)" -ForegroundColor Red
            $script:totalErrors++
            return $false
        }
        
        if ($Verbose) {
            Write-Host "  [OK] $($FilePath | Split-Path -Leaf)" -ForegroundColor Green
        }
        
        return $true
        
    } catch {
        Write-Host "  [AVISO] Erro ao ler arquivo: $($FilePath | Split-Path -Leaf)" -ForegroundColor Yellow
        Write-Host "    Erro: $($_.Exception.Message)" -ForegroundColor Yellow
        $script:totalWarnings++
        return $false
    }
}

# ==================================================================
# Funcao: Verificar encoding UTF-8
# ==================================================================
function Test-FileEncoding {
    param([string]$FilePath, [string]$FileType)
    
    try {
        $bytes = [System.IO.File]::ReadAllBytes($FilePath)
        
        # Verifica UTF-8 com BOM
        $hasUtf8Bom = ($bytes.Length -ge 3) -and 
                      ($bytes[0] -eq 0xEF) -and 
                      ($bytes[1] -eq 0xBB) -and 
                      ($bytes[2] -eq 0xBF)
        
        # Verifica se e ASCII puro
        $isAscii = $true
        foreach ($byte in $bytes) {
            if ($byte -ge 128) {
                $isAscii = $false
                break
            }
        }
        
        # Verifica se pode ser decodificado como UTF-8
        try {
            $utf8Content = [System.Text.Encoding]::UTF8.GetString($bytes)
            $isValidUtf8 = $true
        } catch {
            $isValidUtf8 = $false
        }
        
        if (-not ($hasUtf8Bom -or $isAscii -or $isValidUtf8)) {
            Write-Host "  [AVISO] $FileType pode ter encoding invalido: $($FilePath | Split-Path -Leaf)" -ForegroundColor Yellow
            $script:totalWarnings++
            return $false
        }
        
        return $true
        
    } catch {
        Write-Host "  [AVISO] Erro ao verificar encoding: $($FilePath | Split-Path -Leaf)" -ForegroundColor Yellow
        $script:totalWarnings++
        return $false
    }
}

# ==================================================================
# Verificar Scripts PowerShell
# ==================================================================
Write-Host "Verificando Scripts PowerShell..." -ForegroundColor White

$psFiles = Get-ChildItem -Path "$projectRoot\installation\inst_scripts" -Filter "*.ps1" -Recurse

foreach ($file in $psFiles) {
    Test-FileForEmojis -FilePath $file.FullName -FileType "Script PS1"
    Test-FileEncoding -FilePath $file.FullName -FileType "Script PS1"
}

# ==================================================================
# Verificar Arquivos Markdown
# ==================================================================
Write-Host "`nVerificando Arquivos Markdown..." -ForegroundColor White

$mdFiles = @(
    Get-ChildItem -Path "$projectRoot\docs" -Filter "*.md" -Recurse -ErrorAction SilentlyContinue
    Get-ChildItem -Path "$projectRoot" -Filter "*.md" -File -ErrorAction SilentlyContinue
    Get-ChildItem -Path "$projectRoot\installation\inst_docs" -Filter "*.md" -Recurse -ErrorAction SilentlyContinue
)

foreach ($file in $mdFiles) {
    if ($file -ne $null) {
        Test-FileForEmojis -FilePath $file.FullName -FileType "Markdown"
        Test-FileEncoding -FilePath $file.FullName -FileType "Markdown"
    }
}

# ==================================================================
# Verificar Arquivo VBA
# ==================================================================
Write-Host "`nVerificando Arquivo VBA..." -ForegroundColor White

$vbaFile = "$projectRoot\source\main\monolithicMod.bas"
if (Test-Path $vbaFile) {
    Test-FileForEmojis -FilePath $vbaFile -FileType "VBA"
    Test-FileEncoding -FilePath $vbaFile -FileType "VBA"
}

# ==================================================================
# Verificar Testes PowerShell
# ==================================================================
Write-Host "`nVerificando Testes PowerShell..." -ForegroundColor White

$testFiles = Get-ChildItem -Path "$projectRoot\tests" -Filter "*.ps1" -ErrorAction SilentlyContinue

foreach ($file in $testFiles) {
    Test-FileForEmojis -FilePath $file.FullName -FileType "Teste PS1"
    Test-FileEncoding -FilePath $file.FullName -FileType "Teste PS1"
}

# ==================================================================
# Relatorio Final
# ==================================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "RESULTADO DA VALIDACAO" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Arquivos verificados: $filesChecked" -ForegroundColor White
Write-Host "Erros encontrados:    $totalErrors" -ForegroundColor $(if ($totalErrors -eq 0) { "Green" } else { "Red" })
Write-Host "Avisos encontrados:   $totalWarnings" -ForegroundColor $(if ($totalWarnings -eq 0) { "Green" } else { "Yellow" })
Write-Host "========================================`n" -ForegroundColor Cyan

if ($totalErrors -eq 0) {
    Write-Host "SUCESSO: Nenhum emoji encontrado!" -ForegroundColor Green
    exit 0
} else {
    Write-Host "FALHA: Emojis encontrados nos arquivos acima!" -ForegroundColor Red
    exit 1
}
