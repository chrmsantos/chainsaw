# =============================================================================
# CHAINSAW - Remover Emojis via Bytes UTF-8
# =============================================================================
# Remove emojis diretamente manipulando bytes UTF-8
# =============================================================================

param(
    [switch]$WhatIf = $false
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "CHAINSAW - Remocao de Emojis (Bytes)" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

if ($WhatIf) {
    Write-Host "MODO SIMULACAO - Nenhuma alteracao sera feita`n" -ForegroundColor Yellow
}

$totalFilesModified = 0
$totalEmojisRemoved = 0

# ==================================================================
# Funcao: Remover emojis de um arquivo
# ==================================================================
function Remove-EmojisFromFile {
    param([string]$FilePath)
    
    $fileName = Split-Path $FilePath -Leaf
    $bytes = [System.IO.File]::ReadAllBytes($FilePath)
    $newBytes = [System.Collections.ArrayList]::new()
    $emojisRemoved = 0
    
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        # Detecta sequencias de emoji F0 9F XX XX (4 bytes)
        if ($i -lt ($bytes.Length - 3) -and 
            $bytes[$i] -eq 0xF0 -and $bytes[$i+1] -eq 0x9F) {
            
            # Pula os 4 bytes do emoji
            $i += 3
            $emojisRemoved++
            continue
        }
        
        # Detecta simbolos decorativos E2 9C-AD XX (3 bytes)
        if ($i -lt ($bytes.Length - 2) -and 
            $bytes[$i] -eq 0xE2 -and 
            $bytes[$i+1] -ge 0x9C -and $bytes[$i+1] -le 0xAD) {
            
            # Pula os 3 bytes do simbolo
            $i += 2
            $emojisRemoved++
            continue
        }
        
        # Byte normal, adiciona
        [void]$newBytes.Add($bytes[$i])
    }
    
    if ($emojisRemoved -gt 0) {
        if (-not $WhatIf) {
            [System.IO.File]::WriteAllBytes($FilePath, $newBytes.ToArray())
        }
        
        Write-Host "[MODIFICADO] $fileName - $emojisRemoved emojis removidos" -ForegroundColor Green
        $script:totalFilesModified++
        $script:totalEmojisRemoved += $emojisRemoved
    }
}

# ==================================================================
# Processar Arquivos Markdown
# ==================================================================
Write-Host "Processando Arquivos Markdown...`n" -ForegroundColor White

$mdFiles = @(
    Get-ChildItem -Path "$projectRoot\docs" -Filter "*.md" -Recurse -ErrorAction SilentlyContinue
    Get-ChildItem -Path "$projectRoot" -Filter "*.md" -File -ErrorAction SilentlyContinue
    Get-ChildItem -Path "$projectRoot\installation\inst_docs" -Filter "*.md" -Recurse -ErrorAction SilentlyContinue
)

foreach ($file in $mdFiles) {
    if ($file -ne $null) {
        Remove-EmojisFromFile -FilePath $file.FullName
    }
}

# ==================================================================
# Processar Scripts PowerShell
# ==================================================================
Write-Host "`nProcessando Scripts PowerShell...`n" -ForegroundColor White

$psFiles = Get-ChildItem -Path "$projectRoot\installation\inst_scripts" -Filter "*.ps1" -Recurse

foreach ($file in $psFiles) {
    Remove-EmojisFromFile -FilePath $file.FullName
}

# ==================================================================
# Relatorio Final
# ==================================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "RESULTADO" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Arquivos modificados: $totalFilesModified" -ForegroundColor White
Write-Host "Total de emojis removidos: $totalEmojisRemoved" -ForegroundColor White
Write-Host "========================================`n" -ForegroundColor Cyan

if ($WhatIf) {
    Write-Host "Execute sem -WhatIf para aplicar as alteracoes." -ForegroundColor Yellow
} else {
    Write-Host "Alteracoes aplicadas com sucesso!" -ForegroundColor Green
}
