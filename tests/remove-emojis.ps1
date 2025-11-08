# =============================================================================
# CHAINSAW - Remover Emojis de Arquivos
# =============================================================================
# Script para substituir emojis por equivalentes em texto simples
# =============================================================================

param(
    [switch]$WhatIf = $false
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot

# Mapeamento de emojis para texto usando [char]::ConvertFromUtf32()
$emojiReplacements = @{}

# Simbolos de validacao
$emojiReplacements[[char]::ConvertFromUtf32(0x2713)] = '[OK]'      # checkmark
$emojiReplacements[[char]::ConvertFromUtf32(0x2717)] = '[ERRO]'    # ballot X
$emojiReplacements[[char]::ConvertFromUtf32(0x274C)] = '[ERRO]'    # cross mark  
$emojiReplacements[[char]::ConvertFromUtf32(0x2705)] = '[OK]'      # white check mark
$emojiReplacements[[char]::ConvertFromUtf32(0x26A0)] = '[AVISO]'   # warning sign
$emojiReplacements[[char]::ConvertFromUtf32(0x2B50)] = '*'         # star
$emojiReplacements[[char]::ConvertFromUtf32(0x2728)] = '*'         # sparkles
$emojiReplacements[[char]::ConvertFromUtf32(0x2753)] = '?'         # question mark

# Setas e indicadores
$emojiReplacements[[char]::ConvertFromUtf32(0x27A1)] = '->'        # arrow right
$emojiReplacements[[char]::ConvertFromUtf32(0x25B6)] = '>'         # play button
$emojiReplacements[[char]::ConvertFromUtf32(0x1F504)] = ''         # counterclockwise arrows

# Objetos e documentos
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4C1)] = ''         # file folder
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4C4)] = ''         # page facing up
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4DA)] = ''         # books
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4DD)] = ''         # memo
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4E6)] = ''         # package
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4CA)] = ''         # bar chart
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4BE)] = ''         # floppy disk
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4CB)] = ''         # clipboard
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4D6)] = ''         # open book
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4DC)] = ''         # scroll
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4CC)] = ''         # pushpin
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4D8)] = ''         # blue book
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4E2)] = ''         # loudspeaker
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4E7)] = ''         # e-mail
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4CD)] = ''         # round pushpin
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4DE)] = ''         # telephone receiver
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4E4)] = ''         # outbox tray
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4E5)] = ''         # inbox tray

# Seguranca e ferramentas
$emojiReplacements[[char]::ConvertFromUtf32(0x1F512)] = ''         # lock
$emojiReplacements[[char]::ConvertFromUtf32(0x1F511)] = ''         # key
$emojiReplacements[[char]::ConvertFromUtf32(0x1F510)] = ''         # closed lock with key
$emojiReplacements[[char]::ConvertFromUtf32(0x1F6E1)] = ''         # shield
$emojiReplacements[[char]::ConvertFromUtf32(0x1F50D)] = ''         # magnifying glass
$emojiReplacements[[char]::ConvertFromUtf32(0x1F527)] = ''         # wrench
$emojiReplacements[[char]::ConvertFromUtf32(0x1F517)] = ''         # link

# Outros icones
$emojiReplacements[[char]::ConvertFromUtf32(0x1F4A1)] = ''         # light bulb
$emojiReplacements[[char]::ConvertFromUtf32(0x1F550)] = ''         # clock
$emojiReplacements[[char]::ConvertFromUtf32(0x1F6A8)] = ''         # police car light
$emojiReplacements[[char]::ConvertFromUtf32(0x1F3DB)] = ''         # classical building
$emojiReplacements[[char]::ConvertFromUtf32(0x1F3C6)] = ''         # trophy
$emojiReplacements[[char]::ConvertFromUtf32(0x1F419)] = ''         # octopus
$emojiReplacements[[char]::ConvertFromUtf32(0x1F9EA)] = ''         # test tube
$emojiReplacements[[char]::ConvertFromUtf32(0x1F680)] = ''         # rocket
$emojiReplacements[[char]::ConvertFromUtf32(0x1F393)] = ''         # graduation cap

$totalFilesModified = 0
$totalReplacements = 0

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "CHAINSAW - Remocao de Emojis" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

if ($WhatIf) {
    Write-Host "MODO SIMULACAO - Nenhuma alteracao sera feita`n" -ForegroundColor Yellow
}

# ==================================================================
# Funcao: Processar arquivo
# ==================================================================
function Process-File {
    param([string]$FilePath)
    
    $fileName = Split-Path $FilePath -Leaf
    $modified = $false
    $fileReplacements = 0
    
    try {
        # Le conteudo do arquivo
        $content = Get-Content $FilePath -Raw -Encoding UTF8
        $originalContent = $content
        
        # Aplica todas as substituicoes
        foreach ($emoji in $emojiReplacements.Keys) {
            if ($content -match [regex]::Escape($emoji)) {
                $replacement = $emojiReplacements[$emoji]
                $beforeCount = ($content.ToCharArray() | Where-Object { $_ -eq $emoji }).Count
                $content = $content -replace [regex]::Escape($emoji), $replacement
                $fileReplacements += $beforeCount
                $modified = $true
                
                Write-Host "  - Substituindo '$emoji' por '$replacement' ($beforeCount ocorrencias)" -ForegroundColor Gray
            }
        }
        
        # Salva se houve modificacoes
        if ($modified) {
            if (-not $WhatIf) {
                [System.IO.File]::WriteAllText($FilePath, $content, [System.Text.Encoding]::UTF8)
            }
            
            Write-Host "[MODIFICADO] $fileName - $fileReplacements substituicoes" -ForegroundColor Green
            $script:totalFilesModified++
            $script:totalReplacements += $fileReplacements
        }
        
    } catch {
        Write-Host "[ERRO] Falha ao processar $fileName : $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ==================================================================
# Processar Scripts PowerShell
# ==================================================================
Write-Host "Processando Scripts PowerShell...`n" -ForegroundColor White

$psFiles = Get-ChildItem -Path "$projectRoot\installation\inst_scripts" -Filter "*.ps1" -Recurse

foreach ($file in $psFiles) {
    Process-File -FilePath $file.FullName
}

# ==================================================================
# Processar Arquivos Markdown
# ==================================================================
Write-Host "`nProcessando Arquivos Markdown...`n" -ForegroundColor White

$mdFiles = @(
    Get-ChildItem -Path "$projectRoot\docs" -Filter "*.md" -Recurse -ErrorAction SilentlyContinue
    Get-ChildItem -Path "$projectRoot" -Filter "*.md" -File -ErrorAction SilentlyContinue
    Get-ChildItem -Path "$projectRoot\installation\inst_docs" -Filter "*.md" -Recurse -ErrorAction SilentlyContinue
)

foreach ($file in $mdFiles) {
    if ($file -ne $null) {
        Process-File -FilePath $file.FullName
    }
}

# ==================================================================
# Processar Arquivo VBA
# ==================================================================
Write-Host "`nProcessando Arquivo VBA...`n" -ForegroundColor White

$vbaFile = "$projectRoot\source\backups\main\monolithicMod.bas"
if (Test-Path $vbaFile) {
    Process-File -FilePath $vbaFile
}

# ==================================================================
# Processar Testes PowerShell
# ==================================================================
Write-Host "`nProcessando Testes PowerShell...`n" -ForegroundColor White

$testFiles = Get-ChildItem -Path "$projectRoot\tests" -Filter "*.ps1" -ErrorAction SilentlyContinue

foreach ($file in $testFiles) {
    # Pular este proprio script
    if ($file.Name -ne "remove-emojis.ps1") {
        Process-File -FilePath $file.FullName
    }
}

# ==================================================================
# Relatorio Final
# ==================================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "RESULTADO" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Arquivos modificados: $totalFilesModified" -ForegroundColor White
Write-Host "Total de substituicoes: $totalReplacements" -ForegroundColor White
Write-Host "========================================`n" -ForegroundColor Cyan

if ($WhatIf) {
    Write-Host "Execute sem -WhatIf para aplicar as alteracoes." -ForegroundColor Yellow
} else {
    Write-Host "Alteracoes aplicadas com sucesso!" -ForegroundColor Green
}
