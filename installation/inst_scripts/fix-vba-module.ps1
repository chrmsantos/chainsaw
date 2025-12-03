# =============================================================================
# Fix VBA Module in Normal.dotm
# =============================================================================
# This script replaces the old "Módulo1" with the updated "monolithicMod"

$ErrorActionPreference = "Stop"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  VBA Module Replacement Tool" -ForegroundColor White
Write-Host "========================================`n" -ForegroundColor Cyan

# Paths
$vbaPath = Join-Path $PSScriptRoot "..\..\source\main\monolithicMod.bas"
$normalPath = Join-Path $env:APPDATA "Microsoft\Templates\Normal.dotm"

# Validate source file
if (-not (Test-Path $vbaPath)) {
    Write-Host "[ERROR] Source module not found: $vbaPath" -ForegroundColor Red
    exit 1
}

Write-Host "[1/5] Closing all Word processes..." -ForegroundColor Yellow
Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 3

Write-Host "[2/5] Removing lock files..." -ForegroundColor Yellow
Remove-Item (Join-Path $env:APPDATA "Microsoft\Templates\~`$Normal.dotm") -Force -ErrorAction SilentlyContinue

Write-Host "[3/5] Opening Normal.dotm..." -ForegroundColor Yellow
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open($normalPath)
    
    Write-Host "[4/5] Removing old modules..." -ForegroundColor Yellow
    $oldModules = @("Módulo1", "Module1", "monolithicMod", "CHAINSAW_MODX", "Chainsaw_ModX", "chainsawModX")
    
    foreach ($moduleName in $oldModules) {
        try {
            $module = $doc.VBProject.VBComponents.Item($moduleName)
            if ($module) {
                $doc.VBProject.VBComponents.Remove($module)
                Write-Host "  ✓ Removed: $moduleName" -ForegroundColor Green
            }
        }
        catch {
            # Module doesn't exist, continue
        }
    }
    
    Write-Host "[5/5] Importing monolithicMod..." -ForegroundColor Yellow
    $doc.VBProject.VBComponents.Import($vbaPath) | Out-Null
    Write-Host "  ✓ Imported: monolithicMod" -ForegroundColor Green
    
    # Save changes
    $doc.Save()
    Write-Host "  ✓ Saved Normal.dotm" -ForegroundColor Green
    
    # Verify
    Write-Host "`nVerifying modules:" -ForegroundColor Cyan
    $doc.VBProject.VBComponents | Where-Object {$_.Type -eq 1} | ForEach-Object {
        Write-Host "  → $($_.Name)" -ForegroundColor White
    }
    
    $doc.Close($false)
    $word.Quit()
    
    Write-Host "`n[SUCCESS] VBA module replacement complete!" -ForegroundColor Green
    Write-Host "You can now open Word and press Alt+F11 to verify.`n" -ForegroundColor Gray
}
catch {
    Write-Host "`n[ERROR] Failed: $_" -ForegroundColor Red
    exit 1
}
finally {
    if ($word) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
