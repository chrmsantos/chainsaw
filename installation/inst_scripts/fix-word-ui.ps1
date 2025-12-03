# =============================================================================
# Fix Word UI Customizations
# =============================================================================
# This script re-imports Word UI customizations (Ribbon, QAT, etc.)

$ErrorActionPreference = "Stop"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Word UI Customization Fix Tool" -ForegroundColor White
Write-Host "========================================`n" -ForegroundColor Cyan

# Paths
$exportedConfigPath = Join-Path $PSScriptRoot "..\exported-config"
$localOfficePath = Join-Path $env:LOCALAPPDATA "Microsoft\Office"
$roamingOfficePath = Join-Path $env:APPDATA "Microsoft\Office"

Write-Host "[1/4] Closing all Word processes..." -ForegroundColor Yellow
Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 3

Write-Host "[2/4] Removing lock files..." -ForegroundColor Yellow
Remove-Item (Join-Path $env:APPDATA "Microsoft\Office\*.officeUI.lock") -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path $env:LOCALAPPDATA "Microsoft\Office\*.officeUI.lock") -Force -ErrorAction SilentlyContinue

Write-Host "[3/4] Backing up existing UI files..." -ForegroundColor Yellow
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Backup existing files
Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Office" -Filter "Word.officeUI" -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    $backupPath = $_.FullName + ".backup_$timestamp"
    Copy-Item $_.FullName $backupPath -Force
    Write-Host "  ✓ Backed up: $($_.FullName)" -ForegroundColor Gray
}

Get-ChildItem "$env:APPDATA\Microsoft\Office" -Filter "Word.officeUI" -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    $backupPath = $_.FullName + ".backup_$timestamp"
    Copy-Item $_.FullName $backupPath -Force
    Write-Host "  ✓ Backed up: $($_.FullName)" -ForegroundColor Gray
}

Write-Host "[4/4] Importing new UI customizations..." -ForegroundColor Yellow

# Import Ribbon customization
$ribbonSource = Join-Path $exportedConfigPath "OfficeCustomUI\Word.officeUI"
if (Test-Path $ribbonSource) {
    # Copy to Local AppData (primary location)
    $ribbonDestLocal = Join-Path $localOfficePath "Word.officeUI"
    Copy-Item $ribbonSource $ribbonDestLocal -Force
    Write-Host "  ✓ Imported Ribbon to Local: $ribbonDestLocal" -ForegroundColor Green
    
    # Copy to Roaming AppData (backup location)
    $ribbonDestRoaming = Join-Path $roamingOfficePath "Word.officeUI"
    Copy-Item $ribbonSource $ribbonDestRoaming -Force
    Write-Host "  ✓ Imported Ribbon to Roaming: $ribbonDestRoaming" -ForegroundColor Green
}
else {
    Write-Host "  [WARNING] Ribbon file not found: $ribbonSource" -ForegroundColor Yellow
}

# Import other UI customizations from RibbonCustomization folder
$ribbonCustomSource = Join-Path $exportedConfigPath "RibbonCustomization\Word.officeUI"
if (Test-Path $ribbonCustomSource) {
    $ribbonCustomDest = Join-Path $localOfficePath "RibbonCustomization\Word.officeUI"
    $ribbonCustomDir = Split-Path $ribbonCustomDest -Parent
    
    if (-not (Test-Path $ribbonCustomDir)) {
        New-Item -Path $ribbonCustomDir -ItemType Directory -Force | Out-Null
    }
    
    Copy-Item $ribbonCustomSource $ribbonCustomDest -Force
    Write-Host "  ✓ Imported Custom Ribbon: $ribbonCustomDest" -ForegroundColor Green
}

Write-Host "`n[SUCCESS] Word UI customizations reinstalled!" -ForegroundColor Green
Write-Host "`nNOTE: You must restart Word to see the changes.`n" -ForegroundColor Cyan

# Summary
Write-Host "Files installed:" -ForegroundColor White
Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Office" -Filter "Word.officeUI" -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "  → $($_.FullName)" -ForegroundColor Gray
}
Get-ChildItem "$env:APPDATA\Microsoft\Office" -Filter "Word.officeUI" -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "  → $($_.FullName)" -ForegroundColor Gray
}
Write-Host ""
