<#!
.SYNOPSIS
  Counts lines of code (LOC) for VBA modules in /src.
.DESCRIPTION
  By default counts only active modular source files and excludes legacy / deprecated
  large modules (modMain.bas, legacy snapshot if present). Use -IncludeLegacy to
  include every .bas file.
.PARAMETER IncludeLegacy
  Include legacy / deprecated modules in the LOC tally.
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts/count-loc.ps1
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts/count-loc.ps1 -IncludeLegacy
.NOTES
  Safe to run on Windows PowerShell 5.1 or newer.
#>
param(
    [switch]$IncludeLegacy
)

$ErrorActionPreference = 'Stop'
$srcPath = Join-Path $PSScriptRoot '..' | Join-Path -ChildPath 'src'
if(-not (Test-Path $srcPath)) { Write-Error "Source path not found: $srcPath" }

$allFiles = Get-ChildItem -Path $srcPath -Filter *.bas -File | Sort-Object Name
$legacyNames = @('modMain.bas','legacy_chainsaw_snapshot.bas')

if($IncludeLegacy) {
    $target = $allFiles
} else {
    $target = $allFiles | Where-Object { $legacyNames -notcontains $_.Name }
}

if(-not $target) { Write-Warning 'No .bas files found to process.'; return }

$total = 0
$rows = foreach($f in $target){
    $lineCount = (Get-Content -Raw -Encoding UTF8 $f.FullName | Measure-Object -Line).Lines
    $total += $lineCount
    [pscustomobject]@{ File=$f.Name; Lines=$lineCount }
}

$rows | Format-Table -AutoSize
Write-Host ('TOTAL_ACTIVE_LINES=' + $total)

if(-not $IncludeLegacy){
    $legacyTotal = 0
    foreach($l in $allFiles | Where-Object { $legacyNames -contains $_.Name }){
        $legacyTotal += (Get-Content -Raw -Encoding UTF8 $l.FullName | Measure-Object -Line).Lines
    }
    Write-Host ('LEGACY_LINES=' + $legacyTotal)
    Write-Host ('GRAND_TOTAL=' + ($total + $legacyTotal))
}

exit 0
