<#
.SYNOPSIS
    Idempotently create and (optionally) pin an application shortcut to the Windows taskbar.
.DESCRIPTION
    Generic utility (no product coupling) that:
      - Creates or updates a Start Menu shortcut (.lnk)
      - Attempts to detect if an equivalent shortcut is already pinned
      - Pins the shortcut using Shell verbs (best-effort) if not already pinned
      - Can preload Windows clipboard history with text snippets

    Designed to run UN-ELEVATED. When elevated, the pin verb usually disappears; the script aborts unless -AllowElevated.

.PARAMETER ShortcutName
    Display name for the shortcut (.lnk). Default: 'MyApp'.
.PARAMETER TargetExecutable
    Target executable path or resolvable filename (e.g. notepad.exe). Required for correct shortcut creation.
.PARAMETER Arguments
    Optional command-line arguments passed to the target.
.PARAMETER WorkingDirectory
    Working directory for the process launched by the shortcut. Defaults to directory of TargetExecutable if not supplied.
.PARAMETER IconPath
    Path to an .ico or compatible resource (exe/dll) for the shortcut icon. If omitted, target's default icon is used.
.PARAMETER Force
    Recreate the shortcut even if it already exists (does NOT guarantee unpin first).
.PARAMETER ClipboardListPath
    Path to UTF-8 text file: each non-empty line becomes a clipboard history entry (Win+V) in order.
.PARAMETER NoClipboardLoad
    Skip clipboard history seeding even if a list file exists.
.PARAMETER PreserveClipboard
    Restore the original clipboard contents after seeding entries.
.PARAMETER ClipboardSetDelayMs
    Delay (milliseconds) between Set-Clipboard operations (default 120). Increase if entries are missing.
.PARAMETER AllowElevated
    Permit execution while elevated (pin may silently fail).
.PARAMETER DryRun
    Show operations without modifying filesystem or clipboard.

.EXAMPLE
    PS> ./pin-taskbar.ps1 -TargetExecutable notepad.exe -ShortcutName "My Notepad"

.EXAMPLE
    PS> ./pin-taskbar.ps1 -TargetExecutable C:\Tools\myapp.exe -Arguments "--fast" -IconPath C:\Icons\myapp.ico -ClipboardListPath .\snippets.txt

.EXAMPLE
    PS> ./pin-taskbar.ps1 -TargetExecutable notepad.exe -NoClipboardLoad -DryRun

.EXIT CODES
    0  Success (already pinned or pinned now)
    1  Target executable not found
    2  Pin verb unavailable (manual pin needed)
    3  Missing required parameter(s)
    10 Elevated and not allowed
    11 Dry run completed (no pin attempted)
    99 Unexpected error

.LICENSE
    SPDX-License-Identifier: MIT
#>
[CmdletBinding()]param(
    [string]$ShortcutName = 'MyApp',
    [Parameter(Mandatory=$true)][string]$TargetExecutable,
    [string]$Arguments = '',
    [string]$WorkingDirectory = '',
    [string]$IconPath = '',
    [switch]$Force,
    [string]$ClipboardListPath = '',
    [switch]$NoClipboardLoad,
    [switch]$PreserveClipboard,
    [int]$ClipboardSetDelayMs = 120,
    [switch]$AllowElevated,
    [switch]$DryRun,
    # Dump verb list and attempt multiple localized patterns / canonical verb names
    [switch]$PinDebug
)

$ErrorActionPreference = 'Stop'

function Write-Info($m){Write-Host "[INFO] $m" -ForegroundColor Cyan}
function Write-Warn($m){Write-Host "[WARN] $m" -ForegroundColor Yellow}
function Write-Err($m){Write-Host "[ERROR] $m" -ForegroundColor Red}
function Write-Success($m){Write-Host "[SUCCESS] $m" -ForegroundColor Green}

# --- Elevation Guard ---
try { $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch { $isAdmin=$false }
if($isAdmin -and -not $AllowElevated){ Write-Err 'Script is elevated. Re-run in a normal PowerShell window or use -AllowElevated.'; exit 10 }

# --- Basic Validation ---
if(-not $TargetExecutable){ Write-Err 'TargetExecutable is required.'; exit 3 }
$resolvedTarget = $null
if(Test-Path -LiteralPath $TargetExecutable){ $resolvedTarget = (Resolve-Path -LiteralPath $TargetExecutable).Path } else {
    $cmd = Get-Command $TargetExecutable -ErrorAction SilentlyContinue
    if($cmd){ $resolvedTarget = $cmd.Source }
}
if(-not $resolvedTarget){ Write-Err "Target executable '$TargetExecutable' not found."; exit 1 }
if(-not $WorkingDirectory){ $WorkingDirectory = Split-Path -Parent $resolvedTarget }

# --- Paths ---
$startMenuDir = Join-Path $env:APPDATA 'Microsoft\\Windows\\Start Menu\\Programs'
$shortcutPath = Join-Path $startMenuDir ("$ShortcutName.lnk")

# --- Clipboard Loader ---
function Invoke-ClipboardBatchLoad {
    param([string]$ListPath,[switch]$Preserve,[int]$DelayMs=120,[switch]$Dry)
    if($Dry){ Write-Info "[DRY] Would load clipboard lines from $ListPath"; return }
    if(-not (Test-Path -LiteralPath $ListPath)){ Write-Warn "Clipboard list '$ListPath' not found."; return }
    $lines = Get-Content -LiteralPath $ListPath -ErrorAction SilentlyContinue | ForEach-Object { $_.TrimEnd() } | Where-Object { $_ -ne '' }
    if(-not $lines){ Write-Warn 'Clipboard list has no usable lines.'; return }
    $original = $null
    if($Preserve){ try { $original = Get-Clipboard -Raw -EA SilentlyContinue } catch {} }
    $i=0
    foreach($l in $lines){
        try { Set-Clipboard -Value $l -EA Stop; Start-Sleep -Milliseconds $DelayMs; $i++ } catch { Write-Warn "Failed line #$($i+1): $($_.Exception.Message)" }
    }
    if($Preserve -and $null -ne $original){ try { Set-Clipboard -Value $original } catch { Write-Warn 'Could not restore original clipboard.' } }
    Write-Info "Loaded $i clipboard entrie(s)."
}

# --- Create / Update Shortcut ---
function New-OrUpdateShortcut {
    param([string]$Path,[string]$Target,[string]$Args,[string]$WorkDir,[string]$Icon,[switch]$Dry)
    if($Dry){ Write-Info "[DRY] Would create/update shortcut $Path => $Target $Args"; return }
    if(-not (Test-Path $startMenuDir)){ New-Item -ItemType Directory -Path $startMenuDir | Out-Null }
    $shell = New-Object -ComObject WScript.Shell
    $sc = $shell.CreateShortcut($Path)
    $sc.TargetPath = $Target
    if($Args){ $sc.Arguments = $Args }
    if($WorkDir){ $sc.WorkingDirectory = $WorkDir }
    if($Icon -and (Test-Path $Icon)){ $sc.IconLocation = $Icon }
    $sc.WindowStyle = 1
    $sc.Description = "Pinned shortcut for $ShortcutName"
    $sc.Save()
}

# --- Detect Pinned ---
function Test-TaskbarPinned {
    param([string]$Name,[string]$TargetPath)
    $pinnedDir = Join-Path $env:APPDATA 'Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar'
    if(-not (Test-Path $pinnedDir)){ return $false }
    Get-ChildItem -LiteralPath $pinnedDir -Filter '*.lnk' -ErrorAction SilentlyContinue | ForEach-Object {
        try { $shell = New-Object -ComObject WScript.Shell; $sc=$shell.CreateShortcut($_.FullName); if($sc.TargetPath -and (Split-Path $sc.TargetPath -Leaf) -ieq (Split-Path $TargetPath -Leaf) -and $_.BaseName -like "*$Name*") { throw 'PINNED' } } catch { if($_.Exception.Message -eq 'PINNED'){ throw } }
    }; return $false
} trap { if($_.Exception.Message -eq 'PINNED'){ return $true } else { continue } }

# --- Pin (best-effort) ---
function Invoke-PinToTaskbar { param([string]$Path,[switch]$Dry,[switch]$Debug)
    if($Dry){ Write-Info "[DRY] Would attempt pin of $Path"; return $true }
    $file = Get-Item -LiteralPath $Path -EA Stop
    $shell = New-Object -ComObject Shell.Application
    $dir = $shell.Namespace($file.DirectoryName)
    $item = $dir.ParseName($file.Name)
    if(-not $item){ throw 'Shell item not found.' }
    $allVerbs = @(); try { $allVerbs = @($item.Verbs()) } catch {}
    if($Debug){
        Write-Info 'Available verbs:'
        $allVerbs | ForEach-Object { Write-Host ('  - ' + $_.Name) }
    }
    # Localized / variant patterns
    $patterns = @(
        'Pin to taskbar','Fixar na barra de tarefas','Fixar .*barra de tarefas','Anclar a la barra de tareas',
        'Anheften an Taskleiste','Aggiungi.*barra delle applicazioni','Adicionar.*barra de tarefas'
    )
    foreach($p in $patterns){
        $verb = $allVerbs | Where-Object { $_.Name -match $p }
        if($verb){
            Write-Info "Using verb pattern match: '$($verb.Name)'"
            $verb.DoIt(); Start-Sleep -Milliseconds 500; return $true
        }
    }
    # Attempt canonical hidden verb invocation (may not appear in enumeration)
    $canonAttempts = 'taskbarpin','TaskbarPin'
    foreach($c in $canonAttempts){
        try { $item.InvokeVerb($c); Start-Sleep -Milliseconds 500; Write-Info "Attempted canonical verb '$c'"; return $true } catch { }
    }
    if($Debug){ Write-Warn 'No matching verb found after all attempts.' }
    return $false
}

try {
    Write-Info "Preparing shortcut: $ShortcutName"

    if($DryRun){ Write-Warn 'DRY RUN mode enabled (no writes/pins).'
    }

    # Clipboard seeding (optional)
    if(-not $NoClipboardLoad){
        if(-not $ClipboardListPath){ $auto = Join-Path (Split-Path -Parent $PSCommandPath) 'snippets.txt'; if(Test-Path $auto){ $ClipboardListPath=$auto } }
        if($ClipboardListPath){ Write-Info "Seeding clipboard from $ClipboardListPath"; Invoke-ClipboardBatchLoad -ListPath $ClipboardListPath -Preserve:$PreserveClipboard -DelayMs $ClipboardSetDelayMs -Dry:$DryRun }
        else { Write-Info 'No snippets file found; skipping clipboard load.' }
    } else { Write-Info 'Clipboard load suppressed.' }

    if((Test-Path $shortcutPath) -and $Force){ if($DryRun){ Write-Info "[DRY] Would remove existing $shortcutPath" } else { Remove-Item -LiteralPath $shortcutPath -Force -EA SilentlyContinue } }

    if(-not (Test-Path $shortcutPath)) { New-OrUpdateShortcut -Path $shortcutPath -Target $resolvedTarget -Args $Arguments -WorkDir $WorkingDirectory -Icon $IconPath -Dry:$DryRun } else { Write-Info 'Shortcut already exists.' }

    $alreadyPinned = $false
    try { $alreadyPinned = Test-TaskbarPinned -Name $ShortcutName -TargetPath $resolvedTarget } catch { $alreadyPinned = $true }

    if($alreadyPinned -and -not $Force){ Write-Info 'Already pinned (idempotent).'; if($DryRun){ exit 11 } else { exit 0 } }

    if($Force -and $alreadyPinned){ Write-Info 'Force specified: attempting re-pin (unpin not automated).' }

    Write-Info 'Attempting pin...'
    $pinOk = Invoke-PinToTaskbar -Path $shortcutPath -Dry:$DryRun -Debug:$PinDebug
    if($DryRun){ Write-Info '[DRY] Pin attempt simulated.'; exit 11 }
    if($pinOk){ Write-Success "Taskbar pin ensured for '$ShortcutName'"; exit 0 }
    else { Write-Warn 'Pin verb unavailable; pin manually via Start Menu context menu.'; exit 2 }
}
catch {
    Write-Err $_.Exception.Message
    if($DryRun){ exit 11 } else { exit 99 }
}
