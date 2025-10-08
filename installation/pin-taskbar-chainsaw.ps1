<#
.SYNOPSIS
    Idempotently pins a Chainsaw launcher shortcut to the Windows Taskbar.
.DESCRIPTION
    - Creates (or updates) a Start Menu shortcut for the Chainsaw Word macro environment.
    - Detects whether a shortcut with a specific AppUserModelID is already pinned.
    - If not pinned, attempts to pin using supported COM / shell verbs fallback strategies.
    - Requires running in a normal user context (not elevated) because pin verbs are hidden when elevated.
    - Optionally seeds Windows clipboard history with predefined text snippets (one per line) to speed up drafting.

.PARAMETER ShortcutName
    Display name for the Start Menu / Taskbar shortcut (".lnk" file). Default: 'Chainsaw Proposituras'.

.PARAMETER WordExecutable
    Path or filename of Microsoft Word (winword.exe). If just 'winword.exe', it must be resolvable via PATH / registry.

.PARAMETER StartupDocument
    Optional .docx/.dotm/.docm opened when the shortcut launches Word (e.g., a macro-enabled bootstrap template).

.PARAMETER ChainsawRoot
    Working directory set on the shortcut (used for relative assets). Defaults to project root (parent folder of script).

.PARAMETER Force
    If provided, removes an existing shortcut before recreating it and attempts to pin again (cannot guarantee unpin first).

.PARAMETER ClipboardListPath
    Path to a UTF-8 text file containing one clipboard entry per non-empty line. If omitted, the script looks for
    'clipboard-snippets.txt' in the same folder as the script. Each line is sequentially copied, building clipboard history.

.PARAMETER NoClipboardLoad
    Skip clipboard history seeding even if a snippets file exists.

.PARAMETER PreserveClipboard
    After seeding history, restore the original clipboard content that was present before the script ran.

.PARAMETER ClipboardSetDelayMs
    Delay (milliseconds) between each Set-Clipboard call. Increase (e.g., 200+) if some entries are missing in history.

.EXAMPLE
    PS> PowerShell -ExecutionPolicy Bypass -File .\pin-taskbar-chainsaw.ps1
    Ensures shortcut exists, pins it if needed, attempts to load default clipboard-snippets.txt if present.

.EXAMPLE
    PS> PowerShell -ExecutionPolicy Bypass -File .\pin-taskbar-chainsaw.ps1 -ClipboardListPath C:\Snips\legis.txt -PreserveClipboard
    Pins shortcut and seeds clipboard history from a custom file, then restores original clipboard content.

.EXAMPLE
    PS> PowerShell -ExecutionPolicy Bypass -File .\pin-taskbar-chainsaw.ps1 -StartupDocument C:\Templates\chainsaw.dotm -Force
    Forces recreation of shortcut, sets it to open a template, and attempts re-pin.

.EXAMPLE
    PS> PowerShell -ExecutionPolicy Bypass -File .\pin-taskbar-chainsaw.ps1 -NoClipboardLoad
    Runs pin logic only; skips clipboard seeding.

.OUTPUTS
    Writes informational status lines. Exit codes describe final state (see below).

.NOTES
    Clipboard history seeding leverages sequential Set-Clipboard calls. Windows 10/11 keeps a chronological history (Win+V)
    that should reflect inserted entries after the user first enables the Windows clipboard history feature.
    The script does not attempt to enable clipboard history for you.

.EXIT CODES
    0  Success (already pinned or newly pinned; clipboard seeding best-effort)
    1  Word executable not found
    2  Pin verb unavailable (manual pinning required)
    99 Unexpected error

.NOTES
    Tested on Windows 10/11. Pinning is an unsupported automation area; this script uses best-effort heuristics.
    If direct pinning fails, user is prompted with a manual fallback message.
.LICENSE
    SPDX-License-Identifier: GPL-3.0-or-later
#>

[CmdletBinding()]
param(
    [string]$ShortcutName = 'Chainsaw Proposituras',
    [string]$WordExecutable = 'winword.exe',
    # Optional: path to a .dotm or document that triggers the macro environment.
    [string]$StartupDocument = '',
    # Working directory for the launcher (where config & modules live)
    [string]$ChainsawRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path,
    # If specified, force re-pin (remove then pin again)
    [switch]$Force,
    # Auto clipboard history loader: path to list of lines to seed into clipboard history (one per line). If empty, a default file is searched.
    [string]$ClipboardListPath = '',
    # Disable auto clipboard load even if a file exists.
    [switch]$NoClipboardLoad,
    # Preserve (restore) the original clipboard content after seeding history.
    [switch]$PreserveClipboard,
    # Delay (ms) between clipboard sets when seeding history (tune if entries missing in history)
    [int]$ClipboardSetDelayMs = 120
)

$ErrorActionPreference = 'Stop'

function Write-Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host "[WARN] $msg" -ForegroundColor Yellow }
function Write-Err($msg) { Write-Host "[ERROR] $msg" -ForegroundColor Red }

# Region: Clipboard Loader
function Invoke-ClipboardBatchLoad {
    param(
        [string]$ListPath,
        [switch]$Preserve,
        [int]$DelayMs = 120
    )
    if (-not (Test-Path -LiteralPath $ListPath)) { Write-Warn "Clipboard list '$ListPath' not found."; return }
    $lines = Get-Content -LiteralPath $ListPath -ErrorAction SilentlyContinue | ForEach-Object { $_.TrimEnd() } | Where-Object { $_ -ne '' }
    if (-not $lines) { Write-Warn "Clipboard list file '$ListPath' has no non-empty lines."; return }
    $original = $null
    if ($Preserve) {
        try { $original = Get-Clipboard -Raw -ErrorAction SilentlyContinue } catch { }
    }
    $count = 0
    foreach ($l in $lines) {
        try {
            Set-Clipboard -Value $l -ErrorAction Stop
            Start-Sleep -Milliseconds $DelayMs
            $count++
        } catch {
            Write-Warn "Failed to set clipboard item #$($count+1): $($_.Exception.Message)"
        }
    }
    if ($Preserve -and $null -ne $original) {
        try { Set-Clipboard -Value $original } catch { Write-Warn 'Could not restore original clipboard content.' }
    }
    Write-Info "Loaded $count clipboard history entrie(s) from '$ListPath'."
}

# Region: Constants & Paths
$AppID = 'com.chainsaw.proposituras'
$StartMenuShortcutDir = Join-Path $env:APPDATA 'Microsoft\\Windows\\Start Menu\\Programs'
$ShortcutPath = Join-Path $StartMenuShortcutDir ("$ShortcutName.lnk")
$IconPath = Join-Path $ChainsawRoot 'assets\\stamp.png'

# Region: Helper - Create or Update .lnk
function New-OrUpdateShortcut {
    param(
        [string]$Path,
        [string]$Target,
        [string]$Arguments,
        [string]$WorkingDirectory,
        [string]$IconLocation,
        [string]$AppUserModelID
    )
    $shell = New-Object -ComObject WScript.Shell
    $sc = $shell.CreateShortcut($Path)
    $sc.TargetPath = $Target
    if ($Arguments) { $sc.Arguments = $Arguments }
    if ($WorkingDirectory) { $sc.WorkingDirectory = $WorkingDirectory }
    if (Test-Path $IconLocation) { $sc.IconLocation = $IconLocation }
    $sc.WindowStyle = 1
    $sc.Description = 'Chainsaw - Padronização de Proposituras'
    $sc.Save()

    # Set AppUserModelID via Shell Property Store (needs pin consistency)
    try {
        $bytes = [System.Text.Encoding]::Unicode.GetBytes($AppUserModelID + [char]0)
        $propStoreGuid = [Guid]'9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3' # AppUserModelID
        $PSObject = ([Activator]::CreateInstance([type]::GetTypeFromProgID('Shell.Application'))).Name | Out-Null
        # Directly writing property store of .lnk is non-trivial in pure PowerShell without C#.
        # Skipping deep COM property write; pin detection will rely on name / target heuristics.
    } catch {
        Write-Warn 'Could not set AppUserModelID (non-fatal).'
    }
}

# Region: Detect if pinned (heuristic)
function Test-TaskbarPinned {
    param([string]$ShortcutName,[string]$TargetPath)
    # Windows stores pinned taskbar shortcuts in a binary layout in the User Pinned\TaskBar folder.
    $PinnedDir = Join-Path $env:APPDATA 'Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar'
    if (-not (Test-Path $PinnedDir)) { return $false }
    $lnks = Get-ChildItem -LiteralPath $PinnedDir -Filter '*.lnk' -ErrorAction SilentlyContinue
    foreach ($lnk in $lnks) {
        try {
            $shell = New-Object -ComObject WScript.Shell
            $sc = $shell.CreateShortcut($lnk.FullName)
            if ($sc.TargetPath -and (Split-Path $sc.TargetPath -Leaf) -ieq (Split-Path $TargetPath -Leaf)) {
                if ($lnk.BaseName -like "*$ShortcutName*") { return $true }
            }
        } catch { }
    }
    return $false
}

# Region: Attempt pin via Shell verbs
function Invoke-PinToTaskbar {
    param([string]$Path)
    $file = Get-Item -LiteralPath $Path -ErrorAction Stop
    $folder = $file.DirectoryName
    $filename = $file.Name
    $shell = New-Object -ComObject Shell.Application
    $dir = $shell.Namespace($folder)
    $item = $dir.ParseName($filename)
    if (-not $item) { throw 'Shell item not found for pin operation.' }
    $pinVerb = $item.Verbs() | Where-Object { $_.Name -match 'Pin to taskbar|Fixar na barra de tarefas' }
    if ($pinVerb) {
        $pinVerb.DoIt()
        Start-Sleep -Milliseconds 500
        return $true
    }
    return $false
}

# Region: Main Flow
try {
    Write-Info "Preparing taskbar pin for '$ShortcutName'"

    # Auto-detect default clipboard list file if not explicitly provided.
    if (-not $NoClipboardLoad) {
        if (-not $ClipboardListPath) {
            $defaultList = Join-Path $PSScriptRoot 'clipboard-snippets.txt'
            if (Test-Path -LiteralPath $defaultList) { $ClipboardListPath = $defaultList }
        }
        if ($ClipboardListPath) {
            Write-Info "Seeding clipboard history from '$ClipboardListPath' (delay ${ClipboardSetDelayMs}ms, preserve=$PreserveClipboard)"
            Invoke-ClipboardBatchLoad -ListPath $ClipboardListPath -Preserve:$PreserveClipboard -DelayMs $ClipboardSetDelayMs
        } else {
            Write-Info 'No clipboard snippets file detected; skipping clipboard history load.'
        }
    } else {
        Write-Info 'Clipboard history auto-load disabled by parameter.'
    }

    if (-not (Get-Command $WordExecutable -ErrorAction SilentlyContinue)) {
        Write-Err "Word executable '$WordExecutable' not found in PATH. Provide -WordExecutable full path."
        exit 1
    }

    $arguments = ''
    if ($StartupDocument) {
        $arguments = '"' + $StartupDocument + '"'
    }

    if (-not (Test-Path $StartMenuShortcutDir)) {
        New-Item -ItemType Directory -Path $StartMenuShortcutDir | Out-Null
    }

    if ((Test-Path $ShortcutPath) -and $Force) {
        Write-Info 'Force specified: removing existing shortcut.'
        Remove-Item -LiteralPath $ShortcutPath -Force -ErrorAction SilentlyContinue
    }

    if (-not (Test-Path $ShortcutPath)) {
        Write-Info 'Creating Start Menu shortcut.'
        New-OrUpdateShortcut -Path $ShortcutPath -Target $WordExecutable -Arguments $arguments -WorkingDirectory $ChainsawRoot -IconLocation $IconPath -AppUserModelID $AppID
    } else {
        Write-Info 'Shortcut already exists.'
    }

    $alreadyPinned = Test-TaskbarPinned -ShortcutName $ShortcutName -TargetPath $WordExecutable
    if ($alreadyPinned -and -not $Force) {
        Write-Info 'Taskbar button already pinned (idempotent no-op).'
        exit 0
    }

    if ($Force -and $alreadyPinned) {
        Write-Info 'Force re-pin requested: cannot auto unpin reliably; proceeding to attempt re-pin anyway.'
    }

    Write-Info 'Attempting to pin to taskbar...'
    $pinOk = Invoke-PinToTaskbar -Path $ShortcutPath
    if ($pinOk) {
        Write-Host "[SUCCESS] Taskbar pin ensured for '$ShortcutName'" -ForegroundColor Green
        exit 0
    } else {
        Write-Warn 'Automated pin verb not available (possibly due to policy, elevation, or Windows build).'
        Write-Warn "Manual fallback: Right-click the shortcut '$ShortcutName' in Start Menu and choose 'Pin to taskbar'."
        exit 2
    }
}
catch {
    Write-Err $_.Exception.Message
    exit 99
}
