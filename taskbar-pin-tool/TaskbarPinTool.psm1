function New-ApplicationTaskbarPin {
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
        [switch]$DryRun
    )
    $ErrorActionPreference = 'Stop'
    function Write-Info($m){Write-Host "[INFO] $m" -ForegroundColor Cyan}
    function Write-Warn($m){Write-Host "[WARN] $m" -ForegroundColor Yellow}
    function Write-Err($m){Write-Host "[ERROR] $m" -ForegroundColor Red}
    function Write-Success($m){Write-Host "[SUCCESS] $m" -ForegroundColor Green}
    try { $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch { $isAdmin=$false }
    if($isAdmin -and -not $AllowElevated){ Write-Err 'Elevated session detected; rerun without admin or pass -AllowElevated.'; return }
    if(-not $TargetExecutable){ Write-Err 'TargetExecutable is required.'; return }
    $resolvedTarget = $null
    if(Test-Path -LiteralPath $TargetExecutable){ $resolvedTarget = (Resolve-Path -LiteralPath $TargetExecutable).Path } else { $cmd = Get-Command $TargetExecutable -ErrorAction SilentlyContinue; if($cmd){ $resolvedTarget = $cmd.Source } }
    if(-not $resolvedTarget){ Write-Err "Target executable '$TargetExecutable' not found."; return }
    if(-not $WorkingDirectory){ $WorkingDirectory = Split-Path -Parent $resolvedTarget }
    $startMenuDir = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'
    $shortcutPath = Join-Path $startMenuDir ("$ShortcutName.lnk")
    function Invoke-ClipboardBatchLoad {
        param([string]$ListPath,[switch]$Preserve,[int]$DelayMs=120,[switch]$Dry)
        if($Dry){ Write-Info "[DRY] Would load clipboard lines from $ListPath"; return }
        if(-not (Test-Path -LiteralPath $ListPath)){ Write-Warn "Clipboard list '$ListPath' not found."; return }
        $lines = Get-Content -LiteralPath $ListPath -EA SilentlyContinue | ForEach-Object { $_.TrimEnd() } | Where-Object { $_ -ne '' }
        if(-not $lines){ Write-Warn 'Clipboard list has no usable lines.'; return }
        $orig=$null; if($Preserve){ try{$orig=Get-Clipboard -Raw -EA SilentlyContinue}catch{} }
        $i=0; foreach($l in $lines){ try{ Set-Clipboard -Value $l -EA Stop; Start-Sleep -Milliseconds $DelayMs; $i++ } catch { Write-Warn "Failed line #$($i+1): $($_.Exception.Message)" } }
        if($Preserve -and $null -ne $orig){ try{ Set-Clipboard -Value $orig } catch { Write-Warn 'Could not restore original clipboard.' } }
        Write-Info "Loaded $i clipboard entrie(s)."
    }
    function New-OrUpdateShortcut {
        param([string]$Path,[string]$Target,[string]$CmdArgs,[string]$WorkDir,[string]$Icon,[switch]$Dry)
        if($Dry){ Write-Info "[DRY] Would create/update shortcut $Path => $Target $CmdArgs"; return }
        if(-not (Test-Path $startMenuDir)){ New-Item -ItemType Directory -Path $startMenuDir | Out-Null }
        $shell = New-Object -ComObject WScript.Shell
        $sc = $shell.CreateShortcut($Path)
        $sc.TargetPath = $Target
    if($CmdArgs){ $sc.Arguments = $CmdArgs }
        if($WorkDir){ $sc.WorkingDirectory = $WorkDir }
        if($Icon -and (Test-Path $Icon)){ $sc.IconLocation = $Icon }
        $sc.WindowStyle = 1
        $sc.Description = "Pinned shortcut for $ShortcutName"
        $sc.Save()
    }
    function Test-TaskbarPinned { param([string]$Name,[string]$TargetPath)
        $pinnedDir = Join-Path $env:APPDATA 'Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar'
        if(-not (Test-Path $pinnedDir)){ return $false }
        Get-ChildItem -LiteralPath $pinnedDir -Filter '*.lnk' -EA SilentlyContinue | ForEach-Object { try { $shell = New-Object -ComObject WScript.Shell; $sc=$shell.CreateShortcut($_.FullName); if($sc.TargetPath -and (Split-Path $sc.TargetPath -Leaf) -ieq (Split-Path $TargetPath -Leaf) -and $_.BaseName -like "*$Name*") { throw 'PINNED' } } catch { if($_.Exception.Message -eq 'PINNED'){ throw } } }; return $false
    } trap { if($_.Exception.Message -eq 'PINNED'){ return $true } else { continue } }
    function Invoke-PinToTaskbar { param([string]$Path,[switch]$Dry)
        if($Dry){ Write-Info "[DRY] Would attempt pin of $Path"; return $true }
        $file = Get-Item -LiteralPath $Path -EA Stop
        $shell = New-Object -ComObject Shell.Application
        $dir = $shell.Namespace($file.DirectoryName)
        $item = $dir.ParseName($file.Name)
        if(-not $item){ throw 'Shell item not found.' }
        $verb = $item.Verbs() | Where-Object { $_.Name -match 'Pin to taskbar|Fixar na barra de tarefas' }
        if($verb){ $verb.DoIt(); Start-Sleep -Milliseconds 500; return $true }
        return $false
    }
    Write-Info "Preparing shortcut: $ShortcutName"
    if(-not $NoClipboardLoad){ if(-not $ClipboardListPath){ $auto = Join-Path (Split-Path -Parent $PSCommandPath) 'snippets.txt'; if(Test-Path $auto){ $ClipboardListPath=$auto } }
        if($ClipboardListPath){ Write-Info "Seeding clipboard from $ClipboardListPath"; Invoke-ClipboardBatchLoad -ListPath $ClipboardListPath -Preserve:$PreserveClipboard -DelayMs $ClipboardSetDelayMs -Dry:$DryRun } else { Write-Info 'No snippets file found; skipping clipboard load.' } }
    if((Test-Path $shortcutPath) -and $Force){ if($DryRun){ Write-Info "[DRY] Would remove existing $shortcutPath" } else { Remove-Item -LiteralPath $shortcutPath -Force -EA SilentlyContinue } }
    if(-not (Test-Path $shortcutPath)) { New-OrUpdateShortcut -Path $shortcutPath -Target $resolvedTarget -CmdArgs $Arguments -WorkDir $WorkingDirectory -Icon $IconPath -Dry:$DryRun } else { Write-Info 'Shortcut already exists.' }
    $alreadyPinned=$false; try { $alreadyPinned = Test-TaskbarPinned -Name $ShortcutName -TargetPath $resolvedTarget } catch { $alreadyPinned=$true }
    if($alreadyPinned -and -not $Force){ Write-Info 'Already pinned (idempotent).'; return }
    if($Force -and $alreadyPinned){ Write-Info 'Force specified: attempting re-pin (unpin not automated).' }
    Write-Info 'Attempting pin...'
    $pinOk = Invoke-PinToTaskbar -Path $shortcutPath -Dry:$DryRun
    if($DryRun){ Write-Info '[DRY] Pin attempt simulated.'; return }
    if($pinOk){ Write-Success "Taskbar pin ensured for '$ShortcutName'" } else { Write-Warn 'Pin verb unavailable; pin manually via Start Menu context menu.' }
}
Export-ModuleMember -Function New-ApplicationTaskbarPin
