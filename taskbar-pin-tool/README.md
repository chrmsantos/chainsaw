# Taskbar Pin Tool

Generic PowerShell utility to create/update and (attempt to) pin a Start Menu shortcut to the Windows taskbar, plus optional clipboard history seeding.

## Features
- Idempotent shortcut creation (.lnk) in the user's Start Menu
- Best-effort taskbar pin attempt (Shell verb based)
- Clipboard history preload from text file (Win+V integration)
- Dry-run mode (preview operations)
- Elevation guard (refuses to run elevated unless overridden)
- Force recreation of shortcut

## Requirements
- Windows 10 or 11
- PowerShell 5.1+ (compatible with pwsh Core, though pin verb reliability varies)
- Clipboard history feature enabled (if using snippets)

## Usage
```powershell
# Basic
PowerShell -ExecutionPolicy Bypass -File .\pin-taskbar.ps1 -TargetExecutable notepad.exe

# With custom icon and arguments
PowerShell -ExecutionPolicy Bypass -File .\pin-taskbar.ps1 -TargetExecutable C:\Apps\myapp.exe -Arguments "--fast" -IconPath C:\Icons\myapp.ico

# Seed clipboard history (auto-detects snippets.txt if present)
PowerShell -ExecutionPolicy Bypass -File .\pin-taskbar.ps1 -TargetExecutable notepad.exe -ClipboardListPath .\snippets.txt -PreserveClipboard

# Dry run only
PowerShell -ExecutionPolicy Bypass -File .\pin-taskbar.ps1 -TargetExecutable notepad.exe -DryRun
```

## Parameters (Summary)
| Parameter | Description |
|-----------|-------------|
| ShortcutName | Display name for shortcut (default MyApp) |
| TargetExecutable | Executable path or name (required) |
| Arguments | Optional CLI arguments |
| WorkingDirectory | Process working dir (default: target folder) |
| IconPath | .ico or exe/dll resource path |
| Force | Recreate existing shortcut |
| ClipboardListPath | File with one snippet per line |
| NoClipboardLoad | Skip clipboard seeding |
| PreserveClipboard | Restore original clipboard after seeding |
| ClipboardSetDelayMs | Delay between clipboard sets (default 120) |
| AllowElevated | Allow running while elevated (not recommended) |
| DryRun | Simulate actions only |

## Exit Codes
| Code | Meaning |
|------|---------|
| 0 | Success (already pinned or pinned) |
| 1 | Target executable not found |
| 2 | Pin verb unavailable (manual pin required) |
| 3 | Missing required parameter(s) |
| 10 | Elevated without -AllowElevated |
| 11 | Dry run completed |
| 99 | Unexpected error |

## Notes
- Pinning via automation is not officially supported; reliability depends on Windows build and policy.
- Elevated sessions usually hide the pin verb.
- Clipboard history order should reflect the sequence of Set-Clipboard calls; adjust -ClipboardSetDelayMs if entries are missing.

## License
SPDX-License-Identifier: MIT
