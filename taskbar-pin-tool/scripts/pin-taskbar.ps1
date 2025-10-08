# Wrapper that invokes module function if imported standalone
$modulePath = Split-Path -Parent $PSCommandPath
$root = Split-Path -Parent $modulePath
$modFile = Join-Path $root 'TaskbarPinTool.psm1'
if(Test-Path $modFile){ Import-Module $modFile -Force -ErrorAction SilentlyContinue }

param(
    [string]$ShortcutName = 'MyApp'
  , [Parameter(Mandatory=$true)][string]$TargetExecutable
  , [string]$Arguments = ''
  , [string]$WorkingDirectory = ''
  , [string]$IconPath = ''
  , [switch]$Force
  , [string]$ClipboardListPath = ''
  , [switch]$NoClipboardLoad
  , [switch]$PreserveClipboard
  , [int]$ClipboardSetDelayMs = 120
  , [switch]$AllowElevated
  , [switch]$DryRun
)

New-ApplicationTaskbarPin @PSBoundParameters
