function Get-RepoRoot {
    $candidate = $null
    if ($PSScriptRoot) {
        $candidate = Split-Path -Parent $PSScriptRoot
    }
    if (-not $candidate -and $MyInvocation -and $MyInvocation.MyCommand.Path) {
        $candidate = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
    }
    if (-not $candidate) { $candidate = (Get-Location).ProviderPath }
    return $candidate
}
function Get-PowerShellScripts {
    $root = Get-RepoRoot
    $p = Join-Path $root 'tools\export'
    return Get-ChildItem -Path $p -Filter *.ps1 -File -ErrorAction SilentlyContinue
}
function Get-VbaFiles {
    $root = Get-RepoRoot
    $sourceMain = Join-Path $root 'source\main'
    $sourceBackups = Join-Path $root 'source\backups'

    $files = @()
    if (Test-Path $sourceMain) {
        $files += Get-ChildItem -Path $sourceMain -Filter *.bas -Recurse -File -ErrorAction SilentlyContinue
    }
    if (Test-Path $sourceBackups) {
        $files += Get-ChildItem -Path $sourceBackups -Filter *.bas -Recurse -File -ErrorAction SilentlyContinue
    }
    return $files
}
function Get-Docs {
    $root = Get-RepoRoot
    $p = Join-Path $root 'docs'
    return Get-ChildItem -Path $p -Filter *.md -File -ErrorAction SilentlyContinue
}
function Get-ProjectFile {
    param([string]$RelativePath)
    $root = Get-RepoRoot
    return Join-Path $root $RelativePath
}
