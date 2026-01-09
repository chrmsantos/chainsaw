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
    $candidates = @(
        (Join-Path $root 'tests'),
        $root
    )

    $files = @()
    foreach ($p in $candidates) {
        if (Test-Path $p) {
            $files += Get-ChildItem -Path $p -Filter *.ps1 -File -Recurse -ErrorAction SilentlyContinue
        }
    }

    return $files | Sort-Object FullName -Unique
}
function Get-VbaFiles {
    $root = Get-RepoRoot
    $sourceMain = Join-Path $root 'source\main'

    $files = @()
    if (Test-Path $sourceMain) {
        $files += Get-ChildItem -Path $sourceMain -Filter *.bas -Recurse -File -ErrorAction SilentlyContinue
    }
    return $files
}
function Get-Docs {
    $root = Get-RepoRoot
    return Get-ChildItem -Path $root -Filter *.md -File -ErrorAction SilentlyContinue
}
function Get-ProjectFile {
    param([string]$RelativePath)
    $root = Get-RepoRoot
    return Join-Path $root $RelativePath
}
