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
    $p = Join-Path $root 'installation\inst_scripts'
    return Get-ChildItem -Path $p -Filter *.ps1 -File -ErrorAction SilentlyContinue
}
function Get-VbaFiles {
    $root = Get-RepoRoot
    $p = Join-Path $root 'source\backups'
    return Get-ChildItem -Path $p -Filter *.bas -Recurse -File -ErrorAction SilentlyContinue
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
