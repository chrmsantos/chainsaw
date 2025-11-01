param(
    [string]$Path
)

Set-StrictMode -Version Latest

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Error "File not found: $Path"
    exit 1
}

$text = [System.IO.File]::ReadAllText($Path)
$declared = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)

# Gather declarations
$declPatterns = @(
    '(?is)\b(?:dim|private|public|static|global)\s+([^\n\'']+)',
    '(?is)\bconst\s+([A-Za-z_][A-Za-z0-9_]*)'
)
foreach ($pattern in $declPatterns) {
    foreach ($m in [regex]::Matches($text, $pattern)) {
        $declLine = $m.Groups[1].Value -replace '\r', ''
        foreach ($part in $declLine -split ',') {
            $trim = $part.Trim()
            if ($trim.Length -eq 0) { continue }
            if ($trim -match '([A-Za-z_][A-Za-z0-9_]*)') {
                [void]$declared.Add($matches[1])
            }
        }
    }
}

# Function/Sub names and params
$fnPattern = '(?is)\b(?:sub|function)\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(([^)]*)\)'
foreach ($m in [regex]::Matches($text, $fnPattern)) {
    [void]$declared.Add($m.Groups[1].Value)
    $params = $m.Groups[2].Value
    if ($params.Trim().Length -eq 0) { continue }
    foreach ($param in $params -split ',') {
        $p = $param.Trim()
        if ($p.Length -eq 0) { continue }
        $p = $p -replace '(?i)^(optional\s+)?(byref|byval)\s+', ''
        $p = $p.TrimStart('_').Trim()
        $p = ($p -replace '(?i)\bas\s+[A-Za-z0-9_.]+$', '').Trim()
        if ($p -match '([A-Za-z_][A-Za-z0-9_]*)') {
            [void]$declared.Add($matches[1])
        }
    }
}

# For loops
$forEachPattern = '(?is)\bfor\s+each\s+([A-Za-z_][A-Za-z0-9_]*)\s+in\b'
foreach ($m in [regex]::Matches($text, $forEachPattern)) {
    [void]$declared.Add($m.Groups[1].Value)
}
$forPattern = '(?is)\bfor\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\b'
foreach ($m in [regex]::Matches($text, $forPattern)) {
    [void]$declared.Add($m.Groups[1].Value)
}

$keywords = @('set','let')

$assignmentPattern = '(?im)^(\s*)(?:set\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*='
$results = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($m in [regex]::Matches($text, $assignmentPattern)) {
    $name = $m.Groups[2].Value
    if ($declared.Contains($name)) { continue }
    if ($keywords -contains $name.ToLowerInvariant()) { continue }
    [void]$results.Add($name)
}

$results | Sort-Object
