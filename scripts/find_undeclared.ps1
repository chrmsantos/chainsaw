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
$keywords = @(
    'and','as','byref','byval','call','case','const','debug','dim','do','each','else','elseif','end','enum','error','exit','false','for','function','goto','if','in','is','loop','mod','next','not','nothing','on','optional','or','private','public','resume','select','set','step','sub','then','true','to','until','wend','while','with','let','type','line','lock','unlock','lset','rset'
)

$typePattern = '(?is)\b(?:dim|private|public|static|global)\s+([^\n\'']+)'
foreach ($m in [regex]::Matches($text, $typePattern)) {
    $declLine = $m.Groups[1].Value -replace '\r', ''
    foreach ($part in $declLine -split ',') {
        $trimmed = $part.Trim()
        if ($trimmed.Length -eq 0) { continue }
        if ($trimmed -match '([A-Za-z_][A-Za-z0-9_]*)') {
            [void]$declared.Add($matches[1])
        }
    }
}

$constPattern = '(?is)\bconst\s+([A-Za-z_][A-Za-z0-9_]*)'
foreach ($m in [regex]::Matches($text, $constPattern)) {
    [void]$declared.Add($m.Groups[1].Value)
}

$fnPattern = '(?is)\b(?:sub|function)\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(([^)]*)\)'
foreach ($m in [regex]::Matches($text, $fnPattern)) {
    [void]$declared.Add($m.Groups[1].Value)
    $params = $m.Groups[2].Value
    if ($params.Trim().Length -eq 0) { continue }
    foreach ($param in $params -split ',') {
        $p = $param.Trim()
        if ($p.Length -eq 0) { continue }
        $p = $p -replace '(?i)^(optional\s+)?(byref|byval)\s+', ''
        if ($p -match '([A-Za-z_][A-Za-z0-9_]*)') {
            [void]$declared.Add($matches[1])
        }
    }
}

$forEachPattern = '(?is)\bfor\s+each\s+([A-Za-z_][A-Za-z0-9_]*)\s+in\b'
foreach ($m in [regex]::Matches($text, $forEachPattern)) {
    [void]$declared.Add($m.Groups[1].Value)
}

$forPattern = '(?is)\bfor\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\b'
foreach ($m in [regex]::Matches($text, $forPattern)) {
    [void]$declared.Add($m.Groups[1].Value)
}

$tokens = [regex]::Matches($text, '[A-Za-z_][A-Za-z0-9_]*')
$suspects = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($tok in $tokens) {
    $word = $tok.Value
    if ($declared.Contains($word)) { continue }
    $lower = $word.ToLowerInvariant()
    if ($keywords -contains $lower) { continue }
    if ($word.StartsWith('wd') -or $word.StartsWith('vb') -or $word.StartsWith('xl') -or $word.StartsWith('mso') -or $word.StartsWith('pp') -or $word.StartsWith('ad')) { continue }
    if ($word.StartsWith('_')) { continue }
    if ($word -match '^[A-Z]') { continue }
    [void]$suspects.Add($word)
}

$suspects | Sort-Object
