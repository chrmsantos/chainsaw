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

$fnPattern = '(?is)\b(?:sub|function)\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(([^)]*)\)'
foreach ($m in [regex]::Matches($text, $fnPattern)) {
    [void]$declared.Add($m.Groups[1].Value)
    $params = $m.Groups[2].Value
    if ($params.Trim().Length -eq 0) { continue }
    foreach ($param in $params -split ',') {
        $p = $param.Trim()
        if ($p.Length -eq 0) { continue }
        $p = $p -replace '(?i)^(optional\s+)?(byref|byval)\s+', ''
        $p = $p -replace '(?i)\bas\s+[A-Za-z0-9_.]+$', ''
        $p = $p.Trim()
        if ($p.Length -eq 0) { continue }
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
$ignorePrefixes = @('wd','vb','xl','mso','pp','ad','wdmso','cmsbo')
$ignoreWords = @('Application','Document','Documents','Range','Paragraph','Paragraphs','Selection','Err','Now','MsgBox','Debug','Chr','ChrW','Len','Mid','Left','Right','Trim','UCase','LCase','Replace','Format','Array','Timer','Environ','Dir','DoEvents','Shell','String','IIf','Split','Join','UBound','LBound','MsgBox','CreateObject','FreeFile','EOF','Close','Open','Print','Input','Output','Append','Random','Lock','Unlock','Log','Int','Fix','Sgn','Abs','Environ$','Hex','Val','Asc','Str','CStr','CLng','CInt','CDbl','CSng','Date','Time','DateAdd','DateDiff','DatePart','DateSerial','DateValue','TimeSerial','TimeValue','Weekday','Year','Month','Day','Hour','Minute','Second','ActiveDocument','ActiveWindow','Documents','Selection','Sections','Headers','Footers','Shapes','InlineShapes','ShapeRange','Font','Format','Type')

$undefined = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($tok in $tokens) {
    $word = $tok.Value
    if ($declared.Contains($word)) { continue }
    foreach ($prefix in $ignorePrefixes) {
        if ($word.StartsWith($prefix)) { continue 2 }
    }
    if ($ignoreWords -contains $word) { continue }
    [void]$undefined.Add($word)
}

$undefined | Sort-Object
