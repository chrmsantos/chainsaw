# =============================================================================
# CHAINSAW - Encoding Validator Helper
# =============================================================================
# Funcoes auxiliares para validacao de encoding em arquivos
# =============================================================================

<#
.SYNOPSIS
    Valida o encoding de um arquivo

.PARAMETER FilePath
    Caminho do arquivo a validar

.PARAMETER ExpectedEncoding
    Encoding esperado: UTF8, UTF8-BOM, ASCII, ou UTF16

.EXAMPLE
    Test-FileEncoding -FilePath "script.ps1" -ExpectedEncoding "UTF8"
#>
function Test-FileEncoding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $false)]
        [ValidateSet('UTF8', 'UTF8-BOM', 'ASCII', 'UTF16-LE', 'UTF16-BE')]
        [string]$ExpectedEncoding = 'UTF8'
    )

    if (-not (Test-Path $FilePath)) {
        throw "Arquivo nao encontrado: $FilePath"
    }

    $bytes = [System.IO.File]::ReadAllBytes($FilePath)

    # Detecta BOM
    $detectedEncoding = 'Unknown'

    if ($bytes.Length -ge 3) {
        # UTF-8 BOM (EF BB BF)
        if (($bytes[0] -eq 0xEF) -and ($bytes[1] -eq 0xBB) -and ($bytes[2] -eq 0xBF)) {
            $detectedEncoding = 'UTF8-BOM'
        }
    }

    if ($bytes.Length -ge 2) {
        # UTF-16 LE BOM (FF FE)
        if (($bytes[0] -eq 0xFF) -and ($bytes[1] -eq 0xFE)) {
            $detectedEncoding = 'UTF16-LE'
        }

        # UTF-16 BE BOM (FE FF)
        if (($bytes[0] -eq 0xFE) -and ($bytes[1] -eq 0xFF)) {
            $detectedEncoding = 'UTF16-BE'
        }
    }

    # Se nao tem BOM, verifica se e ASCII ou UTF-8
    if ($detectedEncoding -eq 'Unknown') {
        $isAscii = $true
        foreach ($byte in $bytes) {
            if ($byte -ge 128) {
                $isAscii = $false
                break
            }
        }

        if ($isAscii) {
            $detectedEncoding = 'ASCII'
        }
        else {
            # Assume UTF-8 sem BOM
            $detectedEncoding = 'UTF8'
        }
    }

    return [PSCustomObject]@{
        FilePath         = $FilePath
        DetectedEncoding = $detectedEncoding
        ExpectedEncoding = $ExpectedEncoding
        IsValid          = ($detectedEncoding -eq $ExpectedEncoding) -or
        (($ExpectedEncoding -eq 'UTF8') -and ($detectedEncoding -eq 'ASCII'))
    }
}

<#
.SYNOPSIS
    Verifica se um arquivo contem caracteres corrompidos

.PARAMETER FilePath
    Caminho do arquivo a verificar

.EXAMPLE
    Test-CorruptedCharacters -FilePath "arquivo.txt"
#>
function Test-CorruptedCharacters {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    if (-not (Test-Path $FilePath)) {
        throw "Arquivo nao encontrado: $FilePath"
    }

    $content = Get-Content $FilePath -Raw -Encoding UTF8

    $issues = @()

    # Verifica caractere de substituicao Unicode (U+FFFD)
    $replacementChar = [string][char]0xFFFD
    if ($content -like ("*" + $replacementChar + "*")) {
        $issues += "Contem caractere de substituicao Unicode (U+FFFD)"
    }

    # Verifica null bytes
    if ($content -match '\x00') {
        $issues += "Contem null bytes"
    }

    # Verifica caracteres de controle invalidos
    $controlCharsPattern = '[\x01-\x08\x0B\x0C\x0E-\x1F]'
    if ($content -match $controlCharsPattern) {
        $issues += "Contem caracteres de controle invalidos"
    }

    return [PSCustomObject]@{
        FilePath  = $FilePath
        HasIssues = ($issues.Count -gt 0)
        Issues    = $issues
    }
}

<#
.SYNOPSIS
    Verifica se caracteres acentuados portugueses sao lidos corretamente

.PARAMETER FilePath
    Caminho do arquivo a verificar

.EXAMPLE
    Test-PortugueseAccents -FilePath "script.ps1"
#>
function Test-PortugueseAccents {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    if (-not (Test-Path $FilePath)) {
        throw "Arquivo nao encontrado: $FilePath"
    }

    $contentUtf8 = Get-Content $FilePath -Raw -Encoding UTF8
    $contentDefault = Get-Content $FilePath -Raw

    $accentedCodePoints = @(
        0x00E3, 0x00E1, 0x00E0, 0x00E2, 0x00E9, 0x00EA, 0x00ED, 0x00F3, 0x00F4, 0x00F5, 0x00FA, 0x00E7,
        0x00C3, 0x00C1, 0x00C0, 0x00C2, 0x00C9, 0x00CA, 0x00CD, 0x00D3, 0x00D4, 0x00D5, 0x00DA, 0x00C7
    )
    $accentedChars = $accentedCodePoints | ForEach-Object { [string][char]$_ }

    $foundAccents = @()
    foreach ($char in $accentedChars) {
        if ($contentUtf8 -match [regex]::Escape($char)) {
            $foundAccents += $char
        }
    }

    return [PSCustomObject]@{
        FilePath        = $FilePath
        HasAccents      = ($foundAccents.Count -gt 0)
        FoundAccents    = $foundAccents
        Utf8Length      = $contentUtf8.Length
        DefaultLength   = $contentDefault.Length
        EncodingMatches = ($contentUtf8.Length -eq $contentDefault.Length)
    }
}

<#
.SYNOPSIS
    Valida line endings de um arquivo

.PARAMETER FilePath
    Caminho do arquivo a verificar

.PARAMETER ExpectedLineEnding
    Line ending esperado: CRLF (Windows), LF (Unix), ou CR (Mac)

.EXAMPLE
    Test-LineEndings -FilePath "script.ps1" -ExpectedLineEnding "CRLF"
#>
function Test-LineEndings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $false)]
        [ValidateSet('CRLF', 'LF', 'CR', 'Mixed')]
        [string]$ExpectedLineEnding = 'CRLF'
    )

    if (-not (Test-Path $FilePath)) {
        throw "Arquivo nao encontrado: $FilePath"
    }

    $bytes = [System.IO.File]::ReadAllBytes($FilePath)

    $hasCRLF = $false
    $hasLF = $false
    $hasCR = $false

    for ($i = 0; $i -lt ($bytes.Length - 1); $i++) {
        if (($bytes[$i] -eq 0x0D) -and ($bytes[$i + 1] -eq 0x0A)) {
            $hasCRLF = $true
        }
        elseif ($bytes[$i] -eq 0x0A) {
            $hasLF = $true
        }
        elseif ($bytes[$i] -eq 0x0D) {
            $hasCR = $true
        }
    }

    $detectedLineEnding = 'None'
    if ($hasCRLF -and -not $hasLF -and -not $hasCR) {
        $detectedLineEnding = 'CRLF'
    }
    elseif ($hasLF -and -not $hasCRLF -and -not $hasCR) {
        $detectedLineEnding = 'LF'
    }
    elseif ($hasCR -and -not $hasCRLF -and -not $hasLF) {
        $detectedLineEnding = 'CR'
    }
    elseif ($hasCRLF -or $hasLF -or $hasCR) {
        $detectedLineEnding = 'Mixed'
    }

    return [PSCustomObject]@{
        FilePath           = $FilePath
        DetectedLineEnding = $detectedLineEnding
        ExpectedLineEnding = $ExpectedLineEnding
        IsValid            = ($detectedLineEnding -eq $ExpectedLineEnding)
        HasCRLF            = $hasCRLF
        HasLF              = $hasLF
        HasCR              = $hasCR
    }
}

<#
.SYNOPSIS
    Gera relatorio de encoding para todos os arquivos em um diretorio

.PARAMETER Path
    Caminho do diretorio a analisar

.PARAMETER Filter
    Filtro de arquivos (*.ps1, *.md, etc)

.EXAMPLE
    Get-EncodingReport -Path ".\tests" -Filter "*.ps1"
#>
function Get-EncodingReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [string]$Filter = "*.*"
    )

    if (-not (Test-Path $Path)) {
        throw "Diretorio nao encontrado: $Path"
    }

    $files = Get-ChildItem -Path $Path -Filter $Filter -Recurse -File

    $report = @()

    foreach ($file in $files) {
        $encodingInfo = Test-FileEncoding -FilePath $file.FullName
        $corruptedInfo = Test-CorruptedCharacters -FilePath $file.FullName
        $lineEndingInfo = Test-LineEndings -FilePath $file.FullName
        $accentInfo = Test-PortugueseAccents -FilePath $file.FullName

        $report += [PSCustomObject]@{
            FileName         = $file.Name
            FullPath         = $file.FullName
            Encoding         = $encodingInfo.DetectedEncoding
            EncodingValid    = $encodingInfo.IsValid
            HasCorruption    = $corruptedInfo.HasIssues
            CorruptionIssues = $corruptedInfo.Issues -join '; '
            LineEnding       = $lineEndingInfo.DetectedLineEnding
            LineEndingValid  = $lineEndingInfo.IsValid
            HasAccents       = $accentInfo.HasAccents
            AccentsFound     = $accentInfo.FoundAccents -join ', '
        }
    }

    return $report
}

# Exporta funcoes
Export-ModuleMember -Function @(
    'Test-FileEncoding',
    'Test-CorruptedCharacters',
    'Test-PortugueseAccents',
    'Test-LineEndings',
    'Get-EncodingReport'
)
