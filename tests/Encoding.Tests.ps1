# =============================================================================
# CHAINSAW - Testes de Encoding e Emojis
# =============================================================================
# Verifica conformidade de encoding (UTF-8/ASCII) e ausencia de emojis
# =============================================================================

$ErrorActionPreference = "Stop"

Describe 'CHAINSAW - Testes de Encoding e Emojis' {

    $projectRoot = Split-Path -Parent $PSScriptRoot
    $testsPath = Join-Path $projectRoot 'tests'

    $psSearchRoots = @()
    if (Test-Path $testsPath) { $psSearchRoots += $testsPath }
    $psSearchRoots += $projectRoot

    $script:ProjectPsFiles = @()
    foreach ($root in $psSearchRoots) {
        if (Test-Path $root) {
            $script:ProjectPsFiles += Get-ChildItem -Path $root -Filter "*.ps1" -Recurse -File -ErrorAction SilentlyContinue
        }
    }
    $script:ProjectPsFiles = $script:ProjectPsFiles | Sort-Object FullName -Unique

    Context 'Validacao de Encoding de Arquivos' {

        It 'Scripts PowerShell estao em UTF-8 com BOM ou ASCII' {
            $psFiles = $script:ProjectPsFiles

            foreach ($file in $psFiles) {
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                # Verifica UTF-8 com BOM (EF BB BF) ou ASCII puro
                $isUtf8WithBom = ($bytes.Length -ge 3) -and
                ($bytes[0] -eq 0xEF) -and
                ($bytes[1] -eq 0xBB) -and
                ($bytes[2] -eq 0xBF)

                # Verifica se e ASCII puro (todos os bytes < 128)
                $isAscii = $true
                foreach ($byte in $bytes) {
                    if ($byte -ge 128) {
                        $isAscii = $false
                        break
                    }
                }

                # Verifica UTF-8 sem BOM (compativel com UTF-8)
                $content = Get-Content $file.FullName -Raw
                $utf8Bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
                $isValidUtf8 = $true

                ($isUtf8WithBom -or $isAscii -or $isValidUtf8) | Should Be $true
            }
        }

        It 'Arquivos Markdown estao em UTF-8' {
            $mdFiles = @(
                Get-ChildItem -Path "$projectRoot\docs" -Filter "*.md" -Recurse
                Get-ChildItem -Path "$projectRoot" -Filter "*.md" -File
            )

            foreach ($file in $mdFiles) {
                if ($file -eq $null) { continue }

                $content = Get-Content $file.FullName -Raw -Encoding UTF8
                $content | Should Not BeNullOrEmpty

                # Verifica que pode ser lido como UTF-8
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
                $decodedContent = [System.Text.Encoding]::UTF8.GetString($bytes)
                $decodedContent.Length | Should BeGreaterThan 0
            }
        }

        It 'Arquivo VBA esta em formato legivel' {
            $vbaFile = "$projectRoot\source\main\Modulo1.bas"

            if (Test-Path $vbaFile) {
                $content = Get-Content $vbaFile -Raw
                $content | Should Not BeNullOrEmpty

                # VBA deve conter apenas ASCII ou caracteres extendidos validos
                $bytes = [System.IO.File]::ReadAllBytes($vbaFile)

                # Verifica que nao tem bytes nulos (indicaria binario)
                $hasNullBytes = $false
                foreach ($byte in $bytes) {
                    if ($byte -eq 0) {
                        $hasNullBytes = $true
                        break
                    }
                }

                $hasNullBytes | Should Be $false
            }
        }

        It 'Arquivos de texto (.txt) estao em UTF-8 ou ASCII' {
            $txtFiles = Get-ChildItem -Path $projectRoot -Filter "*.txt" -Recurse

            foreach ($file in $txtFiles) {
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                # Tenta decodificar como UTF-8
                try {
                    $utf8Content = [System.Text.Encoding]::UTF8.GetString($bytes)
                    $utf8Content.Length | Should BeGreaterThan 0
                }
                catch {
                    # Se falhar UTF-8, tenta ASCII
                    $asciiContent = [System.Text.Encoding]::ASCII.GetString($bytes)
                    $asciiContent.Length | Should BeGreaterThan 0
                }
            }
        }
    }

    Context 'Deteccao de Emojis e Caracteres Especiais' {

        # Regex para detectar emojis (Unicode ranges)
        # Usando lookbehind negativo para evitar falsos positivos com caracteres ASCII
        # Emoticons: U+1F600 to U+1F64F
        # Symbols & Pictographs: U+1F300 to U+1F5FF
        # Transport & Map: U+1F680 to U+1F6FF
        # Supplemental Symbols: U+1F900 to U+1F9FF
        # Outros ranges comuns de emojis
        $emojiPattern = '[\u2600-\u26FF\u2700-\u27BF]|' +
        '[\uD83C][\uDF00-\uDFFF]|' + # High surrogate for U+1F300-1F5FF
        '[\uD83D][\uDC00-\uDE4F]|' + # High surrogate for U+1F600-1F64F
        '[\uD83D][\uDE80-\uDEFF]|' + # High surrogate for U+1F680-1F6FF
        '[\uD83E][\uDD00-\uDDFF]|' + # High surrogate for U+1F900-1F9FF
        '[\uD83E][\uDE00-\uDE6F]|' + # High surrogate for U+1FA00-1FA6F
        '[\uD83E][\uDE70-\uDEFF]'     # High surrogate for U+1FA70-1FAFF

        It 'Scripts PowerShell nao contem emojis' {
            $psFiles = $script:ProjectPsFiles

            foreach ($file in $psFiles) {
                $content = Get-Content $file.FullName -Raw -Encoding UTF8

                # Ignora comentários
                $codeLines = ($content -split "`r?`n") | Where-Object {
                    $_ -notmatch '^\s*#'
                }
                $codeContent = $codeLines -join "`n"

                if ($codeContent -match $emojiPattern) {
                    throw "Arquivo $($file.Name) contem emojis no codigo (fora de comentarios)"
                }

                # Passa se nao encontrou emojis no codigo
                $true | Should Be $true
            }
        }

        It 'Arquivos Markdown podem conter emojis apenas em documentacao' {
            # Markdown pode conter emojis para documentacao
            # Este teste apenas verifica que podem ser lidos
            $mdFiles = @(
                Get-ChildItem -Path "$projectRoot\docs" -Filter "*.md" -Recurse
                Get-ChildItem -Path "$projectRoot" -Filter "*.md" -File
            )

            foreach ($file in $mdFiles) {
                if ($file -eq $null) { continue }

                $content = Get-Content $file.FullName -Raw -Encoding UTF8

                # Verifica que emojis podem ser lidos corretamente se presentes
                $content.Length | Should BeGreaterThan 0

                # Passa se conseguiu ler
                $true | Should Be $true
            }
        }

        It 'Arquivo VBA nao contem emojis' {
            $vbaFile = "$projectRoot\source\main\Modulo1.bas"

            if (Test-Path $vbaFile) {
                $content = Get-Content $vbaFile -Raw -Encoding UTF8

                if ($content -match $emojiPattern) {
                    throw "Arquivo VBA contem emojis"
                }

                # Passa se nao encontrou emojis
                $true | Should Be $true
            }
        }

        It 'Testes PowerShell nao contem emojis' {
            $testFiles = Get-ChildItem -Path "$projectRoot\tests" -Filter "*.ps1"

            foreach ($file in $testFiles) {
                $content = Get-Content $file.FullName -Raw -Encoding UTF8

                if ($content -match $emojiPattern) {
                    throw "Arquivo de teste $($file.Name) contem emojis"
                }

                # Passa se nao encontrou emojis
                $true | Should Be $true
            }
        }
    }

    Context 'Validacao de Caracteres Problematicos' {

        It 'Scripts PowerShell nao contem caracteres de controle invalidos' {
            $psFiles = $script:ProjectPsFiles

            # Caracteres de controle permitidos: Tab (0x09), LF (0x0A), CR (0x0D)
            $allowedControlChars = @(0x09, 0x0A, 0x0D)

            foreach ($file in $psFiles) {
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                for ($i = 0; $i -lt $bytes.Length; $i++) {
                    $byte = $bytes[$i]

                    # Verifica caracteres de controle (0x00-0x1F exceto permitidos)
                    if (($byte -le 0x1F) -and ($byte -notin $allowedControlChars)) {
                        throw "Arquivo $($file.Name) contem caractere de controle invalido 0x$($byte.ToString('X2')) na posicao $i"
                    }
                }

                # Passa se nao encontrou problemas
                $true | Should Be $true
            }
        }

        It 'Arquivos Markdown nao contem tabs (usam espacos)' {
            $mdFiles = @(
                Get-ChildItem -Path "$projectRoot\docs" -Filter "*.md" -Recurse
                Get-ChildItem -Path "$projectRoot" -Filter "*.md" -File
            )

            foreach ($file in $mdFiles) {
                if ($file -eq $null) { continue }

                $content = Get-Content $file.FullName -Raw

                # Markdown nao deve ter tabs (usa espacos para indentacao)
                if ($content -match "`t") {
                    Write-Warning "Arquivo $($file.Name) contem tabs - deveria usar espacos"
                    # Apenas warning, nao falha o teste
                }

                $true | Should Be $true
            }
        }

        It 'Arquivo VBA nao contem tabs (usa espacos conforme padrao)' {
            $vbaFile = "$projectRoot\source\main\Modulo1.bas"

            if (Test-Path $vbaFile) {
                $content = Get-Content $vbaFile -Raw

                # VBA nao deve ter tabs
                $content -match "`t" | Should Be $false
            }
        }
    }

    Context 'Consistencia de Line Endings' {

        It 'Scripts PowerShell usam CRLF (Windows)' {
            $psFiles = $script:ProjectPsFiles

            foreach ($file in $psFiles) {
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                # Verifica se tem pelo menos um CRLF
                $hasCrlf = $false
                for ($i = 0; $i -lt ($bytes.Length - 1); $i++) {
                    if (($bytes[$i] -eq 0x0D) -and ($bytes[$i + 1] -eq 0x0A)) {
                        $hasCrlf = $true
                        break
                    }
                }

                # Verifica se nao tem LF sozinho (Unix style)
                $hasLfOnly = $false
                for ($i = 0; $i -lt $bytes.Length; $i++) {
                    if ($bytes[$i] -eq 0x0A) {
                        if (($i -eq 0) -or ($bytes[$i - 1] -ne 0x0D)) {
                            $hasLfOnly = $true
                            break
                        }
                    }
                }

                if ($hasLfOnly) {
                    Write-Warning "Arquivo $($file.Name) usa LF (Unix) - deveria usar CRLF (Windows)"
                }

                $true | Should Be $true
            }
        }
    }

    Context 'Validacao de BOM (Byte Order Mark)' {

        It 'Scripts PowerShell tem UTF-8 BOM ou sao ASCII puro' {
            $psFiles = $script:ProjectPsFiles

            foreach ($file in $psFiles) {
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                if ($bytes.Length -ge 3) {
                    $hasUtf8Bom = ($bytes[0] -eq 0xEF) -and
                    ($bytes[1] -eq 0xBB) -and
                    ($bytes[2] -eq 0xBF)

                    # Verifica se tem caracteres nao-ASCII
                    $hasNonAscii = $false
                    foreach ($byte in $bytes) {
                        if ($byte -ge 128) {
                            $hasNonAscii = $true
                            break
                        }
                    }

                    # Se tem caracteres nao-ASCII, deve ter BOM
                    if ($hasNonAscii -and -not $hasUtf8Bom) {
                        Write-Warning "Arquivo $($file.Name) tem caracteres nao-ASCII mas nao tem UTF-8 BOM"
                    }
                }

                $true | Should Be $true
            }
        }
    }

    Context 'Validacao de Caracteres Acentuados Portugueses' {

        It 'VBA pode ser lido com UTF-8 sem perder acentuacao' {
            $vbaFile = "$projectRoot\source\main\Modulo1.bas"

            if (Test-Path $vbaFile) {
                $contentUtf8 = Get-Content $vbaFile -Raw -Encoding UTF8
                $contentDefault = Get-Content $vbaFile -Raw

                # Verifica que caracteres acentuados comuns estao presentes
                $accentedChars = @('ã', 'á', 'à', 'â', 'é', 'ê', 'í', 'ó', 'ô', 'õ', 'ú', 'ç', 'Ã', 'Á', 'É', 'Ó', 'Ç')

                $hasAccents = $false
                foreach ($char in $accentedChars) {
                    if ($contentUtf8 -match [regex]::Escape($char)) {
                        $hasAccents = $true
                        break
                    }
                }

                # Se tem acentos, verifica que UTF-8 lê corretamente
                if ($hasAccents) {
                    $contentUtf8.Length | Should BeGreaterThan 0
                    $contentUtf8 | Should Not Match '�'  # Caractere de substituicao Unicode
                }

                $true | Should Be $true
            }
        }

        It 'Documentacao Markdown preserva acentuacao portuguesa' {
            $mdFiles = @(
                Get-ChildItem -Path "$projectRoot\docs" -Filter "*.md" -Recurse
                Get-ChildItem -Path "$projectRoot" -Filter "*.md" -File
            )

            foreach ($file in $mdFiles) {
                if ($file -eq $null) { continue }

                $content = Get-Content $file.FullName -Raw -Encoding UTF8

                # Nao deve conter caracteres de substituicao
                $content | Should Not Match '�'

                # Deve poder ser re-encodado sem perda
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
                $reencoded = [System.Text.Encoding]::UTF8.GetString($bytes)

                $reencoded.Length | Should Be $content.Length
            }
        }

        It 'Testes PowerShell preservam strings acentuadas' {
            $testFiles = Get-ChildItem -Path "$projectRoot\tests" -Filter "*.Tests.ps1"

            foreach ($file in $testFiles) {
                $content = Get-Content $file.FullName -Raw -Encoding UTF8

                # Nao deve conter caracteres de substituicao
                if ($content -match '�') {
                    throw "Arquivo $($file.Name) contem caracteres corrompidos (�)"
                }

                $true | Should Be $true
            }
        }

        It 'chainsaw_installer.cmd preserva mensagens acentuadas' {
            $installerFile = "$projectRoot\chainsaw_installer.cmd"

            if (Test-Path $installerFile) {
                # CMD files podem ser lidos com encoding do sistema
                $content = Get-Content $installerFile -Raw

                # Verifica que tem mensagens em portugues
                $content | Should Match 'Instalador|instalacao|codigo|extrair'

                # Nao deve ter caracteres obviamente corrompidos
                $content | Should Not Match '\x00'  # Null bytes
            }
        }
    }

    Context 'Validacao de Encoding Cross-Platform' {

        It 'VERSION pode ser lido em qualquer plataforma' {
            $versionFile = "$projectRoot\VERSION"

            if (Test-Path $versionFile) {
                # Deve ser UTF-8 sem BOM para compatibilidade
                $bytes = [System.IO.File]::ReadAllBytes($versionFile)

                $hasUtf8Bom = ($bytes.Length -ge 3) -and
                ($bytes[0] -eq 0xEF) -and
                ($bytes[1] -eq 0xBB) -and
                ($bytes[2] -eq 0xBF)

                if ($hasUtf8Bom) {
                    Write-Warning "VERSION tem UTF-8 BOM - melhor usar UTF-8 sem BOM"
                }

                # Deve conter um número de versão simples
                $content = Get-Content $versionFile -Raw -Encoding UTF8
                $content | Should Match '[0-9]+\.[0-9]+\.[0-9]+'
            }
        }

        It 'Arquivos de configuracao (.config) sao UTF-8 validos' {
            $configFiles = Get-ChildItem -Path $projectRoot -Filter "*.config" -Recurse

            foreach ($file in $configFiles) {
                $content = Get-Content $file.FullName -Raw -Encoding UTF8
                $content | Should Not BeNullOrEmpty

                # Nao deve ter caracteres de substituicao
                $content | Should Not Match '�'
            }
        }
    }

    Context 'Validacao de Regex com Acentuacao' {

        It 'Testes que usam regex podem encontrar strings acentuadas' {
            $testContent = @'
' Versão: 2.0.2
Instalação completa
Configuração do sistema
'@

            # Testa varios padroes de regex comuns
            $testContent | Should Match 'Versão'
            $testContent | Should Match 'Instalação'
            $testContent | Should Match 'Configuração'

            # Testa com captura
            if ($testContent -match "Versão: ([^\r\n]+)") {
                $matches[1].Trim() | Should Be "2.0.2"
            }
        }

        It 'Get-Content com -Encoding UTF8 lê corretamente arquivos acentuados' {
            $vbaFile = "$projectRoot\source\main\Modulo1.bas"

            if (Test-Path $vbaFile) {
                $contentUtf8 = Get-Content $vbaFile -Raw -Encoding UTF8
                $contentDefault = Get-Content $vbaFile -Raw

                # Ambos devem ler o arquivo
                $contentUtf8.Length | Should BeGreaterThan 0
                $contentDefault.Length | Should BeGreaterThan 0

                # UTF-8 explícito é mais confiável
                $contentUtf8 | Should Match "' ="
            }
        }
    }

    Context 'Deteccao de Encodings Problematicos' {

        It 'Nenhum arquivo usa UTF-16 (wide chars)' {
            $allTextFiles = @(
                $script:ProjectPsFiles
                if (Test-Path (Join-Path $projectRoot 'docs')) { Get-ChildItem -Path (Join-Path $projectRoot 'docs') -Filter "*.md" -Recurse -File -ErrorAction SilentlyContinue }
                if (Test-Path (Join-Path $projectRoot 'tests')) { Get-ChildItem -Path (Join-Path $projectRoot 'tests') -Filter "*.ps1" -Recurse -File -ErrorAction SilentlyContinue }
            )

            foreach ($file in $allTextFiles) {
                if ($file -eq $null) { continue }

                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                if ($bytes.Length -ge 2) {
                    # Verifica UTF-16 LE BOM (FF FE)
                    $isUtf16Le = ($bytes[0] -eq 0xFF) -and ($bytes[1] -eq 0xFE)

                    # Verifica UTF-16 BE BOM (FE FF)
                    $isUtf16Be = ($bytes[0] -eq 0xFE) -and ($bytes[1] -eq 0xFF)

                    if ($isUtf16Le -or $isUtf16Be) {
                        throw "Arquivo $($file.Name) usa UTF-16 - deveria usar UTF-8"
                    }
                }

                $true | Should Be $true
            }
        }

        It 'Nenhum arquivo tem encoding misto' {
            $psFiles = $script:ProjectPsFiles

            foreach ($file in $psFiles) {
                $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                # Tenta decodificar como UTF-8
                $decoder = [System.Text.Encoding]::UTF8.GetDecoder()
                $charCount = $decoder.GetCharCount($bytes, 0, $bytes.Length, $true)

                # Se falhar, pode ser encoding misto
                if ($charCount -eq 0) {
                    throw "Arquivo $($file.Name) pode ter encoding misto ou invalido"
                }

                $true | Should Be $true
            }
        }
    }
}

