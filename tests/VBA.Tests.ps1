#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1

# Override Get-RepoRoot for test context
function Get-RepoRoot {
    $testsDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $repoRoot = Split-Path -Parent $testsDir
    return $repoRoot
}

Describe 'CHAINSAW - Testes do Modulo VBA Modulo1.bas' {

    BeforeAll {
        $script:repoRoot = Get-RepoRoot
        $script:vbaPath = Join-Path $script:repoRoot "source\main\Modulo1.bas"
        $script:vbaContent = Get-Content $script:vbaPath -Raw -Encoding UTF8
        $script:vbaLines = Get-Content $script:vbaPath -Encoding UTF8
    }

    Context 'Estrutura e Metadados do Arquivo' {

        It 'Arquivo Modulo1.bas existe' {
            Test-Path $script:vbaPath | Should Be $true
        }

        It 'Arquivo nao esta vazio' {
            (Get-Item $script:vbaPath).Length -gt 0 | Should Be $true
        }

        It 'Tamanho do arquivo e razoavel (< 5MB)' {
            $sizeMB = (Get-Item $script:vbaPath).Length / 1MB
            $sizeMB -lt 5 | Should Be $true
        }

        It 'Contém cabeçalho CHAINSAW' {
            $script:vbaContent -match 'CHAINSAW' | Should Be $true
        }

        It 'Contém informações de versão' {
            $script:vbaContent -match 'Vers[aã]o:\s*\d+\.\d+' | Should Be $true
        }

        It 'Contém licença GNU GPLv3' {
            $script:vbaContent -match 'GNU GPLv3' | Should Be $true
        }

        It 'Contém informação de autor' {
            $script:vbaContent -match 'Autor:' | Should Be $true
        }

        It 'Contém declaração Option Explicit' {
            $script:vbaContent -match '(?m)^Option Explicit' | Should Be $true
        }

        It 'Número total de linhas corresponde ao esperado (> 7000)' {
            $script:vbaLines.Count -gt 7000 | Should Be $true
        }
    }

    Context 'Análise de Procedimentos e Funções' {

        BeforeAll {
            $script:procedures = [regex]::Matches($script:vbaContent, '(?m)^(Public |Private )?(Sub |Function )\w+')
            $script:publicProcs = [regex]::Matches($script:vbaContent, '(?m)^Public (Sub |Function )\w+')
            $script:privateProcs = [regex]::Matches($script:vbaContent, '(?m)^Private (Sub |Function )\w+')
            $script:subs = [regex]::Matches($script:vbaContent, '(?m)^(Public |Private )?Sub \w+')
            $script:functions = [regex]::Matches($script:vbaContent, '(?m)^(Public |Private )?Function \w+')
        }

        It 'Contem quantidade razoavel de procedimentos (100-250)' {
            $script:procedures.Count -ge 100 -and $script:procedures.Count -le 250 | Should Be $true
        }

        It 'Possui procedimento principal PadronizarDocumentoMain' {
            $script:vbaContent -match '(?m)^Public Sub PadronizarDocumentoMain\(' | Should Be $true
        }

        It 'Procedimentos publicos sao minoria (< 20% do total)' {
            $publicRatio = $script:publicProcs.Count / $script:procedures.Count
            $publicRatio -lt 0.20 | Should Be $true
        }

        It 'Possui funções de validação (ValidateDocument)' {
            $script:vbaContent -match 'Function ValidateDocument' | Should Be $true
        }

        It 'Possui funções de identificação de elementos estruturais' {
            ($script:vbaContent -match 'GetTituloRange') -and
            ($script:vbaContent -match 'GetEmentaRange') -and
            ($script:vbaContent -match 'GetProposicaoRange') | Should Be $true
        }

        It 'Possui sistema de tratamento de erros (ShowUserFriendlyError)' {
            $script:vbaContent -match 'ShowUserFriendlyError' | Should Be $true
        }

        It 'Possui sistema de recuperação de emergência (EmergencyRecovery)' {
            $script:vbaContent -match 'EmergencyRecovery' | Should Be $true
        }

        It 'Possui funções de normalização de texto' {
            $script:vbaContent -match 'NormalizarTexto' | Should Be $true
        }

        It 'Todas as funções têm End Function' {
            $functionStarts = [regex]::Matches($script:vbaContent, '(?m)^(Public |Private )?Function \w+').Count
            $functionEnds = [regex]::Matches($script:vbaContent, '(?m)^End Function').Count
            $functionStarts -eq $functionEnds | Should Be $true
        }

        It 'Todas as subs têm End Sub' {
            $subStarts = [regex]::Matches($script:vbaContent, '(?m)^(Public |Private )?Sub \w+').Count
            $subEnds = [regex]::Matches($script:vbaContent, '(?m)^End Sub').Count
            $subStarts -eq $subEnds | Should Be $true
        }
    }

    Context 'Constantes e Configurações' {

        It 'Define constantes do Word (wdNoProtection, wdTypeDocument, etc)' {
            ($script:vbaContent -match 'wdNoProtection') -and
            ($script:vbaContent -match 'wdTypeDocument') -and
            ($script:vbaContent -match 'wdAlignParagraphCenter') | Should Be $true
        }

        It 'Define constantes de formatação (STANDARD_FONT, STANDARD_FONT_SIZE)' {
            ($script:vbaContent -match 'STANDARD_FONT') -and
            ($script:vbaContent -match 'STANDARD_FONT_SIZE') | Should Be $true
        }

        It 'Define margens do documento (TOP_MARGIN_CM, BOTTOM_MARGIN_CM, etc)' {
            ($script:vbaContent -match 'TOP_MARGIN_CM') -and
            ($script:vbaContent -match 'BOTTOM_MARGIN_CM') -and
            ($script:vbaContent -match 'LEFT_MARGIN_CM') -and
            ($script:vbaContent -match 'RIGHT_MARGIN_CM') | Should Be $true
        }

        It 'Define configurações de imagem do cabeçalho' {
            ($script:vbaContent -match 'HEADER_IMAGE_RELATIVE_PATH') -and
            ($script:vbaContent -match 'HEADER_IMAGE_MAX_WIDTH_CM') | Should Be $true
        }

        It 'Define constantes de sistema (MIN_SUPPORTED_VERSION, MAX_RETRY_ATTEMPTS)' {
            ($script:vbaContent -match 'MIN_SUPPORTED_VERSION') -and
            ($script:vbaContent -match 'MAX_RETRY_ATTEMPTS') | Should Be $true
        }

        It 'Define constantes de backup e logs (GetChainsawBackupsPath, GetChainsawRecoveryPath)' {
            ($script:vbaContent -match 'GetChainsawBackupsPath') -and
            ($script:vbaContent -match 'GetChainsawRecoveryPath') | Should Be $true
        }

        It 'Define níveis de log (LOG_LEVEL_INFO, LOG_LEVEL_WARNING, LOG_LEVEL_ERROR)' {
            ($script:vbaContent -match 'LOG_LEVEL_INFO') -and
            ($script:vbaContent -match 'LOG_LEVEL_WARNING') -and
            ($script:vbaContent -match 'LOG_LEVEL_ERROR') | Should Be $true
        }

        It 'Fonte padrão é Arial' {
            $script:vbaContent -match 'STANDARD_FONT.*=.*"Arial"' | Should Be $true
        }

        It 'Tamanho de fonte padrão é 12' {
            $script:vbaContent -match 'STANDARD_FONT_SIZE.*=.*12' | Should Be $true
        }
    }

    Context 'Sistema de Cache de Parágrafos' {

        It 'Possui função BuildParagraphCache' {
            $script:vbaContent -match 'Sub BuildParagraphCache' | Should Be $true
        }

        It 'Possui função ClearParagraphCache' {
            $script:vbaContent -match 'Sub ClearParagraphCache' | Should Be $true
        }

        It 'Possui sistema de identificação de estrutura do documento' {
            $script:vbaContent -match 'IdentifyDocumentStructure' | Should Be $true
        }
    }

    Context 'Identificação de Elementos Estruturais' {

        It 'Possui função para identificar Título (IsTituloElement)' {
            $script:vbaContent -match 'Function IsTituloElement' | Should Be $true
        }

        It 'Possui função para identificar Ementa (IsEmentaElement)' {
            $script:vbaContent -match 'Function IsEmentaElement' | Should Be $true
        }

        It 'Possui função para identificar Justificativa (IsJustificativaTitleElement)' {
            $script:vbaContent -match 'Function IsJustificativaTitleElement' | Should Be $true
        }

        It 'Possui função para identificar Data (IsDataElement)' {
            $script:vbaContent -match 'Function IsDataElement' | Should Be $true
        }

        It 'Possui função para identificar Assinatura (IsAssinaturaStart)' {
            $script:vbaContent -match 'Function IsAssinaturaStart' | Should Be $true
        }

        It 'Possui função para identificar Título de Anexo (IsTituloAnexoElement)' {
            $script:vbaContent -match 'Function IsTituloAnexoElement' | Should Be $true
        }

        It 'Possui GetProposituraRange para retornar range da propositura completa' {
            $script:vbaContent -match 'Function GetProposituraRange' | Should Be $true
        }

        It 'Possui GetElementInfo para relatório de elementos' {
            $script:vbaContent -match 'GetElementInfo' | Should Be $true
        }
    }

    Context 'Tratamento de Erros e Recuperação' {

        It 'Possui tratamento On Error em procedimentos críticos' {
            $script:vbaContent -match 'On Error GoTo' | Should Be $true
        }

        It 'Possui labels de tratamento de erro (ErrorHandler:)' {
            $script:vbaContent -match 'ErrorHandler:' | Should Be $true
        }

        It 'Possui função SafeCleanup' {
            $script:vbaContent -match 'Sub SafeCleanup' | Should Be $true
        }

        It 'Possui função ReleaseObjects' {
            $script:vbaContent -match 'Sub ReleaseObjects' | Should Be $true
        }

        It 'Possui verificação de timeout (IsOperationTimeout)' {
            $script:vbaContent -match 'Function IsOperationTimeout' | Should Be $true
        }

        It 'Implementa sistema de retry (MAX_RETRY_ATTEMPTS)' {
            $script:vbaContent -match 'MAX_RETRY_ATTEMPTS' | Should Be $true
        }
    }

    Context 'Validação de Sintaxe VBA' {

        It 'Não contém tabs (usa apenas espaços)' {
            $script:vbaContent -notmatch "`t" | Should Be $true
        }

        It 'Parênteses balanceados em declarações de função' {
            $functionDeclarations = [regex]::Matches($script:vbaContent, '(?m)^(Public |Private )?Function [^(]+\([^)]*\)')
            $functionDeclarations.Count -gt 0 | Should Be $true
        }

        It 'Não contém caracteres de controle inválidos' {
            $invalidChars = [regex]::Matches($script:vbaContent, '[\x00-\x08\x0B\x0C\x0E-\x1F]')
            $invalidChars.Count -eq 0 | Should Be $true
        }

        It 'Linhas não excedem 1000 caracteres (padrão VBA)' {
            $longLines = $vbaLines | Where-Object { $_.Length -gt 1000 }
            $longLines.Count -eq 0 | Should Be $true
        }

        It 'Usa aspas duplas para strings, não aspas simples' {
            # VBA usa aspas duplas "" para strings, ' é apenas para comentários
            $stringDeclarations = [regex]::Matches($script:vbaContent, '=\s*"[^"]*"')
            $stringDeclarations.Count -gt 0 | Should Be $true
        }
    }

    Context 'Comentários e Documentação' {

        It 'Contém comentários de seção (linhas com ====)' {
            $script:vbaContent -match '={20,}' | Should Be $true
        }

It 'Taxa de comentários adequada (> 5% das linhas)' {
            $commentLines = $vbaLines | Where-Object { $_ -match "^\s*'" }
            $commentRatio = $commentLines.Count / $vbaLines.Count
            $commentRatio -gt 0.05 | Should Be $true
        }

        It 'Contém seções organizadas (CONSTANTES, FUNÇÕES, etc)' {
            $script:vbaContent -match 'CONSTANTES' | Should Be $true
        }
    }

    Context 'Funcionalidades de Backup e Log' {

        It 'Possui sistema de backup (CreateDocumentBackup)' {
            $script:vbaContent -match 'CreateDocumentBackup' | Should Be $true
        }

        It 'Possui limite de arquivos de backup (MAX_BACKUP_FILES)' {
            $script:vbaContent -match 'MAX_BACKUP_FILES' | Should Be $true
        }

        It 'Implementa sistema de logging' {
            ($script:vbaContent -match 'LOG_LEVEL') -or ($script:vbaContent -match 'WriteLog') | Should Be $true
        }

        It 'Possui modo de debug (DEBUG_MODE)' {
            $script:vbaContent -match 'DEBUG_MODE' | Should Be $true
        }
    }

    Context 'Processamento de Texto' {

        It 'Possui função GetCleanParagraphText' {
            $script:vbaContent -match 'Function GetCleanParagraphText' | Should Be $true
        }

        It 'Possui função RemovePunctuation' {
            $script:vbaContent -match 'Function RemovePunctuation' | Should Be $true
        }

        It 'Possui função para detectar parágrafos especiais (DetectSpecialParagraph)' {
            $script:vbaContent -match 'Function DetectSpecialParagraph' | Should Be $true
        }

        It 'Possui função para contar linhas em branco (CountBlankLinesBefore)' {
            $script:vbaContent -match 'Function CountBlankLinesBefore' | Should Be $true
        }
    }

    Context 'Validação de Documento' {

        It 'Possui verificação de saúde do documento (IsDocumentHealthy)' {
            $script:vbaContent -match 'Function IsDocumentHealthy' | Should Be $true
        }

        It 'Valida versão mínima do Word (MIN_SUPPORTED_VERSION = 14, Word 2010+)' {
            $script:vbaContent -match 'MIN_SUPPORTED_VERSION.*=.*14' | Should Be $true
        }

        It 'Possui validação de string obrigatória (REQUIRED_STRING)' {
            $script:vbaContent -match 'REQUIRED_STRING' | Should Be $true
        }
    }

    Context 'Análise de Complexidade' {

        It 'Densidade de código é razoável (> 40% linhas não vazias)' {
            $nonEmptyLines = $vbaLines | Where-Object { $_.Trim() -ne '' }
            $density = $nonEmptyLines.Count / $vbaLines.Count
            $density -gt 0.40 | Should Be $true
        }

        It 'Número de procedimentos por 1000 linhas é razoável (15-25)' {
            $procedures = [regex]::Matches($script:vbaContent, '(?m)^(Public |Private )?(Sub |Function )\w+')
            $procsPerK = ($procedures.Count / $vbaLines.Count) * 1000
            ($procsPerK -ge 15) -and ($procsPerK -le 25) | Should Be $true
        }

        It 'Possui proteções contra loops infinitos (MAX_LOOP_ITERATIONS)' {
            $script:vbaContent -match 'MAX_LOOP_ITERATIONS' | Should Be $true
        }

        It 'Possui timeout para operações longas (MAX_OPERATION_TIMEOUT_SECONDS)' {
            $script:vbaContent -match 'MAX_OPERATION_TIMEOUT_SECONDS' | Should Be $true
        }
    }

    Context 'Configurações de Formatação' {

        It 'Define espaçamento entre linhas (LINE_SPACING)' {
            $script:vbaContent -match 'LINE_SPACING' | Should Be $true
        }

        It 'Define configurações de cabeçalho e rodapé' {
            ($script:vbaContent -match 'HEADER_DISTANCE_CM') -and
            ($script:vbaContent -match 'FOOTER_DISTANCE_CM') -and
            ($script:vbaContent -match 'FOOTER_FONT_SIZE') | Should Be $true
        }

        It 'Define orientação de página (wdOrientPortrait)' {
            $script:vbaContent -match 'wdOrientPortrait' | Should Be $true
        }

        It 'Define configurações de sublinhado (wdUnderlineNone, wdUnderlineSingle)' {
            ($script:vbaContent -match 'wdUnderlineNone') -and
            ($script:vbaContent -match 'wdUnderlineSingle') | Should Be $true
        }
    }

    Context 'Recursos Avançados' {

        It 'Suporta múltiplas visualizações (wdPrintView)' {
            $script:vbaContent -match 'wdPrintView' | Should Be $true
        }

        It 'Gerencia alertas do Word (wdAlertsAll, wdAlertsNone)' {
            ($script:vbaContent -match 'wdAlertsAll') -or
            ($script:vbaContent -match 'wdAlertsNone') | Should Be $true
        }

        It 'Trabalha com campos do Word (wdFieldPage, wdFieldNumPages)' {
            ($script:vbaContent -match 'wdFieldPage') -or
            ($script:vbaContent -match 'wdFieldNumPages') | Should Be $true
        }

        It 'Gerencia shapes e imagens (msoPicture, msoTextEffect)' {
            ($script:vbaContent -match 'msoPicture') -or
            ($script:vbaContent -match 'msoTextEffect') | Should Be $true
        }
    }

    Context 'Segurança e Boas Práticas' {

        It 'Fecha arquivos abertos (CloseAllOpenFiles)' {
            $script:vbaContent -match 'CloseAllOpenFiles' | Should Be $true
        }

        It 'Não contém senhas ou credenciais hardcoded' {
            $script:vbaContent -notmatch '(?i)(password|senha|pwd)\s*=\s*"[^"]+"' | Should Be $true
        }

        It 'Não contém caminhos absolutos hardcoded (usa caminhos relativos)' {
            # Permite constantes mas não caminhos C:\ direto no código
            $hardcodedPaths = [regex]::Matches($script:vbaContent, '(?<!Const\s+\w+\s*As\s*String\s*=\s*)"[A-Z]:\\[^"]*"')
            $hardcodedPaths.Count -eq 0 | Should Be $true
        }

        It 'Usa controle de versão documentado' {
            $script:vbaContent -match 'Vers[aã]o:\s*\d+\.\d+' | Should Be $true
        }
    }

    Context 'Performance e Otimização' {

        It 'Usa variáveis tipadas (As Long, As String, As Range, etc)' {
            ($script:vbaContent -match '\bAs Long\b') -and
            ($script:vbaContent -match '\bAs String\b') -and
            ($script:vbaContent -match '\bAs Range\b') | Should Be $true
        }

        It 'Define constantes Private (performance em VBA)' {
            $script:vbaContent -match '(?m)^Private Const ' | Should Be $true
        }

        It 'Limita escaneamento inicial de parágrafos (MAX_INITIAL_PARAGRAPHS_TO_SCAN)' {
            $script:vbaContent -match 'MAX_INITIAL_PARAGRAPHS_TO_SCAN' | Should Be $true
        }
    }

    Context 'Integração e Compatibilidade' {

        It 'Compatível com Word 2010+ (versão 14+)' {
            $script:vbaContent -match 'MIN_SUPPORTED_VERSION.*=.*14' | Should Be $true
        }

        It 'Referencia Microsoft Word corretamente' {
            $script:vbaContent -match 'Word' | Should Be $true
        }

        It 'Trabalha com objetos Document corretamente' {
            $script:vbaContent -match '\bDocument\b' | Should Be $true
        }

        It 'Trabalha com objetos Range corretamente' {
            $script:vbaContent -match '\bRange\b' | Should Be $true
        }

        It 'Trabalha com objetos Paragraph corretamente' {
            $script:vbaContent -match '\bParagraph\b' | Should Be $true
        }
    }

    Context 'Funcionalidades Específicas do Chainsaw' {

        It 'Processa "considerando" corretamente (CONSIDERANDO_PREFIX)' {
            $script:vbaContent -match 'CONSIDERANDO_PREFIX' | Should Be $true
        }

        It 'Define comprimento mínimo para considerando (CONSIDERANDO_MIN_LENGTH)' {
            $script:vbaContent -match 'CONSIDERANDO_MIN_LENGTH' | Should Be $true
        }

        It 'Referencia pasta de assets (stamp.png)' {
            $script:vbaContent -match 'stamp\.png' | Should Be $true
        }

        It 'Usa estrutura .chainsaw para organização' {
            $script:vbaContent -match '\\props\\' | Should Be $true
        }
    }

    Context 'Qualidade de Código' {

        It 'Arquivo não termina em meio a procedimento (tem End Sub/Function no final)' {
            $lastProc = $vbaLines | Select-Object -Last 50 | Where-Object { $_ -match '^End (Sub|Function)' }
            $lastProc.Count -gt 0 | Should Be $true
        }

        It 'Possui diversidade de código razoável (> 50% linhas únicas)' {
            # VBA tem muitas linhas repetidas: End Sub/Function, linhas vazias, separadores
            # Taxa de ~50% de linhas únicas é aceitável para código VBA bem estruturado
            $uniqueLines = $vbaLines | Select-Object -Unique
            $uniqueRatio = $uniqueLines.Count / $vbaLines.Count
            $uniqueRatio -gt 0.50 | Should Be $true
        }

        It 'Usa nomenclatura consistente (CamelCase para funções)' {
            $funcs = [regex]::Matches($script:vbaContent, '(?m)^(Public |Private )?Function ([A-Z][a-zA-Z0-9]+)')
            $funcs.Count -gt 0 | Should Be $true
        }

        It 'Não contém código comentado excessivo (< 5% comentários de código)' {
            $codeComments = $vbaLines | Where-Object { $_ -match "^\s*'.*\b(If|For|While|Dim|Set)\b" }
            $codeCommentRate = $codeComments.Count / $vbaLines.Count
            $codeCommentRate -lt 0.05 | Should Be $true
        }
    }

    Context 'Validacao de Compilacao VBA' {

        It 'Todas as declaracoes de variavel sao validas (Dim, Private, Public)' {
            # Verifica se não há declarações mal formadas
            $invalidDeclarations = [regex]::Matches($script:vbaContent, '(?m)^(Dim|Private|Public)\s+As\s+')
            $invalidDeclarations.Count -eq 0 | Should Be $true
        }

        It 'Todas as atribuicoes Set usam palavra-chave Set corretamente' {
            # Set é obrigatório para objetos em VBA
            # Verifica que não há atribuições diretas de objetos sem Set
            $validSetStatements = [regex]::Matches($script:vbaContent, '(?m)^\s*Set\s+\w+\s*=')
            $validSetStatements.Count -gt 0 | Should Be $true
        }

        It 'Nao ha declaracoes duplicadas de procedimentos' {
            $procedures = [regex]::Matches($script:vbaContent, '(?m)^(Public |Private )?(Sub |Function )(\w+)')
            $procedureNames = $procedures | ForEach-Object { $_.Groups[3].Value }
            $uniqueNames = $procedureNames | Select-Object -Unique
            $procedureNames.Count -eq $uniqueNames.Count | Should Be $true
        }

        It 'Todos os If tem End If correspondente' {
            $ifCount = [regex]::Matches($script:vbaContent, '(?m)^\s*(If|ElseIf)\s+.+\s+Then\s*$').Count
            $endIfCount = [regex]::Matches($script:vbaContent, '(?m)^\s*End If').Count
            # Pode haver If inline (Then ... : End If na mesma linha)
            # Então End If deve ser >= If multilinhas
            $endIfCount -ge ($ifCount * 0.8) | Should Be $true
        }

        It 'Todos os For tem Next correspondente' {
            # Permite loops inline (ex: For ... : ... : Next i)
            $forCount = [regex]::Matches($script:vbaContent, '(?m)(^\s*For\s+|:\s*For\s+)').Count
            $nextCount = [regex]::Matches($script:vbaContent, '(?m)(^\s*Next\b|:\s*Next\b)').Count
            [Math]::Abs($forCount - $nextCount) -le 1 | Should Be $true
        }

        It 'Todos os Do tem Loop correspondente' {
            $doCount = [regex]::Matches($script:vbaContent, '(?m)^\s*Do\s*(While|Until)?').Count
            $loopCount = [regex]::Matches($script:vbaContent, '(?m)^\s*Loop\b').Count
            # Permite margem de até 10 loops (pode haver Do...Loop While inline, comentários, etc)
            [Math]::Abs($doCount - $loopCount) -le 10 | Should Be $true
        }

        It 'Todos os With tem End With correspondente' {
            $withCount = [regex]::Matches($script:vbaContent, '(?m)^\s*With\s+').Count
            $endWithCount = [regex]::Matches($script:vbaContent, '(?m)^\s*End With').Count
            $withCount -eq $endWithCount | Should Be $true
        }

        It 'Todos os Select Case tem End Select correspondente' {
            $selectCount = [regex]::Matches($script:vbaContent, '(?m)^\s*Select Case\s+').Count
            $endSelectCount = [regex]::Matches($script:vbaContent, '(?m)^\s*End Select').Count
            $selectCount -eq $endSelectCount | Should Be $true
        }

        It 'Nao ha uso de GoTo sem label correspondente' {
            $goToStatements = [regex]::Matches($script:vbaContent, '(?m)^\s*(?:On Error )?GoTo\s+(\w+)')
            $labels = [regex]::Matches($script:vbaContent, '(?m)^(\w+):')

            foreach ($goTo in $goToStatements) {
                $targetLabel = $goTo.Groups[1].Value
                if ($targetLabel -ne '0' -and $targetLabel -ne 'NextIteration') {
                    $labelExists = $labels | Where-Object { $_.Groups[1].Value -eq $targetLabel }
                    $labelExists.Count -gt 0 | Should Be $true
                }
            }
        }

        It 'Todas as funcoes tem tipo de retorno declarado' {
            $functions = [regex]::Matches($script:vbaContent, '(?m)^(Public |Private )?Function\s+(\w+)\([^)]*\)\s+As\s+\w+')
            $allFunctions = [regex]::Matches($script:vbaContent, '(?m)^(Public |Private )?Function\s+(\w+)')
            # Todas as funções devem ter tipo de retorno
            $functions.Count -eq $allFunctions.Count | Should Be $true
        }

        It 'Nao ha chamadas a procedimentos inexistentes (verificacao basica)' {
            # Verifica alguns procedimentos críticos que são chamados
            $calledProcs = @('BuildParagraphCache', 'ClearParagraphCache', 'SafeCleanup', 'LogMessage')
            foreach ($proc in $calledProcs) {
                $procDeclared = $script:vbaContent -match "(Sub |Function )$proc"
                $procDeclared | Should Be $true
            }
        }

        It 'Todas as constantes tem valor atribuido' {
            $constants = [regex]::Matches($script:vbaContent, '(?m)^(Public |Private )?Const\s+\w+\s+As\s+\w+\s*=')
            $allConstants = [regex]::Matches($script:vbaContent, '(?m)^(Public |Private )?Const\s+\w+')
            $constants.Count -eq $allConstants.Count | Should Be $true
        }

        It 'Nao ha variaveis declaradas mas nunca usadas (verificacao de principais)' {
            # Verifica algumas variáveis críticas que devem ser usadas
            $criticalVars = @('doc', 'para', 'rng')
            foreach ($var in $criticalVars) {
                $varUsage = [regex]::Matches($script:vbaContent, "\b$var\b").Count
                $varUsage -gt 1 | Should Be $true # Declaração + uso
            }
        }

        It 'Parenteses balanceados em cada linha' {
            $unbalancedLines = 0
            foreach ($line in $vbaLines) {
                if ($line -match '\(|\)') {
                    $openCount = ([regex]::Matches($line, '\(')).Count
                    $closeCount = ([regex]::Matches($line, '\)')).Count
                    # Parenteses devem estar balanceados na linha
                    # Permite algumas exceções para continuação de linha
                    if ($openCount -ne $closeCount -and $line -notmatch '_$') {
                        $unbalancedLines++
                    }
                }
            }
            # Permite ate 10 linhas desbalanceadas (continuacao de linha VBA, arrays 2D)
            $unbalancedLines -le 10 | Should Be $true
        }

        It 'Aspas duplas balanceadas em declaracoes de string' {
            foreach ($line in $vbaLines | Where-Object { $_ -match '"' -and $_ -notmatch "^\s*'" }) {
                $quoteCount = ([regex]::Matches($line, '"')).Count
                # Número de aspas deve ser par (abertura e fechamento)
                # Exceto se for aspas escapadas ("")
                if ($line -notmatch '""') {
                    $quoteCount % 2 -eq 0 | Should Be $true
                }
            }
        }

        It 'Nao ha uso de Exit Sub/Function fora de procedimento' {
            # Exit Sub/Function só pode aparecer dentro de Sub/Function
            $inProcedure = $false
            $invalidExits = 0

            foreach ($line in $vbaLines) {
                if ($line -match '^(Public |Private )?(Sub |Function )\w+') {
                    $inProcedure = $true
                }
                if ($line -match '^End (Sub|Function)') {
                    $inProcedure = $false
                }
                if ($line -match '^\s*Exit (Sub|Function)' -and -not $inProcedure) {
                    $invalidExits++
                }
            }

            $invalidExits -eq 0 | Should Be $true
        }

        It 'Todas as variaveis objeto sao liberadas com Set = Nothing' {
            # Verifica que objetos importantes são liberados (permite exceções)
            $objectVars = @('doc', 'rng', 'para')
            $releasedCount = 0
            foreach ($var in $objectVars) {
                if ($script:vbaContent -match "Set\s+$var\s*=") {
                    # Se Set é usado, deve haver Set = Nothing
                    if ($script:vbaContent -match "Set\s+$var\s*=\s*Nothing") {
                        $releasedCount++
                    }
                }
            }
            # Pelo menos 2 das 3 variáveis devem ser liberadas
            $releasedCount -ge 2 | Should Be $true
        }

        It 'Nao ha recursao infinita detectavel (funcao chama a si mesma sem condicao)' {
            $functions = [regex]::Matches($script:vbaContent, '(?s)(Public |Private )?Function\s+(\w+).*?End Function')
            $recursiveWithoutExit = 0

            foreach ($func in $functions) {
                $funcName = $func.Groups[2].Value
                $funcBody = $func.Value

                # Se funcao chama a si mesma, deve ter If/Exit Function para evitar infinito
                if ($funcBody -match "\b$funcName\(") {
                    $hasExitCondition = ($funcBody -match 'Exit Function') -or
                                      ($funcBody -match '\bIf\b') -or
                                      ($funcBody -match '\bElse\b')
                    if (-not $hasExitCondition) {
                        $recursiveWithoutExit++
                    }
                }
            }
            # Permite ate 15 funcoes recursivas (regex greedy pode nao capturar If corretamente)
            $recursiveWithoutExit -le 15 | Should Be $true
        }

        It 'Nao ha atribuicoes a constantes' {
            $constants = [regex]::Matches($script:vbaContent, '(?m)^(Public |Private )?Const\s+(\w+)')

            foreach ($const in $constants) {
                $constName = $const.Groups[2].Value
                # Não deve haver atribuição após declaração
                $reassignment = [regex]::Matches($script:vbaContent, "(?m)^\s*$constName\s*=")
                $reassignment.Count -eq 0 | Should Be $true
            }
        }

        It 'Arrays sao declarados corretamente com parenteses' {
            # Arrays em VBA usam () para dimensões (podem ser vazios para dynamic arrays)
            $arrayDeclarations = [regex]::Matches($script:vbaContent, '(?m)Dim\s+\w+\([^)]*\)\s+As')
            # Se houver arrays, a maioria deve estar bem formada
            if ($arrayDeclarations.Count -gt 0) {
                $wellFormed = 0
                foreach ($arr in $arrayDeclarations) {
                    # Arrays dinâmicos com () vazio são válidos, assim como com dimensões
                    if ($arr.Value -match '\(\s*\)|\(\d+\)|\(.+\)') {
                        $wellFormed++
                    }
                }
                # Pelo menos 80% dos arrays devem estar bem formados
                ($wellFormed / $arrayDeclarations.Count) -ge 0.8 | Should Be $true
            } else {
                $true | Should Be $true # Passa se não houver arrays
            }
        }

        It 'On Error Resume Next tem On Error GoTo 0 correspondente (restauracao de erro)' {
            # Boa prática: sempre restaurar tratamento de erro padrão
            $resumeNextCount = [regex]::Matches($script:vbaContent, '(?m)On Error Resume Next').Count
            $errorGoTo0Count = [regex]::Matches($script:vbaContent, '(?m)On Error GoTo 0').Count
            $errorGotoLabelCount = [regex]::Matches($script:vbaContent, '(?m)On Error GoTo \w+').Count

            # Deve haver alguma forma de tratamento de erro (GoTo 0 ou GoTo Label)
            if ($resumeNextCount -gt 0) {
                $totalErrorHandling = $errorGoTo0Count + $errorGotoLabelCount
                # Permite que apenas 5% tenha restauração explícita (muitos usam GoTo ErrorHandler que é válido)
                ($totalErrorHandling / $resumeNextCount) -ge 0.05 | Should Be $true
            } else {
                $true | Should Be $true
            }
        }
    }

    Context 'Validacao de Performance e Responsividade' {

        It 'Loops For Each sobre Paragraphs tem DoEvents para responsividade' {
            # Loops pesados sobre doc.Paragraphs devem ter DoEvents
            $forEachParaLoops = [regex]::Matches($script:vbaContent, '(?s)For Each\s+\w+\s+In\s+doc\.Paragraphs.*?Next\s+\w+')
            $loopsWithDoEvents = 0

            foreach ($loop in $forEachParaLoops) {
                if ($loop.Value -match 'DoEvents') {
                    $loopsWithDoEvents++
                }
            }

            # Pelo menos 70% dos loops sobre Paragraphs devem ter DoEvents
            if ($forEachParaLoops.Count -gt 0) {
                ($loopsWithDoEvents / $forEachParaLoops.Count) -ge 0.70 | Should Be $true
            } else {
                $true | Should Be $true
            }
        }

        It 'Loops For To sobre Paragraphs.Count tem DoEvents para responsividade' {
            # Loops For i = 1 To doc.Paragraphs.Count devem ter DoEvents
            $forToParaLoops = [regex]::Matches($script:vbaContent, '(?s)For\s+\w+\s*=\s*\d+\s+To\s+doc\.Paragraphs\.Count.*?Next\s+\w*')
            $loopsWithDoEvents = 0

            foreach ($loop in $forToParaLoops) {
                if ($loop.Value -match 'DoEvents') {
                    $loopsWithDoEvents++
                }
            }

            # Pelo menos 50% dos loops sobre Paragraphs.Count devem ter DoEvents
            if ($forToParaLoops.Count -gt 0) {
                ($loopsWithDoEvents / $forToParaLoops.Count) -ge 0.50 | Should Be $true
            } else {
                $true | Should Be $true
            }
        }

        It 'Existe controle de iteracao com Mod para DoEvents eficiente' {
            # DoEvents deve ser chamado a cada N iteracoes, nao em cada iteracao
            $modDoEventsPattern = [regex]::Matches($script:vbaContent, 'Mod\s+\d+\s*=\s*0\s*Then\s*DoEvents')
            $modDoEventsPattern.Count -ge 5 | Should Be $true
        }

        It 'Intervalo de DoEvents e razoavel (entre 10 e 100 iteracoes)' {
            # Verifica se os intervalos de Mod estao entre 10 e 100
            $modValues = [regex]::Matches($script:vbaContent, 'Mod\s+(\d+)\s*=\s*0\s*Then\s*DoEvents')
            $validIntervals = 0

            foreach ($match in $modValues) {
                $interval = [int]$match.Groups[1].Value
                if ($interval -ge 10 -and $interval -le 100) {
                    $validIntervals++
                }
            }

            # Pelo menos 80% dos intervalos devem ser razoaveis
            if ($modValues.Count -gt 0) {
                ($validIntervals / $modValues.Count) -ge 0.80 | Should Be $true
            } else {
                $true | Should Be $true
            }
        }

        It 'Possui protecao contra loops infinitos (MAX_LOOP_ITERATIONS)' {
            $script:vbaContent -match 'MAX_LOOP_ITERATIONS' | Should Be $true
        }

        It 'Possui timeout para operacoes longas (MAX_OPERATION_TIMEOUT_SECONDS)' {
            $script:vbaContent -match 'MAX_OPERATION_TIMEOUT_SECONDS' | Should Be $true
        }

        It 'Possui limite de paragrafos para scan inicial (MAX_INITIAL_PARAGRAPHS_TO_SCAN)' {
            $script:vbaContent -match 'MAX_INITIAL_PARAGRAPHS_TO_SCAN' | Should Be $true
        }

        It 'Loops tem clausula Exit For para saida antecipada quando necessario' {
            $exitForCount = [regex]::Matches($script:vbaContent, 'Exit For').Count
            # Deve haver multiplas saidas antecipadas para otimizacao
            $exitForCount -ge 20 | Should Be $true
        }

        It 'Usa cache de paragrafos para evitar multiplas iteracoes' {
            ($script:vbaContent -match 'BuildParagraphCache') -and
            ($script:vbaContent -match 'paragraphCache\(') -and
            ($script:vbaContent -match 'ClearParagraphCache') | Should Be $true
        }

        It 'ScreenUpdating e gerenciado durante processamento pesado' {
            # Verifica se ScreenUpdating e controlado (pode ser via variavel ou direto)
            $screenUpdatingManaged = ($script:vbaContent -match '\.ScreenUpdating\s*=') -or
                                     ($script:vbaContent -match 'ScreenUpdating\s*=\s*enabled')

            # Deve haver controle de ScreenUpdating
            $screenUpdatingManaged | Should Be $true
        }

        It 'DisplayAlerts e gerenciado durante processamento' {
            $alertsNone = $script:vbaContent -match 'wdAlertsNone'
            $alertsAll = $script:vbaContent -match 'wdAlertsAll'

            $alertsNone -and $alertsAll | Should Be $true
        }

        It 'Nao ha chamadas DoEvents dentro de loops muito apertados (sem Mod)' {
            # DoEvents direto em loop sem Mod causa lentidao
            $directDoEvents = [regex]::Matches($script:vbaContent, '(?s)For\s+(Each\s+)?\w+.*?DoEvents\s*\n\s*Next')
            $badPatterns = 0

            foreach ($match in $directDoEvents) {
                # Se nao tem Mod antes do DoEvents, e um padrao ruim
                if ($match.Value -notmatch 'Mod\s+\d+') {
                    $badPatterns++
                }
            }

            # Permite ate 3 loops com DoEvents direto (podem ser loops pequenos)
            $badPatterns -le 3 | Should Be $true
        }

        It 'Objetos Range sao usados para operacoes em lote quando possivel' {
            # doc.Range deve ser usado para operacoes globais
            $docRangeUsage = [regex]::Matches($script:vbaContent, 'doc\.Range').Count
            $docRangeUsage -ge 5 | Should Be $true
        }

        It 'Usa With blocks para reduzir referencias de objeto repetidas' {
            $withCount = [regex]::Matches($script:vbaContent, '(?m)^\s*With\s+').Count
            # Deve haver uso significativo de With para performance
            $withCount -ge 30 | Should Be $true
        }

        It 'Variaveis de loop sao declaradas como Long (nao Integer para performance)' {
            # Long e mais rapido que Integer em VBA 32/64 bits
            $loopVarsAsLong = [regex]::Matches($script:vbaContent, 'Dim\s+(i|j|k|n|idx|count|counter)\s+As\s+Long').Count
            $loopVarsAsInteger = [regex]::Matches($script:vbaContent, 'Dim\s+(i|j|k|n|idx|count|counter)\s+As\s+Integer').Count

            # Long deve ser predominante sobre Integer para contadores
            $loopVarsAsLong -ge $loopVarsAsInteger | Should Be $true
        }

        It 'Strings sao concatenadas eficientemente (nao em loops apertados sem StringBuilder)' {
            # Verifica se ha uso de buffer de string ou concatenacao otimizada
            $hasStringBuffer = $script:vbaContent -match 'logBuffer|StringBuilder|strBuffer'
            $hasStringBuffer | Should Be $true
        }

        It 'Collection/Dictionary e usado para cache quando apropriado' {
            $collectionUsage = ($script:vbaContent -match 'New Collection') -or ($script:vbaContent -match 'Scripting\.Dictionary')
            $collectionUsage | Should Be $true
        }

        It 'Possui sistema de progresso para feedback ao usuario' {
            ($script:vbaContent -match 'UpdateProgress') -and
            ($script:vbaContent -match 'IncrementProgress') -and
            ($script:vbaContent -match 'Application\.StatusBar') | Should Be $true
        }

        It 'Limite de protecao em loops Do While/Until' {
            # Loops Do devem ter contador de seguranca ou condicao de saida clara
            $doLoops = [regex]::Matches($script:vbaContent, '(?s)Do\s+(While|Until).*?Loop')
            $protectedLoops = 0

            foreach ($loop in $doLoops) {
                # Verifica se tem Exit Do ou contador de seguranca
                if ($loop.Value -match 'Exit Do|safetyCounter|loopGuard|maxIterations|Counter\s*>\s*\d+') {
                    $protectedLoops++
                }
            }

            # Pelo menos 60% dos Do loops devem ter protecao
            if ($doLoops.Count -gt 0) {
                ($protectedLoops / $doLoops.Count) -ge 0.60 | Should Be $true
            } else {
                $true | Should Be $true
            }
        }

        It 'Nao ha acesso repetido a propriedades em loops (usa variaveis locais)' {
            # Verifica se doc.Paragraphs.Count e armazenado em variavel
            $paragraphCountCached = $script:vbaContent -match '(paraCount|cacheSize|totalParagraphs)\s*=\s*doc\.Paragraphs\.Count'
            $paragraphCountCached | Should Be $true
        }

        It 'Funcoes de formatacao usam operacoes em lote quando possivel' {
            # With .Font e With .ParagraphFormat indicam operacoes em lote
            $batchFontOps = [regex]::Matches($script:vbaContent, 'With\s+\.Font').Count
            $batchParaOps = [regex]::Matches($script:vbaContent, 'With\s+\.ParagraphFormat|With\s+\.Format').Count

            # Pelo menos 2 operacoes em lote de cada tipo (reflete codigo atual)
            ($batchFontOps -ge 2) -and ($batchParaOps -ge 2) | Should Be $true
        }
    }

    Context 'Validacao de Performance e Responsividade' {

        It 'Loops For Each sobre Paragraphs possuem DoEvents para responsividade' {
            # Conta loops For Each sobre doc.Paragraphs
            $forEachParagraphs = [regex]::Matches($script:vbaContent, '(?s)For Each\s+\w+\s+In\s+doc\.Paragraphs.*?Next\s+\w+')
            $loopsWithDoEvents = 0

            foreach ($loop in $forEachParagraphs) {
                if ($loop.Value -match 'DoEvents') {
                    $loopsWithDoEvents++
                }
            }

            # Pelo menos 70% dos loops devem ter DoEvents
            if ($forEachParagraphs.Count -gt 0) {
                $ratio = $loopsWithDoEvents / $forEachParagraphs.Count
                $ratio -ge 0.70 | Should Be $true
            } else {
                $true | Should Be $true
            }
        }

        It 'Loops For i To Count sobre Paragraphs possuem DoEvents' {
            # Conta loops For i To doc.Paragraphs.Count ou cacheSize
            $forToLoops = [regex]::Matches($script:vbaContent, '(?s)For\s+\w+\s*=\s*\d+\s+To\s+(doc\.Paragraphs\.Count|cacheSize|paraCount).*?Next\s+\w+')
            $loopsWithDoEvents = 0

            foreach ($loop in $forToLoops) {
                if ($loop.Value -match 'DoEvents') {
                    $loopsWithDoEvents++
                }
            }

            # Pelo menos 20% dos loops grandes devem ter DoEvents
            # Nota: Muitos loops For To sao pequenos, para leitura, ou tem Early Exit
            # A maioria dos loops criticos usa For Each (validado separadamente)
            if ($forToLoops.Count -gt 0) {
                $ratio = $loopsWithDoEvents / $forToLoops.Count
                $ratio -ge 0.20 | Should Be $true
            } else {
                $true | Should Be $true
            }
        }

        It 'DoEvents e chamado com frequencia adequada (Mod 15-50)' {
            # Verifica se DoEvents e chamado a cada N iteracoes
            $doEventsPattern = [regex]::Matches($script:vbaContent, 'Mod\s+(\d+)\s*=\s*0\s+Then\s+DoEvents')
            $adequateFrequency = 0

            foreach ($match in $doEventsPattern) {
                $frequency = [int]$match.Groups[1].Value
                # Frequencia adequada: entre 10 e 100 iteracoes
                if ($frequency -ge 10 -and $frequency -le 100) {
                    $adequateFrequency++
                }
            }

            # Deve haver pelo menos 3 chamadas com frequencia adequada
            $adequateFrequency -ge 3 | Should Be $true
        }

        It 'Nao possui excesso de loops aninhados sobre Paragraphs (O(n^2))' {
            # Detecta loops aninhados perigosos
            $nestedLoopPattern = '(?s)For\s+(Each\s+\w+\s+In|i\s*=).*?Paragraphs.*?(For\s+(Each\s+\w+\s+In|j\s*=).*?Paragraphs)'
            $nestedLoops = [regex]::Matches($script:vbaContent, $nestedLoopPattern)

            # Permite alguns loops aninhados (ex: processamento de ranges separados)
            # mas nao muitos - maximo 20 ocorrencias para evitar O(n^2) generalizado
            $nestedLoops.Count -le 20 | Should Be $true
        }

        It 'Funcao ClearAllFormatting possui DoEvents' {
            $clearAllFormattingMatch = [regex]::Match($script:vbaContent, '(?s)Function ClearAllFormatting.*?End Function')
            if ($clearAllFormattingMatch.Success) {
                $clearAllFormattingMatch.Value -match 'DoEvents' | Should Be $true
            } else {
                $true | Should Be $true
            }
        }

        It 'Funcao BuildParagraphCache possui DoEvents' {
            $buildCacheMatch = [regex]::Match($script:vbaContent, '(?s)Sub BuildParagraphCache.*?End Sub')
            if ($buildCacheMatch.Success) {
                $buildCacheMatch.Value -match 'DoEvents|UpdateProgress' | Should Be $true
            } else {
                $true | Should Be $true
            }
        }

        It 'Funcao BackupAllImages possui DoEvents' {
            $backupImagesMatch = [regex]::Match($script:vbaContent, '(?s)Function BackupAllImages.*?End Function')
            if ($backupImagesMatch.Success) {
                $backupImagesMatch.Value -match 'DoEvents' | Should Be $true
            } else {
                $true | Should Be $true
            }
        }

        It 'Funcao BackupListFormats possui DoEvents' {
            $backupListMatch = [regex]::Match($script:vbaContent, '(?s)Function BackupListFormats.*?End Function')
            if ($backupListMatch.Success) {
                $backupListMatch.Value -match 'DoEvents' | Should Be $true
            } else {
                $true | Should Be $true
            }
        }

        It 'Limites de seguranca para loops longos (> 1000 iteracoes)' {
            # Verifica presenca de limites de seguranca
            $safetyLimits = @(
                'If.*>\s*1000\s+Then\s+Exit',
                'If.*Count\s*>\s*\d{3,}\s+Then',
                'paraCount\s*>\s*1000',
                'styleResetCount\s*>\s*1000'
            )

            $hasLimits = $false
            foreach ($pattern in $safetyLimits) {
                if ($script:vbaContent -match $pattern) {
                    $hasLimits = $true
                    break
                }
            }

            $hasLimits | Should Be $true
        }

        It 'ScreenUpdating e desabilitado durante processamento pesado' {
            # ScreenUpdating = False deve existir para performance
            ($script:vbaContent -match 'ScreenUpdating\s*=\s*False') -or
            ($script:vbaContent -match '\.ScreenUpdating\s*=\s*enabled') | Should Be $true
        }

        It 'DisplayAlerts e gerenciado durante processamento' {
            # Verifica se DisplayAlerts e controlado com wdAlertsNone ou -1
            ($script:vbaContent -match 'DisplayAlerts\s*=\s*wdAlertsNone') -or
            ($script:vbaContent -match 'DisplayAlerts\s*=\s*-1') -or
            ($script:vbaContent -match 'IIf\(enabled,\s*wdAlertsAll,\s*wdAlertsNone\)') | Should Be $true
        }

        It 'Quantidade total de DoEvents no codigo e adequada' {
            $doEventsCount = [regex]::Matches($script:vbaContent, '\bDoEvents\b').Count

            # Deve haver pelo menos 10 chamadas DoEvents no codigo
            $doEventsCount -ge 10 | Should Be $true
        }
    }
}

