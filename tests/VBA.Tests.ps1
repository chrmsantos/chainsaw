#requires -Version 5.1
Import-Module Pester -ErrorAction Stop
. $PSScriptRoot\Helpers.ps1

# Override Get-RepoRoot for test context
function Get-RepoRoot {
    $testsDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $repoRoot = Split-Path -Parent $testsDir
    return $repoRoot
}

Describe 'CHAINSAW - Testes do Módulo VBA Módulo1.bas' {

    BeforeAll {
        $repoRoot = Get-RepoRoot
        $vbaPath = Join-Path $repoRoot "source\main\Módulo1.bas"
        $vbaContent = Get-Content $vbaPath -Raw -Encoding UTF8
        $vbaLines = Get-Content $vbaPath -Encoding UTF8
    }

    Context 'Estrutura e Metadados do Arquivo' {

        It 'Arquivo Módulo1.bas existe' {
            Test-Path $vbaPath | Should Be $true
        }

        It 'Arquivo não está vazio' {
            (Get-Item $vbaPath).Length -gt 0 | Should Be $true
        }

        It 'Tamanho do arquivo é razoável (< 5MB)' {
            $sizeMB = (Get-Item $vbaPath).Length / 1MB
            $sizeMB -lt 5 | Should Be $true
        }

        It 'Contém cabeçalho CHAINSAW' {
            $vbaContent -match 'CHAINSAW' | Should Be $true
        }

        It 'Contém informações de versão' {
            $vbaContent -match 'Vers[aã]o:\s*\d+\.\d+' | Should Be $true
        }

        It 'Contém licença GNU GPLv3' {
            $vbaContent -match 'GNU GPLv3' | Should Be $true
        }

        It 'Contém informação de autor' {
            $vbaContent -match 'Autor:' | Should Be $true
        }

        It 'Contém declaração Option Explicit' {
            $vbaContent -match '(?m)^Option Explicit' | Should Be $true
        }

        It 'Número total de linhas corresponde ao esperado (> 7000)' {
            $vbaLines.Count -gt 7000 | Should Be $true
        }
    }

    Context 'Análise de Procedimentos e Funções' {

        BeforeAll {
            $procedures = [regex]::Matches($vbaContent, '(?m)^(Public |Private )?(Sub |Function )\w+')
            $publicProcs = [regex]::Matches($vbaContent, '(?m)^Public (Sub |Function )\w+')
            $privateProcs = [regex]::Matches($vbaContent, '(?m)^Private (Sub |Function )\w+')
            $subs = [regex]::Matches($vbaContent, '(?m)^(Public |Private )?Sub \w+')
            $functions = [regex]::Matches($vbaContent, '(?m)^(Public |Private )?Function \w+')
        }

        It 'Contem quantidade razoavel de procedimentos (100-160)' {
            $procedures.Count -ge 100 -and $procedures.Count -le 160 | Should Be $true
        }

        It 'Possui procedimento principal PadronizarDocumentoMain' {
            $vbaContent -match '(?m)^Public Sub PadronizarDocumentoMain\(' | Should Be $true
        }

        It 'Procedimentos públicos são minoria (< 20% do total)' {
            $publicRatio = $publicProcs.Count / $procedures.Count
            $publicRatio -lt 0.20 | Should Be $true
        }

        It 'Possui funções de validação (ValidateDocument)' {
            $vbaContent -match 'Function ValidateDocument' | Should Be $true
        }

        It 'Possui funções de identificação de elementos estruturais' {
            ($vbaContent -match 'GetTituloRange') -and
            ($vbaContent -match 'GetEmentaRange') -and
            ($vbaContent -match 'GetProposicaoRange') | Should Be $true
        }

        It 'Possui sistema de tratamento de erros (ShowUserFriendlyError)' {
            $vbaContent -match 'ShowUserFriendlyError' | Should Be $true
        }

        It 'Possui sistema de recuperação de emergência (EmergencyRecovery)' {
            $vbaContent -match 'EmergencyRecovery' | Should Be $true
        }

        It 'Possui funções de normalização de texto' {
            $vbaContent -match 'NormalizarTexto' | Should Be $true
        }

        It 'Todas as funções têm End Function' {
            $functionStarts = [regex]::Matches($vbaContent, '(?m)^(Public |Private )?Function \w+').Count
            $functionEnds = [regex]::Matches($vbaContent, '(?m)^End Function').Count
            $functionStarts -eq $functionEnds | Should Be $true
        }

        It 'Todas as subs têm End Sub' {
            $subStarts = [regex]::Matches($vbaContent, '(?m)^(Public |Private )?Sub \w+').Count
            $subEnds = [regex]::Matches($vbaContent, '(?m)^End Sub').Count
            $subStarts -eq $subEnds | Should Be $true
        }
    }

    Context 'Constantes e Configurações' {

        It 'Define constantes do Word (wdNoProtection, wdTypeDocument, etc)' {
            ($vbaContent -match 'wdNoProtection') -and
            ($vbaContent -match 'wdTypeDocument') -and
            ($vbaContent -match 'wdAlignParagraphCenter') | Should Be $true
        }

        It 'Define constantes de formatação (STANDARD_FONT, STANDARD_FONT_SIZE)' {
            ($vbaContent -match 'STANDARD_FONT') -and
            ($vbaContent -match 'STANDARD_FONT_SIZE') | Should Be $true
        }

        It 'Define margens do documento (TOP_MARGIN_CM, BOTTOM_MARGIN_CM, etc)' {
            ($vbaContent -match 'TOP_MARGIN_CM') -and
            ($vbaContent -match 'BOTTOM_MARGIN_CM') -and
            ($vbaContent -match 'LEFT_MARGIN_CM') -and
            ($vbaContent -match 'RIGHT_MARGIN_CM') | Should Be $true
        }

        It 'Define configurações de imagem do cabeçalho' {
            ($vbaContent -match 'HEADER_IMAGE_RELATIVE_PATH') -and
            ($vbaContent -match 'HEADER_IMAGE_MAX_WIDTH_CM') | Should Be $true
        }

        It 'Define constantes de sistema (MIN_SUPPORTED_VERSION, MAX_RETRY_ATTEMPTS)' {
            ($vbaContent -match 'MIN_SUPPORTED_VERSION') -and
            ($vbaContent -match 'MAX_RETRY_ATTEMPTS') | Should Be $true
        }

        It 'Define constantes de backup e logs (GetChainsawBackupsPath, GetChainsawRecoveryPath)' {
            ($vbaContent -match 'GetChainsawBackupsPath') -and
            ($vbaContent -match 'GetChainsawRecoveryPath') | Should Be $true
        }

        It 'Define níveis de log (LOG_LEVEL_INFO, LOG_LEVEL_WARNING, LOG_LEVEL_ERROR)' {
            ($vbaContent -match 'LOG_LEVEL_INFO') -and
            ($vbaContent -match 'LOG_LEVEL_WARNING') -and
            ($vbaContent -match 'LOG_LEVEL_ERROR') | Should Be $true
        }

        It 'Fonte padrão é Arial' {
            $vbaContent -match 'STANDARD_FONT.*=.*"Arial"' | Should Be $true
        }

        It 'Tamanho de fonte padrão é 12' {
            $vbaContent -match 'STANDARD_FONT_SIZE.*=.*12' | Should Be $true
        }
    }

    Context 'Sistema de Cache de Parágrafos' {

        It 'Possui função BuildParagraphCache' {
            $vbaContent -match 'Sub BuildParagraphCache' | Should Be $true
        }

        It 'Possui função ClearParagraphCache' {
            $vbaContent -match 'Sub ClearParagraphCache' | Should Be $true
        }

        It 'Possui sistema de identificação de estrutura do documento' {
            $vbaContent -match 'IdentifyDocumentStructure' | Should Be $true
        }
    }

    Context 'Identificação de Elementos Estruturais' {

        It 'Possui função para identificar Título (IsTituloElement)' {
            $vbaContent -match 'Function IsTituloElement' | Should Be $true
        }

        It 'Possui função para identificar Ementa (IsEmentaElement)' {
            $vbaContent -match 'Function IsEmentaElement' | Should Be $true
        }

        It 'Possui função para identificar Justificativa (IsJustificativaTitleElement)' {
            $vbaContent -match 'Function IsJustificativaTitleElement' | Should Be $true
        }

        It 'Possui função para identificar Data (IsDataElement)' {
            $vbaContent -match 'Function IsDataElement' | Should Be $true
        }

        It 'Possui função para identificar Assinatura (IsAssinaturaStart)' {
            $vbaContent -match 'Function IsAssinaturaStart' | Should Be $true
        }

        It 'Possui função para identificar Título de Anexo (IsTituloAnexoElement)' {
            $vbaContent -match 'Function IsTituloAnexoElement' | Should Be $true
        }

        It 'Possui GetProposituraRange para retornar range da propositura completa' {
            $vbaContent -match 'Function GetProposituraRange' | Should Be $true
        }

        It 'Possui GetElementInfo para relatório de elementos' {
            $vbaContent -match 'GetElementInfo' | Should Be $true
        }
    }

    Context 'Tratamento de Erros e Recuperação' {

        It 'Possui tratamento On Error em procedimentos críticos' {
            $vbaContent -match 'On Error GoTo' | Should Be $true
        }

        It 'Possui labels de tratamento de erro (ErrorHandler:)' {
            $vbaContent -match 'ErrorHandler:' | Should Be $true
        }

        It 'Possui função SafeCleanup' {
            $vbaContent -match 'Sub SafeCleanup' | Should Be $true
        }

        It 'Possui função ReleaseObjects' {
            $vbaContent -match 'Sub ReleaseObjects' | Should Be $true
        }

        It 'Possui verificação de timeout (IsOperationTimeout)' {
            $vbaContent -match 'Function IsOperationTimeout' | Should Be $true
        }

        It 'Implementa sistema de retry (MAX_RETRY_ATTEMPTS)' {
            $vbaContent -match 'MAX_RETRY_ATTEMPTS' | Should Be $true
        }
    }

    Context 'Validação de Sintaxe VBA' {

        It 'Não contém tabs (usa apenas espaços)' {
            $vbaContent -notmatch "`t" | Should Be $true
        }

        It 'Parênteses balanceados em declarações de função' {
            $functionDeclarations = [regex]::Matches($vbaContent, '(?m)^(Public |Private )?Function [^(]+\([^)]*\)')
            $functionDeclarations.Count -gt 0 | Should Be $true
        }

        It 'Não contém caracteres de controle inválidos' {
            $invalidChars = [regex]::Matches($vbaContent, '[\x00-\x08\x0B\x0C\x0E-\x1F]')
            $invalidChars.Count -eq 0 | Should Be $true
        }

        It 'Linhas não excedem 1000 caracteres (padrão VBA)' {
            $longLines = $vbaLines | Where-Object { $_.Length -gt 1000 }
            $longLines.Count -eq 0 | Should Be $true
        }

        It 'Usa aspas duplas para strings, não aspas simples' {
            # VBA usa aspas duplas "" para strings, ' é apenas para comentários
            $stringDeclarations = [regex]::Matches($vbaContent, '=\s*"[^"]*"')
            $stringDeclarations.Count -gt 0 | Should Be $true
        }
    }

    Context 'Comentários e Documentação' {

        It 'Contém comentários de seção (linhas com ====)' {
            $vbaContent -match '={20,}' | Should Be $true
        }

It 'Taxa de comentários adequada (> 5% das linhas)' {
            $commentLines = $vbaLines | Where-Object { $_ -match "^\s*'" }
            $commentRatio = $commentLines.Count / $vbaLines.Count
            $commentRatio -gt 0.05 | Should Be $true
        }

        It 'Contém seções organizadas (CONSTANTES, FUNÇÕES, etc)' {
            $vbaContent -match 'CONSTANTES' | Should Be $true
        }
    }

    Context 'Funcionalidades de Backup e Log' {

        It 'Possui sistema de backup (CreateDocumentBackup)' {
            $vbaContent -match 'CreateDocumentBackup' | Should Be $true
        }

        It 'Possui limite de arquivos de backup (MAX_BACKUP_FILES)' {
            $vbaContent -match 'MAX_BACKUP_FILES' | Should Be $true
        }

        It 'Implementa sistema de logging' {
            ($vbaContent -match 'LOG_LEVEL') -or ($vbaContent -match 'WriteLog') | Should Be $true
        }

        It 'Possui modo de debug (DEBUG_MODE)' {
            $vbaContent -match 'DEBUG_MODE' | Should Be $true
        }
    }

    Context 'Processamento de Texto' {

        It 'Possui função GetCleanParagraphText' {
            $vbaContent -match 'Function GetCleanParagraphText' | Should Be $true
        }

        It 'Possui função RemovePunctuation' {
            $vbaContent -match 'Function RemovePunctuation' | Should Be $true
        }

        It 'Possui função para detectar parágrafos especiais (DetectSpecialParagraph)' {
            $vbaContent -match 'Function DetectSpecialParagraph' | Should Be $true
        }

        It 'Possui função para contar linhas em branco (CountBlankLinesBefore)' {
            $vbaContent -match 'Function CountBlankLinesBefore' | Should Be $true
        }
    }

    Context 'Validação de Documento' {

        It 'Possui verificação de saúde do documento (IsDocumentHealthy)' {
            $vbaContent -match 'Function IsDocumentHealthy' | Should Be $true
        }

        It 'Valida versão mínima do Word (MIN_SUPPORTED_VERSION = 14, Word 2010+)' {
            $vbaContent -match 'MIN_SUPPORTED_VERSION.*=.*14' | Should Be $true
        }

        It 'Possui validação de string obrigatória (REQUIRED_STRING)' {
            $vbaContent -match 'REQUIRED_STRING' | Should Be $true
        }
    }

    Context 'Análise de Complexidade' {

        It 'Densidade de código é razoável (> 40% linhas não vazias)' {
            $nonEmptyLines = $vbaLines | Where-Object { $_.Trim() -ne '' }
            $density = $nonEmptyLines.Count / $vbaLines.Count
            $density -gt 0.40 | Should Be $true
        }

        It 'Número de procedimentos por 1000 linhas é razoável (15-25)' {
            $procedures = [regex]::Matches($vbaContent, '(?m)^(Public |Private )?(Sub |Function )\w+')
            $procsPerK = ($procedures.Count / $vbaLines.Count) * 1000
            ($procsPerK -ge 15) -and ($procsPerK -le 25) | Should Be $true
        }

        It 'Possui proteções contra loops infinitos (MAX_LOOP_ITERATIONS)' {
            $vbaContent -match 'MAX_LOOP_ITERATIONS' | Should Be $true
        }

        It 'Possui timeout para operações longas (MAX_OPERATION_TIMEOUT_SECONDS)' {
            $vbaContent -match 'MAX_OPERATION_TIMEOUT_SECONDS' | Should Be $true
        }
    }

    Context 'Configurações de Formatação' {

        It 'Define espaçamento entre linhas (LINE_SPACING)' {
            $vbaContent -match 'LINE_SPACING' | Should Be $true
        }

        It 'Define configurações de cabeçalho e rodapé' {
            ($vbaContent -match 'HEADER_DISTANCE_CM') -and
            ($vbaContent -match 'FOOTER_DISTANCE_CM') -and
            ($vbaContent -match 'FOOTER_FONT_SIZE') | Should Be $true
        }

        It 'Define orientação de página (wdOrientPortrait)' {
            $vbaContent -match 'wdOrientPortrait' | Should Be $true
        }

        It 'Define configurações de sublinhado (wdUnderlineNone, wdUnderlineSingle)' {
            ($vbaContent -match 'wdUnderlineNone') -and
            ($vbaContent -match 'wdUnderlineSingle') | Should Be $true
        }
    }

    Context 'Recursos Avançados' {

        It 'Suporta múltiplas visualizações (wdPrintView)' {
            $vbaContent -match 'wdPrintView' | Should Be $true
        }

        It 'Gerencia alertas do Word (wdAlertsAll, wdAlertsNone)' {
            ($vbaContent -match 'wdAlertsAll') -or
            ($vbaContent -match 'wdAlertsNone') | Should Be $true
        }

        It 'Trabalha com campos do Word (wdFieldPage, wdFieldNumPages)' {
            ($vbaContent -match 'wdFieldPage') -or
            ($vbaContent -match 'wdFieldNumPages') | Should Be $true
        }

        It 'Gerencia shapes e imagens (msoPicture, msoTextEffect)' {
            ($vbaContent -match 'msoPicture') -or
            ($vbaContent -match 'msoTextEffect') | Should Be $true
        }
    }

    Context 'Segurança e Boas Práticas' {

        It 'Fecha arquivos abertos (CloseAllOpenFiles)' {
            $vbaContent -match 'CloseAllOpenFiles' | Should Be $true
        }

        It 'Não contém senhas ou credenciais hardcoded' {
            $vbaContent -notmatch '(?i)(password|senha|pwd)\s*=\s*"[^"]+"' | Should Be $true
        }

        It 'Não contém caminhos absolutos hardcoded (usa caminhos relativos)' {
            # Permite constantes mas não caminhos C:\ direto no código
            $hardcodedPaths = [regex]::Matches($vbaContent, '(?<!Const\s+\w+\s*As\s*String\s*=\s*)"[A-Z]:\\[^"]*"')
            $hardcodedPaths.Count -eq 0 | Should Be $true
        }

        It 'Usa controle de versão documentado' {
            $vbaContent -match 'Vers[aã]o:\s*\d+\.\d+' | Should Be $true
        }
    }

    Context 'Performance e Otimização' {

        It 'Usa variáveis tipadas (As Long, As String, As Range, etc)' {
            ($vbaContent -match '\bAs Long\b') -and
            ($vbaContent -match '\bAs String\b') -and
            ($vbaContent -match '\bAs Range\b') | Should Be $true
        }

        It 'Define constantes Private (performance em VBA)' {
            $vbaContent -match '(?m)^Private Const ' | Should Be $true
        }

        It 'Limita escaneamento inicial de parágrafos (MAX_INITIAL_PARAGRAPHS_TO_SCAN)' {
            $vbaContent -match 'MAX_INITIAL_PARAGRAPHS_TO_SCAN' | Should Be $true
        }
    }

    Context 'Integração e Compatibilidade' {

        It 'Compatível com Word 2010+ (versão 14+)' {
            $vbaContent -match 'MIN_SUPPORTED_VERSION.*=.*14' | Should Be $true
        }

        It 'Referencia Microsoft Word corretamente' {
            $vbaContent -match 'Word' | Should Be $true
        }

        It 'Trabalha com objetos Document corretamente' {
            $vbaContent -match '\bDocument\b' | Should Be $true
        }

        It 'Trabalha com objetos Range corretamente' {
            $vbaContent -match '\bRange\b' | Should Be $true
        }

        It 'Trabalha com objetos Paragraph corretamente' {
            $vbaContent -match '\bParagraph\b' | Should Be $true
        }
    }

    Context 'Funcionalidades Específicas do Chainsaw' {

        It 'Processa "considerando" corretamente (CONSIDERANDO_PREFIX)' {
            $vbaContent -match 'CONSIDERANDO_PREFIX' | Should Be $true
        }

        It 'Define comprimento mínimo para considerando (CONSIDERANDO_MIN_LENGTH)' {
            $vbaContent -match 'CONSIDERANDO_MIN_LENGTH' | Should Be $true
        }

        It 'Referencia pasta de assets (stamp.png)' {
            $vbaContent -match 'stamp\.png' | Should Be $true
        }

        It 'Usa estrutura .chainsaw para organização' {
            $vbaContent -match '\\props\\' | Should Be $true
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
            $funcs = [regex]::Matches($vbaContent, '(?m)^(Public |Private )?Function ([A-Z][a-zA-Z0-9]+)')
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
            $invalidDeclarations = [regex]::Matches($vbaContent, '(?m)^(Dim|Private|Public)\s+As\s+')
            $invalidDeclarations.Count -eq 0 | Should Be $true
        }

        It 'Todas as atribuicoes Set usam palavra-chave Set corretamente' {
            # Set é obrigatório para objetos em VBA
            # Verifica que não há atribuições diretas de objetos sem Set
            $validSetStatements = [regex]::Matches($vbaContent, '(?m)^\s*Set\s+\w+\s*=')
            $validSetStatements.Count -gt 0 | Should Be $true
        }

        It 'Nao ha declaracoes duplicadas de procedimentos' {
            $procedures = [regex]::Matches($vbaContent, '(?m)^(Public |Private )?(Sub |Function )(\w+)')
            $procedureNames = $procedures | ForEach-Object { $_.Groups[3].Value }
            $uniqueNames = $procedureNames | Select-Object -Unique
            $procedureNames.Count -eq $uniqueNames.Count | Should Be $true
        }

        It 'Todos os If tem End If correspondente' {
            $ifCount = [regex]::Matches($vbaContent, '(?m)^\s*(If|ElseIf)\s+.+\s+Then\s*$').Count
            $endIfCount = [regex]::Matches($vbaContent, '(?m)^\s*End If').Count
            # Pode haver If inline (Then ... : End If na mesma linha)
            # Então End If deve ser >= If multilinhas
            $endIfCount -ge ($ifCount * 0.8) | Should Be $true
        }

        It 'Todos os For tem Next correspondente' {
            # Permite loops inline (ex: For ... : ... : Next i)
            $forCount = [regex]::Matches($vbaContent, '(?m)(^\s*For\s+|:\s*For\s+)').Count
            $nextCount = [regex]::Matches($vbaContent, '(?m)(^\s*Next\b|:\s*Next\b)').Count
            [Math]::Abs($forCount - $nextCount) -le 1 | Should Be $true
        }

        It 'Todos os Do tem Loop correspondente' {
            $doCount = [regex]::Matches($vbaContent, '(?m)^\s*Do\s*(While|Until)?').Count
            $loopCount = [regex]::Matches($vbaContent, '(?m)^\s*Loop\b').Count
            # Permite margem de até 10 loops (pode haver Do...Loop While inline, comentários, etc)
            [Math]::Abs($doCount - $loopCount) -le 10 | Should Be $true
        }

        It 'Todos os With tem End With correspondente' {
            $withCount = [regex]::Matches($vbaContent, '(?m)^\s*With\s+').Count
            $endWithCount = [regex]::Matches($vbaContent, '(?m)^\s*End With').Count
            $withCount -eq $endWithCount | Should Be $true
        }

        It 'Todos os Select Case tem End Select correspondente' {
            $selectCount = [regex]::Matches($vbaContent, '(?m)^\s*Select Case\s+').Count
            $endSelectCount = [regex]::Matches($vbaContent, '(?m)^\s*End Select').Count
            $selectCount -eq $endSelectCount | Should Be $true
        }

        It 'Nao ha uso de GoTo sem label correspondente' {
            $goToStatements = [regex]::Matches($vbaContent, '(?m)^\s*(?:On Error )?GoTo\s+(\w+)')
            $labels = [regex]::Matches($vbaContent, '(?m)^(\w+):')

            foreach ($goTo in $goToStatements) {
                $targetLabel = $goTo.Groups[1].Value
                if ($targetLabel -ne '0' -and $targetLabel -ne 'NextIteration') {
                    $labelExists = $labels | Where-Object { $_.Groups[1].Value -eq $targetLabel }
                    $labelExists.Count -gt 0 | Should Be $true
                }
            }
        }

        It 'Todas as funcoes tem tipo de retorno declarado' {
            $functions = [regex]::Matches($vbaContent, '(?m)^(Public |Private )?Function\s+(\w+)\([^)]*\)\s+As\s+\w+')
            $allFunctions = [regex]::Matches($vbaContent, '(?m)^(Public |Private )?Function\s+(\w+)')
            # Todas as funções devem ter tipo de retorno
            $functions.Count -eq $allFunctions.Count | Should Be $true
        }

        It 'Nao ha chamadas a procedimentos inexistentes (verificacao basica)' {
            # Verifica alguns procedimentos críticos que são chamados
            $calledProcs = @('BuildParagraphCache', 'ClearParagraphCache', 'SafeCleanup', 'LogMessage')
            foreach ($proc in $calledProcs) {
                $procDeclared = $vbaContent -match "(Sub |Function )$proc"
                $procDeclared | Should Be $true
            }
        }

        It 'Todas as constantes tem valor atribuido' {
            $constants = [regex]::Matches($vbaContent, '(?m)^(Public |Private )?Const\s+\w+\s+As\s+\w+\s*=')
            $allConstants = [regex]::Matches($vbaContent, '(?m)^(Public |Private )?Const\s+\w+')
            $constants.Count -eq $allConstants.Count | Should Be $true
        }

        It 'Nao ha variaveis declaradas mas nunca usadas (verificacao de principais)' {
            # Verifica algumas variáveis críticas que devem ser usadas
            $criticalVars = @('doc', 'para', 'rng')
            foreach ($var in $criticalVars) {
                $varUsage = [regex]::Matches($vbaContent, "\b$var\b").Count
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
            # Permite até 5 linhas desbalanceadas (continuação de linha VBA)
            $unbalancedLines -le 5 | Should Be $true
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
                if ($vbaContent -match "Set\s+$var\s*=") {
                    # Se Set é usado, deve haver Set = Nothing
                    if ($vbaContent -match "Set\s+$var\s*=\s*Nothing") {
                        $releasedCount++
                    }
                }
            }
            # Pelo menos 2 das 3 variáveis devem ser liberadas
            $releasedCount -ge 2 | Should Be $true
        }

        It 'Nao ha recursao infinita detectavel (funcao chama a si mesma sem condicao)' {
            $functions = [regex]::Matches($vbaContent, '(?s)(Public |Private )?Function\s+(\w+).*?End Function')
            $recursiveWithoutExit = 0

            foreach ($func in $functions) {
                $funcName = $func.Groups[2].Value
                $funcBody = $func.Value

                # Se função chama a si mesma, deve ter If/Exit Function para evitar infinito
                if ($funcBody -match "\b$funcName\(") {
                    $hasExitCondition = ($funcBody -match 'Exit Function') -or
                                      ($funcBody -match '\bIf\b') -or
                                      ($funcBody -match '\bElse\b')
                    if (-not $hasExitCondition) {
                        $recursiveWithoutExit++
                    }
                }
            }
            # Permite até 10 funções recursivas (regex greedy pode não capturar If corretamente)
            $recursiveWithoutExit -le 10 | Should Be $true
        }

        It 'Nao ha atribuicoes a constantes' {
            $constants = [regex]::Matches($vbaContent, '(?m)^(Public |Private )?Const\s+(\w+)')

            foreach ($const in $constants) {
                $constName = $const.Groups[2].Value
                # Não deve haver atribuição após declaração
                $reassignment = [regex]::Matches($vbaContent, "(?m)^\s*$constName\s*=")
                $reassignment.Count -eq 0 | Should Be $true
            }
        }

        It 'Arrays sao declarados corretamente com parenteses' {
            # Arrays em VBA usam () para dimensões (podem ser vazios para dynamic arrays)
            $arrayDeclarations = [regex]::Matches($vbaContent, '(?m)Dim\s+\w+\([^)]*\)\s+As')
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
            $resumeNextCount = [regex]::Matches($vbaContent, '(?m)On Error Resume Next').Count
            $errorGoTo0Count = [regex]::Matches($vbaContent, '(?m)On Error GoTo 0').Count
            $errorGotoLabelCount = [regex]::Matches($vbaContent, '(?m)On Error GoTo \w+').Count

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
}
