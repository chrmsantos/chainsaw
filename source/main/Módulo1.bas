' =============================================================================
' CHAINSAW - Sistema de Padronizacao de Proposituras Legislativas
' =============================================================================
' Versao: 2.9.7
' Data: 2025-12-18
' Licenca: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Compatibilidade: Microsoft Word 2010+
' Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
' =============================================================================

Option Explicit

'================================================================================
' CONSTANTES DO WORD
'================================================================================
Private Const wdNoProtection As Long = -1
Private Const wdTypeDocument As Long = 0
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdAlignParagraphLeft As Long = 0
Private Const wdAlignParagraphCenter As Long = 1
Private Const wdAlignParagraphJustify As Long = 3
Private Const wdLineSpaceSingle As Long = 0
Private Const wdLineSpace1pt5 As Long = 1
Private Const wdLineSpacingMultiple As Long = 5
Private Const wdStatisticPages As Long = 2
Private Const msoTrue As Long = -1
Private Const msoFalse As Long = 0
Private Const msoPicture As Long = 13
Private Const msoTextEffect As Long = 15
Private Const wdCollapseEnd As Long = 0
Private Const wdCollapseStart As Long = 1
Private Const wdFieldPage As Long = 33
Private Const wdFieldNumPages As Long = 26
Private Const wdFieldEmpty As Long = -1
Private Const wdRelativeHorizontalPositionPage As Long = 1
Private Const wdRelativeVerticalPositionPage As Long = 1
Private Const wdWrapTopBottom As Long = 3
Private Const wdAlertsAll As Long = 0
Private Const wdAlertsNone As Long = -1
Private Const wdColorAutomatic As Long = -16777216
Private Const wdOrientPortrait As Long = 0
Private Const wdUnderlineNone As Long = 0
Private Const wdUnderlineSingle As Long = 1
Private Const wdTextureNone As Long = 0
Private Const wdPrintView As Long = 3

'================================================================================
' CONSTANTES DE FORMATAÇÃO
'================================================================================
Private Const STANDARD_FONT As String = "Arial"
Private Const STANDARD_FONT_SIZE As Long = 12
Private Const FOOTER_FONT_SIZE As Long = 9
Private Const LINE_SPACING As Single = 14

Private Const TOP_MARGIN_CM As Double = 4.85
Private Const BOTTOM_MARGIN_CM As Double = 2
Private Const LEFT_MARGIN_CM As Double = 3
Private Const RIGHT_MARGIN_CM As Double = 3
Private Const HEADER_DISTANCE_CM As Double = 0.3
Private Const FOOTER_DISTANCE_CM As Double = 0.9

Private Const HEADER_IMAGE_RELATIVE_PATH As String = "\chainsaw\assets\stamp.png"
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 21
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.7
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.19

'================================================================================
' CONSTANTES DE SISTEMA
'================================================================================
Private Const CHAINSAW_VERSION As String = "2.9.7"
Private Const MIN_SUPPORTED_VERSION As Long = 14
Private Const REQUIRED_STRING As String = "$NUMERO$/$ANO$"
Private Const MAX_BACKUP_FILES As Long = 10
Private Const DEBUG_MODE As Boolean = False

Private Const LOG_LEVEL_INFO As Long = 1
Private Const LOG_LEVEL_WARNING As Long = 2
Private Const LOG_LEVEL_ERROR As Long = 3

Private Const MAX_RETRY_ATTEMPTS As Long = 3
Private Const RETRY_DELAY_MS As Long = 1000
Private Const MAX_LOOP_ITERATIONS As Long = 1000
Private Const MAX_INITIAL_PARAGRAPHS_TO_SCAN As Long = 50
Private Const MAX_OPERATION_TIMEOUT_SECONDS As Long = 300

Private Const CONSIDERANDO_PREFIX As String = "considerando"
Private Const CONSIDERANDO_MIN_LENGTH As Long = 12
Private Const JUSTIFICATIVA_TEXT As String = "justificativa"

'================================================================================
' CONSTANTES DE IDENTIFICACAO DE ELEMENTOS ESTRUTURAIS
'================================================================================
' Criterios para identificacao dos elementos da propositura
Private Const TITULO_MIN_LENGTH As Long = 15              ' Comprimento minimo do titulo
Private Const EMENTA_MIN_LEFT_INDENT As Single = 6        ' Recuo minimo a esquerda da ementa (em pontos)
Private Const PLENARIO_TEXT As String = "plenario"        ' Texto identificador da data (parcial)
Private Const ANEXO_TEXT_SINGULAR As String = "anexo"     ' Texto identificador de anexo (singular)
Private Const ANEXO_TEXT_PLURAL As String = "anexos"      ' Texto identificador de anexo (plural)
Private Const ASSINATURA_PARAGRAPH_COUNT As Long = 3      ' Numero de paragrafos da assinatura
Private Const ASSINATURA_BLANK_LINES_BEFORE As Long = 2   ' Linhas em branco antes da assinatura

'================================================================================
' VARIÁVEIS GLOBAIS
'================================================================================
Private undoGroupEnabled As Boolean
Private loggingEnabled As Boolean
Private logFilePath As String
Private formattingCancelled As Boolean
Private executionStartTime As Date
Private backupFilePath As String
Private errorCount As Long
Private warningCount As Long
Private infoCount As Long
Private logFileHandle As Integer
Private logBufferEnabled As Boolean
Private logBuffer As String
Private lastFlushTime As Date

' Cache de parágrafos para otimização
Private Type paragraphCache
    index As Long
    text As String
    cleanText As String
    hasImages As Boolean
    isSpecial As Boolean
    specialType As String
    needsFormatting As Boolean
    ' Identificadores de elementos estruturais da propositura
    isTitulo As Boolean
    isEmenta As Boolean
    isProposicaoContent As Boolean
    isTituloJustificativa As Boolean
    isJustificativaContent As Boolean
    isData As Boolean
    isAssinatura As Boolean
    isTituloAnexo As Boolean
    isAnexoContent As Boolean
End Type

Private paragraphCache() As paragraphCache
Private cacheSize As Long
Private cacheEnabled As Boolean
Private documentDirty As Boolean  ' Flag para otimizar pipeline de 2 passagens

' Barra de progresso
Private totalSteps As Long
Private currentStep As Long

Private Type ImageInfo
    paraIndex As Long
    ImageIndex As Long
    ImageType As String
    ImageData As Variant
    Position As Long
    WrapType As Long
    Width As Single
    Height As Single
    LeftPosition As Single
    TopPosition As Single
    AnchorRange As Range
End Type

Private savedImages() As ImageInfo
Private imageCount As Long

Private Type ViewSettings
    ViewType As Long
    ShowVerticalRuler As Boolean
    ShowHorizontalRuler As Boolean
    ShowFieldCodes As Boolean
    ShowBookmarks As Boolean
    ShowParagraphMarks As Boolean
    ShowSpaces As Boolean
    ShowTabs As Boolean
    ShowHiddenText As Boolean
    ShowOptionalHyphens As Boolean
    ShowAll As Boolean
    ShowDrawings As Boolean
    ShowObjectAnchors As Boolean
    ShowTextBoundaries As Boolean
    ShowHighlight As Boolean
    DraftFont As Boolean
    WrapToWindow As Boolean
    ShowPicturePlaceHolders As Boolean
    ShowFieldShading As Long
    TableGridlines As Boolean
End Type

Private originalViewSettings As ViewSettings

Private Type ListFormatInfo
    paraIndex As Long
    HasList As Boolean
    ListType As Long
    ListLevelNumber As Long
    ListString As String
End Type

Private savedListFormats() As ListFormatInfo
Private listFormatCount As Long

'================================================================================
' VARIAVEIS DE IDENTIFICACAO DE ELEMENTOS ESTRUTURAIS
'================================================================================
' Índices dos elementos identificados no documento (0 = não encontrado)
Private tituloParaIndex As Long
Private ementaParaIndex As Long
Private proposicaoStartIndex As Long
Private proposicaoEndIndex As Long
Private tituloJustificativaIndex As Long
Private justificativaStartIndex As Long
Private justificativaEndIndex As Long
Private dataParaIndex As Long
Private assinaturaStartIndex As Long
Private assinaturaEndIndex As Long
Private tituloAnexoIndex As Long
Private anexoStartIndex As Long
Private anexoEndIndex As Long

'================================================================================
' PONTO DE ENTRADA PRINCIPAL
'================================================================================
Public Sub PadronizarDocumentoMain()
    On Error GoTo CriticalErrorHandler

    executionStartTime = Now
    formattingCancelled = False
    undoGroupEnabled = False ' Reset inicial

    ' Verificações iniciais ANTES de iniciar UndoRecord
    If Not CheckWordVersion() Then
        Application.StatusBar = "Erro: Word 2010 ou superior necessário"
        LogMessage "Versão do Word " & Application.version & " não suportada. Mínimo: " & CStr(MIN_SUPPORTED_VERSION), LOG_LEVEL_ERROR
        MsgBox "Requer Word 2010 ou superior." & vbCrLf & _
               "Versão atual: " & Application.version, vbCritical, "Versão Incompatível"
        Exit Sub
    End If

    Dim doc As Document
    Set doc = Nothing

    On Error Resume Next
    Set doc = ActiveDocument
    If doc Is Nothing Then
        Application.StatusBar = "Erro: Nenhum documento aberto"
        MsgBox "Nenhum documento está aberto para processamento.", vbCritical, "Erro"
        Exit Sub
    End If
    Err.Clear
    On Error GoTo CriticalErrorHandler
    ' ---------------------------------------------------------------------------

    ' Inicializa sistema de logging ANTES de qualquer LogMessage
    If Not InitializeLogging(doc) Then
        Application.StatusBar = "Aviso: Log desabilitado"
    End If

    ' Inicializa sistema de progresso (18 etapas do pipeline - 2 passagens)
    InitializeProgress 18

    If Not SetAppState(False, "Iniciando...") Then
        LogMessage "Falha ao configurar estado da aplicacao", LOG_LEVEL_WARNING
    End If

    IncrementProgress "Verificando documento"
    If Not PreviousChecking(doc) Then
        GoTo CleanUp
    End If

    If doc.Path = "" Then
        If Not SaveDocumentFirst(doc) Then
            Application.StatusBar = "Cancelado: documento não salvo"
            LogMessage "Operação cancelada - documento não foi salvo", LOG_LEVEL_INFO
            GoTo CleanUp
        End If
    End If

    ' Cria backup do documento antes de qualquer modificação
    IncrementProgress "Criando backup"
    If Not CreateDocumentBackup(doc) Then
        LogMessage "Falha ao criar backup - continuando sem backup", LOG_LEVEL_WARNING
    End If

    ' Backup das configurações de visualização originais
    IncrementProgress "Salvando configurações"
    If Not BackupViewSettings(doc) Then
        LogMessage "Aviso: Falha no backup das configurações de visualização", LOG_LEVEL_WARNING
    End If

    ' Backup de imagens antes das formatações
    IncrementProgress "Protegendo imagens"
    If Not BackupAllImages(doc) Then
        LogMessage "Aviso: Falha no backup de imagens - continuando com proteção básica", LOG_LEVEL_WARNING
    End If

    ' Backup de formatacoes de lista antes das formatacoes
    IncrementProgress "Protegendo listas"
    If Not BackupListFormats(doc) Then
        LogMessage "Aviso: Falha no backup de listas - formatacoes de lista podem ser perdidas", LOG_LEVEL_WARNING
    End If

    ' ==========================================================================
    ' PIPELINE DE FORMATACAO (DUPLA PASSAGEM OTIMIZADA)
    ' ==========================================================================

    LogMessage "=== PIPELINE DE FORMATACAO (2 PASSAGENS) ===", LOG_LEVEL_INFO

    ' Constroi cache de paragrafos (inclui identificacao de estrutura)
    IncrementProgress "Indexando paragrafos"
    BuildParagraphCache doc

    ' Executa formatacao em 2 passagens para garantir estabilidade
    ' Segunda passagem so executa se primeira fez alteracoes (flag dirty)
    Dim pipelinePass As Integer
    documentDirty = True  ' Primeira passagem sempre executa

    For pipelinePass = 1 To 2
        ' Pula segunda passagem se documento nao foi modificado
        If pipelinePass = 2 And Not documentDirty Then
            LogMessage "=== PASSAGEM 2 IGNORADA (sem alteracoes na passagem 1) ===", LOG_LEVEL_INFO
            Exit For
        End If

        documentDirty = False  ' Reset flag antes de cada passagem
        LogMessage "=== PASSAGEM " & pipelinePass & " DE 2 ===", LOG_LEVEL_INFO

        ' Formata documento
        IncrementProgress "Formatando documento (" & pipelinePass & "ª passagem)"
        If Not PreviousFormatting(doc) Then
            GoTo CleanUp
        End If

        ' Restaura imagens após formatações
        IncrementProgress "Restaurando imagens (" & pipelinePass & "ª passagem)"
        If Not RestoreAllImages(doc) Then
            LogMessage "Aviso: Algumas imagens podem ter sido afetadas durante o processamento", LOG_LEVEL_WARNING
        End If
    Next pipelinePass

    ' Remove linhas em branco extras e aplica ajustes finais
    IncrementProgress "Removendo linhas em branco extras"
    RemoverLinhasEmBrancoExtras doc

    ' Restaura formatacoes de lista apos formatacoes
    IncrementProgress "Restaurando listas"
    If Not RestoreListFormats(doc) Then
        LogMessage "Aviso: Algumas formatacoes de lista podem nao ter sido restauradas", LOG_LEVEL_WARNING
    End If

    ' Formata paragrafos iniciados com numero (aplica recuo de lista numerada)
    IncrementProgress "Ajustando numeracao"
    If Not FormatNumberedParagraphsIndent(doc) Then
        LogMessage "Aviso: Falha ao formatar recuos de paragrafos numerados", LOG_LEVEL_WARNING
    End If

    ' Formata paragrafos iniciados com marcador (aplica recuo de lista com marcadores)
    IncrementProgress "Ajustando marcadores"
    If Not FormatBulletedParagraphsIndent(doc) Then
        LogMessage "Aviso: Falha ao formatar recuos de paragrafos com marcadores", LOG_LEVEL_WARNING
    End If

    ' Formata recuos de paragrafos com imagens (zera recuo a esquerda)
    IncrementProgress "Ajustando layout"
    If Not FormatImageParagraphsIndents(doc) Then
        LogMessage "Aviso: Falha ao formatar recuos de imagens", LOG_LEVEL_WARNING
    End If

    ' Centraliza imagem entre 5a e 7a linha apos Plenario
    IncrementProgress "Centralizando elementos"
    If Not CenterImageAfterPlenario(doc) Then
        LogMessage "Aviso: Falha ao centralizar imagem apos Plenario", LOG_LEVEL_WARNING
    End If

    ' Restaura configuracoes de visualizacao originais (exceto zoom)
    IncrementProgress "Restaurando visualizacao"
    If Not RestoreViewSettings(doc) Then
        LogMessage "Aviso: Algumas configuracoes de visualizacao podem nao ter sido restauradas", LOG_LEVEL_WARNING
    End If

    If formattingCancelled Then
        GoTo CleanUp
    End If

    IncrementProgress "Finalizando"
    LogMessage "Documento padronizado com sucesso", LOG_LEVEL_INFO

    ' Calcula tempo de execucao em segundos
    Dim execSeconds As Long
    execSeconds = CLng((Now - executionStartTime) * 86400)

    ' Mostra mensagem final na barra de status
    Application.Sta
    tusBar = "Padronizacao concluida em " & execSeconds & "s, com " & errorCount & " erros e " & warningCount & " avisos! (chainsaw)"

CleanUp:
    ' ---------------------------------------------------------------------------
    ' FIM DO GRUPO DE DESFAZER - SEMPRE fecha o UndoRecord
    ' ---------------------------------------------------------------------------
    On Error Resume Next
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
        LogMessage "UndoRecord finalizado com sucesso", LOG_LEVEL_INFO
    End If
    Err.Clear
    On Error GoTo 0
    ' ---------------------------------------------------------------------------

    ClearParagraphCache ' Limpa cache de paragrafos
    SafeCleanup
    CleanupImageProtection ' Nova funcao para limpar variaveis de protecao de imagens
    CleanupViewSettings    ' Nova funcao para limpar variaveis de configuracoes de visualizacao

    ' Restaura estado da aplicacao preservando a StatusBar (mantem mensagem final)
    If Not SetAppState(True, "", True) Then
        LogMessage "Falha ao restaurar estado da aplicacao", LOG_LEVEL_WARNING
    End If

    SafeFinalizeLogging

    ' Exibe mensagem de conclusão com informações completas
    If Not formattingCancelled Then
        Dim executionTimeText As String
        Dim duration As Double

        ' Calcula duração total
        duration = (Now - executionStartTime) * 86400
        If duration < 60 Then
            executionTimeText = Format(duration, "0.0") & " segundos"
        ElseIf duration < 3600 Then
            executionTimeText = Format(Int(duration / 60), "0") & " minuto(s) e " & Format(duration Mod 60, "00") & " segundo(s)"
        Else
            executionTimeText = Format(Int(duration / 3600), "0") & " hora(s) e " & Format(Int((duration Mod 3600) / 60), "00") & " minuto(s)"
        End If

        ' Monta mensagem com informações de erros/avisos
        Dim statusMsg As String
        If errorCount > 0 Then
            statusMsg = vbCrLf & vbCrLf & "[!] ATENÇÃO: " & errorCount & " erro(s) detectado(s) durante a execução." & vbCrLf & _
                       "   Verifique o log para mais detalhes."
        ElseIf warningCount > 0 Then
            statusMsg = vbCrLf & vbCrLf & "[i] INFORMAÇÃO: " & warningCount & " aviso(s) registrado(s) durante a execução." & vbCrLf & _
                       "   Verifique o log para mais detalhes."
        Else
            statusMsg = vbCrLf & vbCrLf & "[OK] Nenhum erro ou aviso detectado durante a execução."
        End If

        ' Mensagem de sucesso com informações completas
        MsgBox "[OK] Processamento concluído com sucesso em " & executionTimeText & "!" & vbCrLf & vbCrLf & _
               "[DIR] Backup criado em:" & vbCrLf & _
               "   " & IIf(backupFilePath <> "", backupFilePath, GetChainsawBackupsPath()) & vbCrLf & vbCrLf & _
               "[LOG] Log salvo em:" & vbCrLf & _
               "   " & logFilePath & statusMsg, _
               vbInformation, "CHAINSAW - Padronização Concluída"
    End If

    ' Posiciona cursor no início do documento
    On Error Resume Next
    If Not doc Is Nothing Then
        doc.Range(0, 0).Select
    End If
    On Error GoTo 0

    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "ERRO CRÍTICO #" & Err.Number & ": " & Err.Description & _
              " em " & Err.Source & " (Linha: " & Erl & ")"

    LogMessage errDesc, LOG_LEVEL_ERROR
    Application.StatusBar = "Erro - verificar logs"

    ShowUserFriendlyError Err.Number, Err.Description
    EmergencyRecovery

    ' CRÍTICO: Garante fechamento do UndoRecord mesmo em erro
    GoTo CleanUp
End Sub

'================================================================================
' TRATAMENTO AMIGÁVEL DE ERROS
'================================================================================
Private Sub ShowUserFriendlyError(errNum As Long, errDesc As String)
    Dim msg As String

    Select Case errNum
        Case 91 ' Object variable not set
            msg = "Erro: Objeto não inicializado." & vbCrLf & vbCrLf & _
                  "Reinicie o Word."

        Case 5 ' Invalid procedure call
            msg = "Erro de configuração." & vbCrLf & vbCrLf & _
                  "Formato válido: .docx"

        Case 70 ' Permission denied
            msg = "Permissão negada." & vbCrLf & vbCrLf & _
                  "Documento protegido ou somente leitura." & vbCrLf & _
                  "Salve uma cópia."

        Case 53 ' File not found
            msg = "Arquivo não encontrado." & vbCrLf & vbCrLf & _
                  "Verifique se foi salvo."

        Case Else
            msg = "Erro #" & errNum & ":" & vbCrLf & vbCrLf & _
                  errDesc & vbCrLf & vbCrLf & _
                  "Verifique o log."
    End Select

    MsgBox msg, vbCritical, "Chainsaw Proposituras v1.0-beta1"
End Sub

'================================================================================
' RECUPERAÇÃO DE EMERGÊNCIA
'================================================================================
Private Sub EmergencyRecovery()
    On Error Resume Next

    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.StatusBar = False
    Application.EnableCancelKey = 0

    ' Fecha UndoRecord se ainda estiver aberto
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
        LogMessage "UndoRecord fechado durante recuperação de emergência", LOG_LEVEL_WARNING
    End If

    ' Limpa variáveis de proteção de imagens em caso de erro
    CleanupImageProtection

    ' Limpa variáveis de configurações de visualização em caso de erro
    CleanupViewSettings

    ' Limpa cache de parágrafos
    ClearParagraphCache

    LogMessage "Recuperação de emergência executada", LOG_LEVEL_ERROR

    CloseAllOpenFiles
End Sub

'================================================================================
' LIMPEZA SEGURA DE RECURSOS
'================================================================================
Private Sub SafeCleanup()
    On Error Resume Next

    ' Não tenta fechar UndoRecord aqui - já foi fechado em CleanUp

    ReleaseObjects
End Sub

'================================================================================
' LIBERAÇÃO DE OBJETOS
'================================================================================
Private Sub ReleaseObjects()
    On Error Resume Next

    Dim nullObj As Object
    Set nullObj = Nothing

    Dim memoryCounter As Long
    For memoryCounter = 1 To 3
        DoEvents
    Next memoryCounter
End Sub

'================================================================================
' FECHAMENTO DE ARQUIVOS ABERTOS
'================================================================================
Private Sub CloseAllOpenFiles()
    On Error Resume Next

    Dim fileNumber As Integer
    For fileNumber = 1 To 511
        Close #fileNumber
    Next fileNumber
    Err.Clear
End Sub

'================================================================================
' FUNÇÕES DE VALIDAÇÃO E COMPATIBILIDADE
'================================================================================
Private Function ValidateDocument(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ValidateDocument = False

    If doc Is Nothing Then
        LogMessage "Documento é Nothing", LOG_LEVEL_ERROR
        Exit Function
    End If

    If doc.Paragraphs.count = 0 Then
        LogMessage "Documento não tem parágrafos", LOG_LEVEL_WARNING
        Exit Function
    End If

    ValidateDocument = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na validação do documento: " & Err.Description, LOG_LEVEL_ERROR
    ValidateDocument = False
End Function

'================================================================================
' IS DOCUMENT HEALTHY - Validação profunda da integridade do documento
'================================================================================
Private Function IsDocumentHealthy(doc As Document) As Boolean
    On Error Resume Next

    IsDocumentHealthy = False

    ' Verifica acessibilidade básica
    If doc Is Nothing Then Exit Function
    If doc.Range Is Nothing Then Exit Function
    If doc.Paragraphs.count = 0 Then Exit Function

    ' Verifica se documento está corrompido
    Dim testAccess As Long
    testAccess = doc.Range.End
    If Err.Number <> 0 Then Exit Function

    ' Testa acesso a parágrafos
    Dim testPara As Paragraph
    Set testPara = doc.Paragraphs(1)
    If Err.Number <> 0 Then Exit Function

    IsDocumentHealthy = True
End Function

'================================================================================
' IS OPERATION TIMEOUT - Verifica timeout de operações longas
'================================================================================
Private Function IsOperationTimeout(startTime As Date) As Boolean
    IsOperationTimeout = (DateDiff("s", startTime, Now) > MAX_OPERATION_TIMEOUT_SECONDS)
End Function

'================================================================================
' FUNÇÕES AUXILIARES DE LIMPEZA DE TEXTO
'================================================================================
Private Function GetCleanParagraphText(para As Paragraph) As String
    On Error Resume Next

    Dim txt As String
    txt = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

    ' Remove pontuação final com proteção contra loop infinito
    Dim safetyCounter As Long
    safetyCounter = 0
    Do While Len(txt) > 0 And InStr(".,;:", Right(txt, 1)) > 0 And safetyCounter < MAX_LOOP_ITERATIONS
        txt = Left(txt, Len(txt) - 1)
        safetyCounter = safetyCounter + 1
    Loop

        GetCleanParagraphText = RemovePunctuation(Trim(LCase(txt)))
End Function

Private Function RemovePunctuation(text As String) As String
    Dim result As String
    result = text

    ' Remove pontuação final com proteção contra loop infinito
    Dim safetyCounter As Long
    safetyCounter = 0
    Do While Len(result) > 0 And InStr(".,;:", Right(result, 1)) > 0 And safetyCounter < 100
        result = Left(result, Len(result) - 1)
        safetyCounter = safetyCounter + 1
    Loop

    RemovePunctuation = Trim(result)
End Function

'================================================================================
' NORMALIZAÇÃO OTIMIZADA DE TEXTO - Única passagem
'================================================================================
Private Function NormalizarTexto(text As String) As String
    Dim result As String
    Dim loopGuard As Long
    result = text

    ' Remove caracteres de controle em uma unica passagem
    result = Replace(result, vbCr, "")
    result = Replace(result, vbLf, "")
    result = Replace(result, vbTab, " ")

    ' Remove espacos multiplos com protecao contra loop infinito
    loopGuard = 0
    Do While InStr(result, "  ") > 0 And loopGuard < 500
        result = Replace(result, "  ", " ")
        loopGuard = loopGuard + 1
    Loop

    NormalizarTexto = Trim(LCase(result))
End Function

'================================================================================
' DETECÇÃO DE TIPO DE PARÁGRAFO ESPECIAL
'================================================================================
Private Function DetectSpecialParagraph(cleanText As String, ByRef specialType As String) As Boolean
    specialType = ""

    ' Remove pontuação final para análise
    Dim textForAnalysis As String
    textForAnalysis = cleanText

    Dim safetyCounter As Long
    safetyCounter = 0
    Do While Len(textForAnalysis) > 0 And InStr(".,;:", Right(textForAnalysis, 1)) > 0 And safetyCounter < 50
        textForAnalysis = Left(textForAnalysis, Len(textForAnalysis) - 1)
        safetyCounter = safetyCounter + 1
    Loop
    textForAnalysis = Trim(textForAnalysis)

    ' Verifica tipos especiais
    If Left(textForAnalysis, CONSIDERANDO_MIN_LENGTH) = CONSIDERANDO_PREFIX Then
        specialType = "considerando"
        DetectSpecialParagraph = True
    ElseIf textForAnalysis = JUSTIFICATIVA_TEXT Then
        specialType = "justificativa"
        DetectSpecialParagraph = True
    ElseIf textForAnalysis = "vereador" Or textForAnalysis = "vereadora" Then
        specialType = "vereador"
        DetectSpecialParagraph = True
    ElseIf Left(textForAnalysis, 17) = "diante do exposto" Then
        specialType = "dianteexposto"
        DetectSpecialParagraph = True
    ElseIf textForAnalysis = "requeiro" Then
        specialType = "requeiro"
        DetectSpecialParagraph = True
    ElseIf textForAnalysis = "anexo" Or textForAnalysis = "anexos" Then
        specialType = "anexo"
        DetectSpecialParagraph = True
    Else
        DetectSpecialParagraph = False
    End If
End Function

'================================================================================
' IDENTIFICAÇÃO DE ELEMENTOS ESTRUTURAIS DA PROPOSITURA
'================================================================================

'--------------------------------------------------------------------------------
' IsTituloElement - Identifica se o parágrafo é o título da propositura
'--------------------------------------------------------------------------------
' Critérios:
' - 1ª linha contendo texto
' - Negrito, sublinhado, caixa alta
' - Recuo = 0
' - Mais de 15 caracteres
' - Termina com "$NUMERO$/$ANO$"
'--------------------------------------------------------------------------------
Private Function IsTituloElement(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler

    IsTituloElement = False

    ' Validação de segurança
    If para Is Nothing Then Exit Function
    If para.Range Is Nothing Then Exit Function

    ' Obtém texto limpo
    Dim paraText As String
    paraText = Trim(para.Range.text)
    If Len(paraText) < TITULO_MIN_LENGTH Then Exit Function

    ' Verifica se termina com a string requerida
    Dim cleanText As String
    cleanText = Replace(Replace(paraText, vbCr, ""), vbLf, "")
    If Not (Right(Trim(cleanText), Len(REQUIRED_STRING)) = REQUIRED_STRING) Then Exit Function

    ' Verifica formatação do parágrafo
    With para.Format
        If .leftIndent <> 0 Then Exit Function
        If .alignment <> wdAlignParagraphLeft Then Exit Function
    End With

    ' Verifica formatação do texto (negrito, sublinhado, caixa alta)
    With para.Range.Font
        If .Bold <> msoTrue Then Exit Function
        If .Underline = wdUnderlineNone Then Exit Function
        If .AllCaps <> msoTrue Then Exit Function
    End With

    IsTituloElement = True
    Exit Function

ErrorHandler:
    IsTituloElement = False
End Function

'--------------------------------------------------------------------------------
' IsEmentaElement - Identifica se o parágrafo é a ementa
'--------------------------------------------------------------------------------
' Critérios:
' - Parágrafo único imediatamente abaixo do título
' - Recuo à esquerda > 6 pontos
' - Contém texto
'--------------------------------------------------------------------------------
Private Function IsEmentaElement(para As Paragraph, prevParaIsTitulo As Boolean) As Boolean
    On Error GoTo ErrorHandler

    IsEmentaElement = False

    ' Validação de segurança
    If para Is Nothing Then Exit Function
    If Not prevParaIsTitulo Then Exit Function

    ' Verifica se contém texto
    Dim paraText As String
    paraText = Trim(para.Range.text)
    If Len(paraText) = 0 Then Exit Function

    ' Verifica recuo à esquerda
    If para.Format.leftIndent <= EMENTA_MIN_LEFT_INDENT Then Exit Function

    IsEmentaElement = True
    Exit Function

ErrorHandler:
    IsEmentaElement = False
End Function

'--------------------------------------------------------------------------------
' IsJustificativaTitleElement - Identifica o título "Justificativa"
'--------------------------------------------------------------------------------
Private Function IsJustificativaTitleElement(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler

    IsJustificativaTitleElement = False

    ' Validação de segurança
    If para Is Nothing Then Exit Function

    ' Verifica se o texto é "Justificativa"
    Dim cleanText As String
    cleanText = GetCleanParagraphText(para)
    If cleanText <> JUSTIFICATIVA_TEXT Then Exit Function

    IsJustificativaTitleElement = True
    Exit Function

ErrorHandler:
    IsJustificativaTitleElement = False
End Function

'--------------------------------------------------------------------------------
' IsDataElement - Identifica o parágrafo de data (Plenário)
'--------------------------------------------------------------------------------
' Critérios:
' - Contém "Plenário "Dr. Tancredo Neves", $DATAATUALEXTENSO$."
'--------------------------------------------------------------------------------
Private Function IsDataElement(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler

    IsDataElement = False

    ' Validação de segurança
    If para Is Nothing Then Exit Function

    ' Verifica se contém o texto do plenário
    Dim paraTextLower As String
    paraTextLower = LCase(Trim(para.Range.text))

    ' Busca por "plenário" e elementos relacionados
    If InStr(paraTextLower, "plenário") > 0 And _
       InStr(paraTextLower, "tancredo neves") > 0 Then
        IsDataElement = True
    End If

    Exit Function

ErrorHandler:
    IsDataElement = False
End Function

'--------------------------------------------------------------------------------
' IsTituloAnexoElement - Identifica o título "Anexo" ou "Anexos"
'--------------------------------------------------------------------------------
' Critérios:
' - Parágrafo unicamente com palavra "Anexo" ou "Anexos"
' - Negrito, recuo 0, alinhado à esquerda
'--------------------------------------------------------------------------------
Private Function IsTituloAnexoElement(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler

    IsTituloAnexoElement = False

    ' Validação de segurança
    If para Is Nothing Then Exit Function

    ' Verifica texto
    Dim cleanText As String
    cleanText = GetCleanParagraphText(para)
    If cleanText <> ANEXO_TEXT_SINGULAR And cleanText <> ANEXO_TEXT_PLURAL Then Exit Function

    ' Verifica formatação
    With para.Format
        If .leftIndent <> 0 Then Exit Function
        If .alignment <> wdAlignParagraphLeft Then Exit Function
    End With

    ' Verifica negrito
    If para.Range.Font.Bold <> msoTrue Then Exit Function

    IsTituloAnexoElement = True
    Exit Function

ErrorHandler:
    IsTituloAnexoElement = False
End Function

'--------------------------------------------------------------------------------
' CountBlankLinesBefore - Conta linhas em branco antes de um parágrafo
'--------------------------------------------------------------------------------
Private Function CountBlankLinesBefore(doc As Document, paraIndex As Long) As Long
    On Error GoTo ErrorHandler

    CountBlankLinesBefore = 0

    If paraIndex <= 1 Then Exit Function
    If paraIndex > doc.Paragraphs.count Then Exit Function

    Dim i As Long
    Dim blankCount As Long
    blankCount = 0

    ' Volta até encontrar parágrafo não-vazio ou até 5 linhas
    For i = paraIndex - 1 To 1 Step -1
        If i > doc.Paragraphs.count Then Exit For

        Dim paraText As String
        paraText = Trim(doc.Paragraphs(i).Range.text)

        If Len(paraText) = 0 Then
            blankCount = blankCount + 1
        Else
            Exit For
        End If

        ' Limita a 5 linhas para evitar loops longos
        If blankCount >= 5 Then Exit For
    Next i

    CountBlankLinesBefore = blankCount
    Exit Function

ErrorHandler:
    CountBlankLinesBefore = 0
End Function

'--------------------------------------------------------------------------------
' IsAssinaturaStart - Identifica o início da assinatura
'--------------------------------------------------------------------------------
' Critérios:
' - 3 parágrafos textuais
' - 2 linhas em branco antes
' - Centralizados
' - Sem linhas em branco entre si
' - Pode ter imagens logo abaixo (sem linhas em branco)
'--------------------------------------------------------------------------------
Private Function IsAssinaturaStart(doc As Document, paraIndex As Long) As Boolean
    On Error GoTo ErrorHandler

    IsAssinaturaStart = False

    ' Validação de segurança
    If paraIndex <= 0 Or paraIndex > doc.Paragraphs.count Then Exit Function

    ' Verifica se há linhas em branco antes (pelo menos 2)
    If CountBlankLinesBefore(doc, paraIndex) < ASSINATURA_BLANK_LINES_BEFORE Then Exit Function

    ' Verifica se há 3 parágrafos consecutivos centralizados com texto
    Dim i As Long
    Dim consecutiveCount As Long
    consecutiveCount = 0

    For i = paraIndex To doc.Paragraphs.count
        If i > doc.Paragraphs.count Then Exit For

        Dim para As Paragraph
        Set para = doc.Paragraphs(i)

        Dim paraText As String
        paraText = Trim(para.Range.text)

        ' Se encontrou parágrafo vazio, para a contagem
        If Len(paraText) = 0 Then
            Exit For
        End If

        ' Verifica se está centralizado
        If para.Format.alignment = wdAlignParagraphCenter Then
            consecutiveCount = consecutiveCount + 1
        Else
            Exit For
        End If

        ' Se já encontrou 3, é uma assinatura
        If consecutiveCount >= ASSINATURA_PARAGRAPH_COUNT Then
            IsAssinaturaStart = True
            Exit Function
        End If

        ' Limite de segurança
        If i - paraIndex > 10 Then Exit For
    Next i

    Exit Function

ErrorHandler:
    IsAssinaturaStart = False
End Function

'--------------------------------------------------------------------------------
' IdentifyDocumentStructure - Identifica todos os elementos estruturais
'--------------------------------------------------------------------------------
' Esta função percorre o documento e identifica:
' - Título, Ementa, Proposição, Justificativa, Data, Assinatura, Anexo
'--------------------------------------------------------------------------------
Private Sub IdentifyDocumentStructure(doc As Document)
    On Error GoTo ErrorHandler

    LogMessage "Identificando estrutura do documento...", LOG_LEVEL_INFO

    ' Reseta todos os índices
    tituloParaIndex = 0
    ementaParaIndex = 0
    proposicaoStartIndex = 0
    proposicaoEndIndex = 0
    tituloJustificativaIndex = 0
    justificativaStartIndex = 0
    justificativaEndIndex = 0
    dataParaIndex = 0
    assinaturaStartIndex = 0
    assinaturaEndIndex = 0
    tituloAnexoIndex = 0
    anexoStartIndex = 0
    anexoEndIndex = 0

    Dim i As Long
    Dim para As Paragraph
    Dim foundTitulo As Boolean
    Dim foundJustificativa As Boolean
    Dim foundData As Boolean

    foundTitulo = False
    foundJustificativa = False
    foundData = False

    ' Percorre todos os parágrafos
    For i = 1 To cacheSize
        ' Proteção contra mudanças no documento durante execução
        If i > doc.Paragraphs.count Then Exit For

        Set para = doc.Paragraphs(i)

        ' Atualiza cache com identificação
        With paragraphCache(i)
            ' Reseta flags
            .isTitulo = False
            .isEmenta = False
            .isProposicaoContent = False
            .isTituloJustificativa = False
            .isJustificativaContent = False
            .isData = False
            .isAssinatura = False
            .isTituloAnexo = False
            .isAnexoContent = False

            ' 1. Identifica TÍTULO (primeira ocorrência)
            If Not foundTitulo And IsTituloElement(para) Then
                .isTitulo = True
                tituloParaIndex = i
                foundTitulo = True
                LogMessage "Título identificado no parágrafo " & i, LOG_LEVEL_INFO

            ' 2. Identifica EMENTA (logo após o título)
            ElseIf foundTitulo And ementaParaIndex = 0 Then
                If IsEmentaElement(para, True) Then
                    .isEmenta = True
                    ementaParaIndex = i
                    proposicaoStartIndex = i + 1 ' Proposição começa logo após a ementa
                    LogMessage "Ementa identificada no parágrafo " & i, LOG_LEVEL_INFO
                End If

            ' 3. Identifica TÍTULO DA JUSTIFICATIVA
            ElseIf Not foundJustificativa And IsJustificativaTitleElement(para) Then
                .isTituloJustificativa = True
                tituloJustificativaIndex = i
                foundJustificativa = True
                ' Proposição termina antes da Justificativa
                If proposicaoStartIndex > 0 Then
                    proposicaoEndIndex = i - 1
                End If
                justificativaStartIndex = i + 1 ' Justificativa começa logo após o título
                LogMessage "Título da Justificativa identificado no parágrafo " & i, LOG_LEVEL_INFO

            ' 4. Identifica DATA (Plenário)
            ElseIf Not foundData And IsDataElement(para) Then
                .isData = True
                dataParaIndex = i
                foundData = True
                ' Justificativa termina antes da Data
                If justificativaStartIndex > 0 Then
                    justificativaEndIndex = i - 1
                End If
                LogMessage "Data (Plenário) identificada no parágrafo " & i, LOG_LEVEL_INFO

            ' 5. Identifica ASSINATURA (após a data, com 2 linhas em branco)
            ElseIf foundData And assinaturaStartIndex = 0 And IsAssinaturaStart(doc, i) Then
                .isAssinatura = True
                assinaturaStartIndex = i
                ' Conta os 3 parágrafos + imagens (se houver)
                Dim j As Long
                Dim assinaturaCount As Long
                assinaturaCount = 0
                For j = i To doc.Paragraphs.count
                    If j > doc.Paragraphs.count Then Exit For
                    Dim tempPara As Paragraph
                    Set tempPara = doc.Paragraphs(j)
                    Dim tempText As String
                    tempText = Trim(tempPara.Range.text)

                    ' Para em linha vazia
                    If Len(tempText) = 0 Then Exit For

                    ' Marca como assinatura
                    paragraphCache(j).isAssinatura = True
                    assinaturaCount = assinaturaCount + 1
                    assinaturaEndIndex = j

                    ' Se já contou 3 parágrafos, verifica se há imagens nos próximos
                    If assinaturaCount >= ASSINATURA_PARAGRAPH_COUNT Then
                        ' Verifica se próximo parágrafo tem imagem (sem linha vazia)
                        If j + 1 <= doc.Paragraphs.count Then
                            Set tempPara = doc.Paragraphs(j + 1)
                            If HasVisualContent(tempPara) Then
                                ' Inclui imagem na assinatura
                                paragraphCache(j + 1).isAssinatura = True
                                assinaturaEndIndex = j + 1
                            End If
                        End If
                        Exit For
                    End If

                    ' Limite de segurança
                    If assinaturaCount > 10 Then Exit For
                Next j
                LogMessage "Assinatura identificada nos parágrafos " & assinaturaStartIndex & " a " & assinaturaEndIndex, LOG_LEVEL_INFO

            ' 6. Identifica TÍTULO DO ANEXO
            ElseIf tituloAnexoIndex = 0 And IsTituloAnexoElement(para) Then
                .isTituloAnexo = True
                tituloAnexoIndex = i
                anexoStartIndex = i + 1 ' Anexo começa logo após o título
                LogMessage "Título do Anexo identificado no parágrafo " & i, LOG_LEVEL_INFO
            End If

            ' Marca conteúdo da PROPOSIÇÃO
            If proposicaoStartIndex > 0 And proposicaoEndIndex > 0 Then
                If i >= proposicaoStartIndex And i <= proposicaoEndIndex Then
                    .isProposicaoContent = True
                End If
            End If

            ' Marca conteúdo da JUSTIFICATIVA
            If justificativaStartIndex > 0 And justificativaEndIndex > 0 Then
                If i >= justificativaStartIndex And i <= justificativaEndIndex Then
                    .isJustificativaContent = True
                End If
            End If

            ' Marca conteúdo do ANEXO
            If anexoStartIndex > 0 And i >= anexoStartIndex Then
                .isAnexoContent = True
                anexoEndIndex = i
            End If
        End With

        ' Atualiza progresso a cada 50 parágrafos
        If i Mod 50 = 0 Then
            DoEvents
        End If
    Next i

    ' Se não encontrou fim da proposição, define até antes da justificativa ou data
    If proposicaoStartIndex > 0 And proposicaoEndIndex = 0 Then
        If tituloJustificativaIndex > 0 Then
            proposicaoEndIndex = tituloJustificativaIndex - 1
        ElseIf dataParaIndex > 0 Then
            proposicaoEndIndex = dataParaIndex - 1
        Else
            proposicaoEndIndex = cacheSize
        End If
    End If

    ' Se não encontrou fim da justificativa, define até antes da data
    If justificativaStartIndex > 0 And justificativaEndIndex = 0 Then
        If dataParaIndex > 0 Then
            justificativaEndIndex = dataParaIndex - 1
        Else
            justificativaEndIndex = cacheSize
        End If
    End If

    ' Relatório de identificação
    LogMessage "=== ESTRUTURA DO DOCUMENTO IDENTIFICADA ===", LOG_LEVEL_INFO
    LogMessage "Título: parágrafo " & tituloParaIndex, LOG_LEVEL_INFO
    LogMessage "Ementa: parágrafo " & ementaParaIndex, LOG_LEVEL_INFO
    LogMessage "Proposição: parágrafos " & proposicaoStartIndex & " a " & proposicaoEndIndex, LOG_LEVEL_INFO
    LogMessage "Título Justificativa: parágrafo " & tituloJustificativaIndex, LOG_LEVEL_INFO
    LogMessage "Justificativa: parágrafos " & justificativaStartIndex & " a " & justificativaEndIndex, LOG_LEVEL_INFO
    LogMessage "Data: parágrafo " & dataParaIndex, LOG_LEVEL_INFO
    LogMessage "Assinatura: parágrafos " & assinaturaStartIndex & " a " & assinaturaEndIndex, LOG_LEVEL_INFO
    LogMessage "Título Anexo: parágrafo " & tituloAnexoIndex, LOG_LEVEL_INFO
    LogMessage "Anexo: parágrafos " & anexoStartIndex & " a " & anexoEndIndex, LOG_LEVEL_INFO
    LogMessage "==========================================", LOG_LEVEL_INFO

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao identificar estrutura do documento: " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' CONSTRUCAO DO CACHE DE PARAGRAFOS - Otimizacao principal
'================================================================================
Private Sub BuildParagraphCache(doc As Document)
    On Error GoTo ErrorHandler

    Dim startTime As Double
    startTime = Timer

    LogMessage "Iniciando construcao do cache de paragrafos...", LOG_LEVEL_INFO

    cacheSize = doc.Paragraphs.count
    ReDim paragraphCache(1 To cacheSize)

    Dim i As Long
    Dim para As Paragraph
    Dim rawText As String

    For i = 1 To cacheSize
        ' DoEvents a cada 20 paragrafos para manter responsividade
        If i Mod 20 = 0 Then DoEvents

        Set para = doc.Paragraphs(i)

        ' Captura o texto bruto uma unica vez
        On Error Resume Next
        rawText = para.Range.text
        On Error GoTo ErrorHandler

        With paragraphCache(i)
            .index = i
            .text = rawText
            .cleanText = NormalizarTexto(rawText)
            .hasImages = HasVisualContent(para)
            .isSpecial = DetectSpecialParagraph(.cleanText, .specialType)
            .needsFormatting = (Len(.cleanText) > 0) And (Not .hasImages)
        End With

        ' Atualiza progresso a cada 100 paragrafos
        If i Mod 100 = 0 Then
            UpdateProgress "Indexando: " & i & "/" & cacheSize, 5 + (i * 5 \ cacheSize)
        End If
    Next i

    cacheEnabled = True

    Dim elapsed As Single
    elapsed = Timer - startTime

    LogMessage "Cache construido: " & cacheSize & " paragrafos em " & Format(elapsed, "0.00") & "s", LOG_LEVEL_INFO

    ' Identifica a estrutura do documento apos construir o cache
    IdentifyDocumentStructure doc

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao construir cache: " & Err.Description, LOG_LEVEL_ERROR
    cacheEnabled = False
End Sub

'================================================================================
' LIMPEZA DO CACHE
'================================================================================
Private Sub ClearParagraphCache()
    On Error Resume Next
    Erase paragraphCache
    cacheSize = 0
    cacheEnabled = False

    ' Limpa também os índices de identificação
    tituloParaIndex = 0
    ementaParaIndex = 0
    proposicaoStartIndex = 0
    proposicaoEndIndex = 0
    tituloJustificativaIndex = 0
    justificativaStartIndex = 0
    justificativaEndIndex = 0
    dataParaIndex = 0
    assinaturaStartIndex = 0
    assinaturaEndIndex = 0
    tituloAnexoIndex = 0
    anexoStartIndex = 0
    anexoEndIndex = 0
End Sub

'================================================================================
' FUNÇÕES PÚBLICAS DE ACESSO AOS ELEMENTOS ESTRUTURAIS
'================================================================================

'--------------------------------------------------------------------------------
' GetTituloRange - Retorna o Range do título
'--------------------------------------------------------------------------------
Public Function GetTituloRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetTituloRange = Nothing

    If tituloParaIndex <= 0 Or tituloParaIndex > doc.Paragraphs.count Then Exit Function
    Set GetTituloRange = doc.Paragraphs(tituloParaIndex).Range
    Exit Function

ErrorHandler:
    Set GetTituloRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetEmentaRange - Retorna o Range da ementa
'--------------------------------------------------------------------------------
Public Function GetEmentaRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetEmentaRange = Nothing

    If ementaParaIndex <= 0 Or ementaParaIndex > doc.Paragraphs.count Then Exit Function
    Set GetEmentaRange = doc.Paragraphs(ementaParaIndex).Range
    Exit Function

ErrorHandler:
    Set GetEmentaRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetProposicaoRange - Retorna o Range da proposição (conjunto de parágrafos)
'--------------------------------------------------------------------------------
Public Function GetProposicaoRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetProposicaoRange = Nothing

    If proposicaoStartIndex <= 0 Or proposicaoEndIndex <= 0 Then Exit Function
    If proposicaoStartIndex > doc.Paragraphs.count Then Exit Function
    If proposicaoEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long
    Dim endPos As Long

    startPos = doc.Paragraphs(proposicaoStartIndex).Range.Start
    endPos = doc.Paragraphs(proposicaoEndIndex).Range.End

    Set GetProposicaoRange = doc.Range(startPos, endPos)
    Exit Function

ErrorHandler:
    Set GetProposicaoRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetTituloJustificativaRange - Retorna o Range do título "Justificativa"
'--------------------------------------------------------------------------------
Public Function GetTituloJustificativaRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetTituloJustificativaRange = Nothing

    If tituloJustificativaIndex <= 0 Or tituloJustificativaIndex > doc.Paragraphs.count Then Exit Function
    Set GetTituloJustificativaRange = doc.Paragraphs(tituloJustificativaIndex).Range
    Exit Function

ErrorHandler:
    Set GetTituloJustificativaRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetJustificativaRange - Retorna o Range da justificativa (conjunto de parágrafos)
'--------------------------------------------------------------------------------
Public Function GetJustificativaRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetJustificativaRange = Nothing

    If justificativaStartIndex <= 0 Or justificativaEndIndex <= 0 Then Exit Function
    If justificativaStartIndex > doc.Paragraphs.count Then Exit Function
    If justificativaEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long
    Dim endPos As Long

    startPos = doc.Paragraphs(justificativaStartIndex).Range.Start
    endPos = doc.Paragraphs(justificativaEndIndex).Range.End

    Set GetJustificativaRange = doc.Range(startPos, endPos)
    Exit Function

ErrorHandler:
    Set GetJustificativaRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetDataRange - Retorna o Range da data (Plenário)
'--------------------------------------------------------------------------------
Public Function GetDataRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetDataRange = Nothing

    If dataParaIndex <= 0 Or dataParaIndex > doc.Paragraphs.count Then Exit Function
    Set GetDataRange = doc.Paragraphs(dataParaIndex).Range
    Exit Function

ErrorHandler:
    Set GetDataRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetAssinaturaRange - Retorna o Range da assinatura (3 parágrafos + imagens)
'--------------------------------------------------------------------------------
Public Function GetAssinaturaRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetAssinaturaRange = Nothing

    If assinaturaStartIndex <= 0 Or assinaturaEndIndex <= 0 Then Exit Function
    If assinaturaStartIndex > doc.Paragraphs.count Then Exit Function
    If assinaturaEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long
    Dim endPos As Long

    startPos = doc.Paragraphs(assinaturaStartIndex).Range.Start
    endPos = doc.Paragraphs(assinaturaEndIndex).Range.End

    Set GetAssinaturaRange = doc.Range(startPos, endPos)
    Exit Function

ErrorHandler:
    Set GetAssinaturaRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetTituloAnexoRange - Retorna o Range do título "Anexo" ou "Anexos"
'--------------------------------------------------------------------------------
Public Function GetTituloAnexoRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetTituloAnexoRange = Nothing

    If tituloAnexoIndex <= 0 Or tituloAnexoIndex > doc.Paragraphs.count Then Exit Function
    Set GetTituloAnexoRange = doc.Paragraphs(tituloAnexoIndex).Range
    Exit Function

ErrorHandler:
    Set GetTituloAnexoRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetAnexoRange - Retorna o Range do anexo (todo conteúdo abaixo do título)
'--------------------------------------------------------------------------------
Public Function GetAnexoRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetAnexoRange = Nothing

    If anexoStartIndex <= 0 Or anexoEndIndex <= 0 Then Exit Function
    If anexoStartIndex > doc.Paragraphs.count Then Exit Function
    If anexoEndIndex > doc.Paragraphs.count Then Exit Function

    Dim startPos As Long
    Dim endPos As Long

    startPos = doc.Paragraphs(anexoStartIndex).Range.Start
    endPos = doc.Paragraphs(anexoEndIndex).Range.End

    Set GetAnexoRange = doc.Range(startPos, endPos)
    Exit Function

ErrorHandler:
    Set GetAnexoRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetProposituraRange - Retorna o Range de toda a propositura (documento completo)
'--------------------------------------------------------------------------------
Public Function GetProposituraRange(doc As Document) As Range
    On Error GoTo ErrorHandler

    Set GetProposituraRange = Nothing

    If doc Is Nothing Then Exit Function
    Set GetProposituraRange = doc.Range
    Exit Function

ErrorHandler:
    Set GetProposituraRange = Nothing
End Function

'--------------------------------------------------------------------------------
' GetElementInfo - Retorna informações sobre todos os elementos identificados
' REFATORADO: Usa funções identificadoras ao invés de acesso direto às variáveis
'--------------------------------------------------------------------------------
Public Function GetElementInfo(doc As Document) As String
    On Error Resume Next

    Dim info As String
    Dim rng As Range

    info = "=== INFORMAÇÕES DOS ELEMENTOS ESTRUTURAIS ===" & vbCrLf

    ' Título - usa GetTituloRange
    Set rng = GetTituloRange(doc)
    If Not rng Is Nothing Then
        info = info & "Título: Parágrafo " & tituloParaIndex & vbCrLf
    Else
        info = info & "Título: Não identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Ementa - usa GetEmentaRange
    Set rng = GetEmentaRange(doc)
    If Not rng Is Nothing Then
        info = info & "Ementa: Parágrafo " & ementaParaIndex & vbCrLf
    Else
        info = info & "Ementa: Não identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Proposição - usa GetProposicaoRange
    Set rng = GetProposicaoRange(doc)
    If Not rng Is Nothing Then
        info = info & "Proposição: Parágrafos " & proposicaoStartIndex & " a " & proposicaoEndIndex & _
                      " (" & (proposicaoEndIndex - proposicaoStartIndex + 1) & " parágrafos)" & vbCrLf
    Else
        info = info & "Proposição: Não identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Título Justificativa - ainda usa variável direta (não tem função Get específica)
    If tituloJustificativaIndex > 0 Then
        info = info & "Título Justificativa: Parágrafo " & tituloJustificativaIndex & vbCrLf
    Else
        info = info & "Título Justificativa: Não identificado" & vbCrLf
    End If

    ' Justificativa - usa GetJustificativaRange
    Set rng = GetJustificativaRange(doc)
    If Not rng Is Nothing Then
        info = info & "Justificativa: Parágrafos " & justificativaStartIndex & " a " & justificativaEndIndex & _
                      " (" & (justificativaEndIndex - justificativaStartIndex + 1) & " parágrafos)" & vbCrLf
    Else
        info = info & "Justificativa: Não identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Data - usa GetDataRange
    Set rng = GetDataRange(doc)
    If Not rng Is Nothing Then
        info = info & "Data (Plenário): Parágrafo " & dataParaIndex & vbCrLf
    Else
        info = info & "Data (Plenário): Não identificado" & vbCrLf
    End If
    Set rng = Nothing

    ' Assinatura - usa GetAssinaturaRange
    Set rng = GetAssinaturaRange(doc)
    If Not rng Is Nothing Then
        info = info & "Assinatura: Parágrafos " & assinaturaStartIndex & " a " & assinaturaEndIndex & _
                      " (" & (assinaturaEndIndex - assinaturaStartIndex + 1) & " parágrafos)" & vbCrLf
    Else
        info = info & "Assinatura: Não identificado" & vbCrLf
    End If
    Set rng = Nothing

    If tituloAnexoIndex > 0 Then
        info = info & "Título Anexo: Parágrafo " & tituloAnexoIndex & vbCrLf
        If anexoStartIndex > 0 And anexoEndIndex > 0 Then
            info = info & "Anexo: Parágrafos " & anexoStartIndex & " a " & anexoEndIndex & _
                          " (" & (anexoEndIndex - anexoStartIndex + 1) & " parágrafos)" & vbCrLf
        End If
    Else
        info = info & "Anexo: Não presente" & vbCrLf
    End If

    info = info & "============================================="

    GetElementInfo = info
End Function

'================================================================================
' ATUALIZACAO DA BARRA DE PROGRESSO
'================================================================================
Private Sub UpdateProgress(message As String, percentComplete As Long)
    ' Mostra apenas "Padronizando..." durante a execucao
    Application.StatusBar = "Padronizando..."

    ' Forca atualizacao da tela
    DoEvents
End Sub

'================================================================================
' CÁLCULO DE PROGRESSO BASEADO EM ETAPAS
'================================================================================
Private Sub InitializeProgress(steps As Long)
    totalSteps = steps
    currentStep = 0
End Sub

Private Sub IncrementProgress(message As String)
    currentStep = currentStep + 1
    Dim percent As Long
    If totalSteps > 0 Then
        percent = CLng((currentStep * 100) / totalSteps)
    Else
        percent = 0
    End If
    UpdateProgress message, percent
End Sub

'================================================================================
' VERIFICAÇÃO DE VERSÃO DO WORD
'================================================================================
Private Function CheckWordVersion() As Boolean
    On Error GoTo ErrorHandler

    Dim version As Double
    ' Uso de CDbl para garantir conversão correta em todas as versões
    version = CDbl(Application.version)

    If version < MIN_SUPPORTED_VERSION Then
        CheckWordVersion = False
        LogMessage "Versão detectada: " & CStr(version) & " - Mínima suportada: " & CStr(MIN_SUPPORTED_VERSION), LOG_LEVEL_ERROR
    Else
        CheckWordVersion = True
        LogMessage "Versão do Word compatível: " & CStr(version), LOG_LEVEL_INFO
    End If

    Exit Function

ErrorHandler:
    ' Se não conseguir detectar a versão, assume incompatibilidade por segurança
    CheckWordVersion = False
    LogMessage "Erro ao detectar versão do Word: " & Err.Description, LOG_LEVEL_ERROR
End Function

'================================================================================
' ACESSO SEGURO A PROPRIEDADES
'================================================================================
Private Function SafeGetCharacterCount(targetRange As Range) As Long
    On Error GoTo FallbackMethod

    ' Método preferido - mais rápido
    SafeGetCharacterCount = targetRange.Characters.count
    Exit Function

FallbackMethod:
    On Error GoTo ErrorHandler
    ' Método alternativo para versões com problemas de .Characters.Count
    SafeGetCharacterCount = Len(targetRange.text)
    Exit Function

ErrorHandler:
    ' Último recurso - valor padrão seguro
    SafeGetCharacterCount = 0
    LogMessage "Erro ao obter contagem de caracteres: " & Err.Description, LOG_LEVEL_WARNING
End Function

Private Function SafeSetFont(targetRange As Range, fontName As String, fontSize As Long) As Boolean
    On Error GoTo ErrorHandler

    ' Aplica formatação de fonte de forma segura
    With targetRange.Font
        If fontName <> "" Then .Name = fontName
        If fontSize > 0 Then .size = fontSize
        .Color = wdColorAutomatic
    End With

    SafeSetFont = True
    Exit Function

ErrorHandler:
    SafeSetFont = False
    LogMessage "Erro ao aplicar fonte: " & Err.Description & " - Range: " & Left(targetRange.text, 20), LOG_LEVEL_WARNING
End Function

Private Function SafeSetParagraphFormat(para As Paragraph, alignment As Long, leftIndent As Single, firstLineIndent As Single) As Boolean
    On Error GoTo ErrorHandler

    With para.Format
        If alignment >= 0 Then .alignment = alignment
        If leftIndent >= 0 Then .leftIndent = leftIndent
        If firstLineIndent >= 0 Then .firstLineIndent = firstLineIndent
    End With

    SafeSetParagraphFormat = True
    Exit Function

ErrorHandler:
    SafeSetParagraphFormat = False
    LogMessage "Erro ao aplicar formatação de parágrafo: " & Err.Description, LOG_LEVEL_WARNING
End Function

Private Function SafeHasVisualContent(para As Paragraph) As Boolean
    On Error GoTo SafeMode

    ' Verificação padrão mais robusta
    Dim hasImages As Boolean
    Dim hasShapes As Boolean

    ' Verifica imagens inline de forma segura
    hasImages = (para.Range.InlineShapes.count > 0)

    ' Verifica shapes flutuantes de forma segura
    hasShapes = False
    If Not hasImages Then
        Dim shp As shape
        For Each shp In para.Range.ShapeRange
            hasShapes = True
            Exit For
        Next shp
    End If

    SafeHasVisualContent = hasImages Or hasShapes
    Exit Function

SafeMode:
    On Error GoTo ErrorHandler
    ' Método alternativo mais simples
    SafeHasVisualContent = (para.Range.InlineShapes.count > 0)
    Exit Function

ErrorHandler:
    ' Em caso de erro, assume que não há conteúdo visual
    SafeHasVisualContent = False
End Function

'================================================================================
' SAFE FIND/REPLACE OPERATIONS
'================================================================================
Private Function SafeFindReplace(doc As Document, findText As String, replaceText As String, Optional useWildcards As Boolean = False) As Long
    On Error GoTo ErrorHandler

    Dim findCount As Long
    findCount = 0

    ' Configuração segura de Find/Replace
    With doc.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = findText
        .Replacement.text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = useWildcards  ' Parâmetro controlado
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        ' Executa a substituição e conta ocorrências
        Do While .Execute(Replace:=True)
            findCount = findCount + 1
            ' Limite de segurança para evitar loops infinitos
            If findCount > 10000 Then
                LogMessage "Limite de substituições atingido para: " & findText, LOG_LEVEL_WARNING
                Exit Do
            End If
        Loop
    End With

    SafeFindReplace = findCount
    Exit Function

ErrorHandler:
    SafeFindReplace = 0
    LogMessage "Erro na operação Find/Replace: " & findText & " -> " & replaceText & " | " & Err.Description, LOG_LEVEL_WARNING
End Function

'================================================================================
' ACESSO SEGURO A CARACTERES
'================================================================================
Private Function SafeGetLastCharacter(rng As Range) As String
    On Error GoTo ErrorHandler

    Dim charCount As Long
    charCount = SafeGetCharacterCount(rng)

    If charCount > 0 Then
        SafeGetLastCharacter = rng.Characters(charCount).text
    Else
        SafeGetLastCharacter = ""
    End If
    Exit Function

ErrorHandler:
    ' Método alternativo usando Right()
    On Error GoTo FinalFallback
    SafeGetLastCharacter = Right(rng.text, 1)
    Exit Function

FinalFallback:
    SafeGetLastCharacter = ""
End Function

'================================================================================
' FUNÇÕES DE CAMINHO - Estrutura do projeto
'================================================================================

'--------------------------------------------------------------------------------
' GetProjectRootPath - Retorna caminho raiz do projeto chainsaw
'--------------------------------------------------------------------------------
Private Function GetProjectRootPath() As String
    GetProjectRootPath = Environ("USERPROFILE") & "\chainsaw"
End Function

'--------------------------------------------------------------------------------
' GetChainsawBackupsPath - Retorna caminho para backups
'--------------------------------------------------------------------------------
Private Function GetChainsawBackupsPath() As String
    GetChainsawBackupsPath = Environ("TEMP") & "\.chainsaw\props\backups"
End Function

'--------------------------------------------------------------------------------
' GetChainsawRecoveryPath - Retorna caminho para recovery temporario
'--------------------------------------------------------------------------------
Private Function GetChainsawRecoveryPath() As String
    GetChainsawRecoveryPath = GetProjectRootPath() & "\props\recovery_tmp"
End Function

'--------------------------------------------------------------------------------
' GetChainsawLogsPath - Retorna caminho para logs
'--------------------------------------------------------------------------------
Private Function GetChainsawLogsPath() As String
    GetChainsawLogsPath = GetProjectRootPath() & "\source\logs"
End Function

'--------------------------------------------------------------------------------
' EnsureChainsawFolders - Cria estrutura de pastas do projeto se não existir
'--------------------------------------------------------------------------------
Private Sub EnsureChainsawFolders()
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim propsPath As String
    propsPath = GetProjectRootPath() & "\props"

    Dim sourcePath As String
    sourcePath = GetProjectRootPath() & "\source"

    ' Cria pasta props
    If Not fso.FolderExists(propsPath) Then
        fso.CreateFolder propsPath
    End If

    ' Cria pasta source
    If Not fso.FolderExists(sourcePath) Then
        fso.CreateFolder sourcePath
    End If

    ' Cria pasta backups (sempre em %TEMP%\.chainsaw\props\backups)
    Dim chainsawTempRoot As String
    chainsawTempRoot = Environ("TEMP") & "\.chainsaw"

    Dim chainsawTempProps As String
    chainsawTempProps = chainsawTempRoot & "\props"

    If Not fso.FolderExists(chainsawTempRoot) Then
        fso.CreateFolder chainsawTempRoot
    End If

    If Not fso.FolderExists(chainsawTempProps) Then
        fso.CreateFolder chainsawTempProps
    End If

    If Not fso.FolderExists(GetChainsawBackupsPath()) Then
        fso.CreateFolder GetChainsawBackupsPath()
    End If

    ' Cria pasta recovery_tmp
    If Not fso.FolderExists(GetChainsawRecoveryPath()) Then
        fso.CreateFolder GetChainsawRecoveryPath()
    End If

    ' Cria pasta logs
    If Not fso.FolderExists(GetChainsawLogsPath()) Then
        fso.CreateFolder GetChainsawLogsPath()
    End If

    Set fso = Nothing
End Sub

'================================================================================
' SISTEMA DE REGISTRO DE LOGS
'================================================================================

'--------------------------------------------------------------------------------
' WriteTextUTF8 - Escreve texto em arquivo com encoding UTF-8
'--------------------------------------------------------------------------------
Private Sub WriteTextUTF8(filePath As String, textContent As String, Optional appendMode As Boolean = False)
    On Error GoTo ErrorHandler

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open

    ' Se modo append, lê conteúdo existente primeiro
    If appendMode And Dir(filePath) <> "" Then
        stream.LoadFromFile filePath
        stream.Position = stream.size
    End If

    ' Escreve o novo conteúdo
    stream.WriteText textContent, 1 ' adWriteLine

    ' Salva com UTF-8
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing

    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
End Sub

Private Sub EnforceLogRetention(logFolder As String, logPrefix As String, Optional maxFiles As Long = 3)
    On Error GoTo CleanExit

    If maxFiles < 1 Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(logFolder) Then GoTo CleanExit

    Dim folder As Object
    Set folder = fso.GetFolder(logFolder)

    Dim sortedList As Object
    Set sortedList = CreateObject("System.Collections.ArrayList")

    Dim fileItem As Object
    Dim prefixLower As String
    prefixLower = LCase(logPrefix)

    For Each fileItem In folder.Files
        If LCase(fileItem.Name) Like prefixLower & "*.log" Then
            sortedList.Add Format(fileItem.DateLastModified, "yyyymmddHHMMSS") & "|" & fileItem.Path
        End If
    Next fileItem

    If sortedList.count <= maxFiles Then GoTo CleanExit

    sortedList.Sort
    sortedList.Reverse

    Dim idx As Long
    For idx = maxFiles To sortedList.count - 1
        Dim parts() As String
        parts = Split(sortedList(idx), "|")
        On Error Resume Next
        fso.DeleteFile parts(1), True
        On Error GoTo CleanExit
    Next idx

CleanExit:
    On Error Resume Next
    Set sortedList = Nothing
    Set folder = Nothing
    Set fso = Nothing
End Sub

Private Function InitializeLogging(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim logFolder As String
    Dim docNameClean As String
    Dim fileNum As Integer
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Garante que a estrutura de pastas do projeto existe
    EnsureChainsawFolders

    ' SEMPRE USA source\logs para todos os documentos
    logFolder = GetChainsawLogsPath() & "\"

    ' Garante que a pasta de logs existe antes de criar o arquivo
    If Not fso.FolderExists(logFolder) Then
        On Error Resume Next
        fso.CreateFolder logFolder
        On Error GoTo ErrorHandler
    End If

    If Not fso.FolderExists(logFolder) Then
        InitializeLogging = False
        loggingEnabled = False
        Exit Function
    End If

    ' Sanitiza nome do documento para uso em arquivo
    docNameClean = doc.Name
    docNameClean = Replace(docNameClean, ".doc", "")
    docNameClean = Replace(docNameClean, ".docx", "")
    docNameClean = Replace(docNameClean, ".docm", "")
    docNameClean = SanitizeFileName(docNameClean)

    ' Define nome do arquivo de log com timestamp
    logFilePath = logFolder & "chainsaw_" & Format(Now, "yyyymmdd_HHmmss") & "_" & docNameClean & ".log"

    ' Inicializa contadores e controles
    errorCount = 0
    warningCount = 0
    infoCount = 0
    logBufferEnabled = False
    logBuffer = ""
    lastFlushTime = Now
    logFileHandle = 0

    ' Cria arquivo de log com informações de contexto usando UTF-8
    Dim headerText As String
    headerText = String(80, "=") & vbCrLf
    headerText = headerText & "CHAINSAW - LOG DE PROCESSAMENTO DE DOCUMENTO" & vbCrLf
    headerText = headerText & String(80, "=") & vbCrLf & vbCrLf
    headerText = headerText & "[SESSÃO]" & vbCrLf
    headerText = headerText & "  Início: " & Format(Now, "dd/mm/yyyy HH:mm:ss") & vbCrLf
    headerText = headerText & "  ID: " & Format(Now, "yyyymmddHHmmss") & vbCrLf & vbCrLf
    headerText = headerText & "[AMBIENTE]" & vbCrLf
    headerText = headerText & "  Usuário: " & Environ("USERNAME") & vbCrLf
    headerText = headerText & "  Computador: " & Environ("COMPUTERNAME") & vbCrLf
    headerText = headerText & "  Domínio: " & Environ("USERDOMAIN") & vbCrLf
    headerText = headerText & "  SO: Windows " & GetWindowsVersion() & vbCrLf
    headerText = headerText & "  Word: " & Application.version & " (" & GetWordVersionName() & ")" & vbCrLf & vbCrLf
    headerText = headerText & "[DOCUMENTO]" & vbCrLf
    headerText = headerText & "  Nome: " & doc.Name & vbCrLf
    headerText = headerText & "  Caminho: " & IIf(doc.Path = "", "(Não salvo)", doc.Path) & vbCrLf
    headerText = headerText & "  Tamanho: " & GetDocumentSize(doc) & vbCrLf
    headerText = headerText & "  Parágrafos: " & doc.Paragraphs.count & vbCrLf
    headerText = headerText & "  Páginas: " & doc.ComputeStatistics(wdStatisticPages) & vbCrLf
    headerText = headerText & "  Proteção: " & GetProtectionType(doc) & vbCrLf
    headerText = headerText & "  Idioma: " & doc.Range.LanguageID & vbCrLf & vbCrLf
    headerText = headerText & "[CONFIGURAÇÃO]" & vbCrLf
    headerText = headerText & "  Debug: " & IIf(DEBUG_MODE, "Ativado", "Desativado") & vbCrLf
    headerText = headerText & "  Log: " & logFilePath & vbCrLf
    headerText = headerText & "  Backup: " & GetChainsawBackupsPath() & "\" & vbCrLf & vbCrLf
    headerText = headerText & String(80, "=") & vbCrLf & vbCrLf

    ' Escreve cabeçalho em UTF-8
    WriteTextUTF8 logFilePath, headerText, False

    ' Enforces log retention limit for this routine
    EnforceLogRetention logFolder, "chainsaw_", 3

    loggingEnabled = True
    InitializeLogging = True

    Exit Function

ErrorHandler:
    On Error Resume Next
    logFileHandle = 0
    loggingEnabled = False
    InitializeLogging = False
    Debug.Print "ERRO CRÍTICO: Falha ao inicializar logging - " & Err.Description
End Function

Private Sub LogMessage(message As String, Optional level As Long = LOG_LEVEL_INFO)
    On Error GoTo ErrorHandler

    If Not loggingEnabled Then Exit Sub

    Dim levelText As String
    Dim levelPrefix As String
    Dim fileNum As Integer
    Dim formattedMessage As String
    Dim timeStamp As String
    Dim elapsedTime As String

    ' Calcula tempo decorrido desde início
    If executionStartTime > 0 Then
        Dim elapsed As Double
        elapsed = (Now - executionStartTime) * 86400 ' Converte para segundos
        elapsedTime = Format(Int(elapsed / 60), "00") & ":" & Format(elapsed Mod 60, "00.0")
    Else
        elapsedTime = "00:00.0"
    End If

    ' Define nível e incrementa contadores
    Select Case level
        Case LOG_LEVEL_INFO
            levelText = "INFO "
            levelPrefix = "?"
            infoCount = infoCount + 1
        Case LOG_LEVEL_WARNING
            levelText = "WARN "
            levelPrefix = "?"
            warningCount = warningCount + 1
        Case LOG_LEVEL_ERROR
            levelText = "ERROR"
            levelPrefix = "?"
            errorCount = errorCount + 1
        Case Else
            levelText = "DEBUG"
            levelPrefix = "?"
    End Select

    ' Formata mensagem com timestamp, tempo decorrido e nível
    timeStamp = Format(Now, "HH:mm:ss.") & Format((Timer * 1000) Mod 1000, "000")
    formattedMessage = timeStamp & " [" & elapsedTime & "] " & levelText & " " & levelPrefix & " " & message

    ' Debug mode output para console VBA
    If DEBUG_MODE Then
        Debug.Print formattedMessage
    End If

    ' Buffer para reduzir I/O quando não for erro crítico
    If level = LOG_LEVEL_ERROR Or Len(logBuffer) > 4096 Or (Now - lastFlushTime) > (5 / 86400) Then
        ' Escreve imediatamente: erros, buffer cheio (>4KB), ou 5+ segundos desde último flush
        FlushLogBuffer

        ' Escreve mensagem em UTF-8
        WriteTextUTF8 logFilePath, formattedMessage, True

        lastFlushTime = Now
    Else
        ' Adiciona ao buffer para flush posterior (otimização de performance)
        logBuffer = logBuffer & formattedMessage & vbCrLf
    End If

    Exit Sub

ErrorHandler:
    On Error Resume Next
    If fileNum > 0 Then Close #fileNum
    Debug.Print "FALHA NO LOG: " & message & " | Erro: " & Err.Description
End Sub

Private Sub FlushLogBuffer()
    On Error Resume Next

    If Len(logBuffer) = 0 Then Exit Sub

    ' Escreve buffer em UTF-8
    WriteTextUTF8 logFilePath, logBuffer, True

    logBuffer = ""
    lastFlushTime = Now
End Sub

'================================================================================
' FUNÇÕES AUXILIARES DE LOG
'================================================================================
Private Sub LogSection(sectionName As String)
    On Error Resume Next

    If Not loggingEnabled Then Exit Sub

    FlushLogBuffer

    ' Cria texto de seção
    Dim sectionText As String
    sectionText = vbCrLf & String(80, "-") & vbCrLf
    sectionText = sectionText & "SEÇÃO: " & UCase(sectionName) & vbCrLf
    sectionText = sectionText & String(80, "-")

    ' Escreve em UTF-8
    WriteTextUTF8 logFilePath, sectionText, True

    lastFlushTime = Now
End Sub

Private Sub LogStepStart(stepName As String)
    On Error Resume Next
    LogMessage "? Iniciando: " & stepName, LOG_LEVEL_INFO
End Sub

Private Sub LogStepComplete(stepName As String, Optional details As String = "")
    On Error Resume Next
    Dim msg As String
    msg = "? Concluído: " & stepName
    If Len(details) > 0 Then msg = msg & " | " & details
    LogMessage msg, LOG_LEVEL_INFO
End Sub

Private Sub LogStepSkipped(stepName As String, reason As String)
    On Error Resume Next
    LogMessage "? Ignorado: " & stepName & " | Motivo: " & reason, LOG_LEVEL_INFO
End Sub

Private Sub LogMetric(metricName As String, value As Variant, Optional unit As String = "")
    On Error Resume Next
    Dim msg As String
    msg = "?? " & metricName & ": " & CStr(value)
    If Len(unit) > 0 Then msg = msg & " " & unit
    LogMessage msg, LOG_LEVEL_INFO
End Sub

Private Sub SafeFinalizeLogging()
    On Error GoTo ErrorHandler

    If Not loggingEnabled Then Exit Sub

    Dim fileNum As Integer
    Dim statusText As String
    Dim statusIcon As String
    Dim duration As Double
    Dim durationText As String
    Dim totalEvents As Long

    ' Flush pendente no buffer
    FlushLogBuffer

    ' Calcula duração total
    duration = (Now - executionStartTime) * 86400
    If duration < 60 Then
        durationText = Format(duration, "0.0") & "s"
    ElseIf duration < 3600 Then
        durationText = Format(Int(duration / 60), "0") & "m " & Format(duration Mod 60, "00") & "s"
    Else
        durationText = Format(Int(duration / 3600), "0") & "h " & Format(Int((duration Mod 3600) / 60), "00") & "m"
    End If

    ' Determina status final
    If formattingCancelled Then
        statusText = "CANCELADO PELO USUÁRIO"
        statusIcon = "?"
    ElseIf errorCount > 0 Then
        statusText = "CONCLUÍDO COM ERROS"
        statusIcon = "?"
    ElseIf warningCount > 0 Then
        statusText = "CONCLUÍDO COM AVISOS"
        statusIcon = "?"
    Else
        statusText = "CONCLUÍDO COM SUCESSO"
        statusIcon = "?"
    End If

    totalEvents = infoCount + warningCount + errorCount

    ' Escreve rodapé estruturado em UTF-8
    Dim footerText As String
    footerText = vbCrLf & String(80, "=") & vbCrLf
    footerText = footerText & "RESUMO DA SESSÃO" & vbCrLf
    footerText = footerText & String(80, "=") & vbCrLf & vbCrLf
    footerText = footerText & "[STATUS]" & vbCrLf
    footerText = footerText & "  Final: " & statusText & " " & statusIcon & vbCrLf
    footerText = footerText & "  Término: " & Format(Now, "dd/mm/yyyy HH:mm:ss") & vbCrLf
    footerText = footerText & "  Duração: " & durationText & vbCrLf & vbCrLf
    footerText = footerText & "[ESTATÍSTICAS]" & vbCrLf
    footerText = footerText & "  Total de eventos: " & totalEvents & vbCrLf
    footerText = footerText & "  Informações: " & infoCount & " (" & Format(infoCount / IIf(totalEvents > 0, totalEvents, 1) * 100, "0.0") & "%)" & vbCrLf
    footerText = footerText & "  Avisos: " & warningCount & " (" & Format(warningCount / IIf(totalEvents > 0, totalEvents, 1) * 100, "0.0") & "%)" & vbCrLf
    footerText = footerText & "  Erros: " & errorCount & " (" & Format(errorCount / IIf(totalEvents > 0, totalEvents, 1) * 100, "0.0") & "%)" & vbCrLf & vbCrLf

    ' Adiciona informações de performance
    If totalEvents > 0 Then
        footerText = footerText & "[PERFORMANCE]" & vbCrLf
        footerText = footerText & "  Eventos/segundo: " & Format(totalEvents / IIf(duration > 0, duration, 1), "0.0") & vbCrLf
        footerText = footerText & "  Tempo médio/evento: " & Format((duration / totalEvents) * 1000, "0.0") & "ms" & vbCrLf & vbCrLf
    End If

    ' Recomendações se houver problemas
    If errorCount > 0 Or warningCount > 5 Then
        footerText = footerText & "[RECOMENDAÇÕES]" & vbCrLf
        If errorCount > 0 Then
            footerText = footerText & "  • Verifique os erros acima e corrija problemas no documento" & vbCrLf
        End If
        If warningCount > 5 Then
            footerText = footerText & "  • Múltiplos avisos detectados - revise o documento manualmente" & vbCrLf
        End If
        If duration > 60 Then
            footerText = footerText & "  • Processamento demorado - considere otimizar o documento" & vbCrLf
        End If
        footerText = footerText & vbCrLf
    End If

    footerText = footerText & String(80, "=") & vbCrLf
    footerText = footerText & "FIM DO LOG" & vbCrLf
    footerText = footerText & String(80, "=")

    ' Escreve footer em UTF-8
    WriteTextUTF8 logFilePath, footerText, True

    ' Limpa variáveis
    loggingEnabled = False
    logBuffer = ""
    logFileHandle = 0

    Exit Sub

ErrorHandler:
    On Error Resume Next
    If fileNum > 0 Then Close #fileNum
    loggingEnabled = False
    Debug.Print "ERRO CRÍTICO ao finalizar logging: " & Err.Description
End Sub

'================================================================================
' UTILITY: GET PROTECTION TYPE
'================================================================================
Private Function GetProtectionType(doc As Document) As String
    On Error Resume Next

    Select Case doc.protectionType
        Case wdNoProtection: GetProtectionType = "Sem proteção"
        Case 1: GetProtectionType = "Protegido contra revisões"
        Case 2: GetProtectionType = "Protegido contra comentários"
        Case 3: GetProtectionType = "Protegido contra formulários"
        Case 4: GetProtectionType = "Protegido contra leitura"
        Case Else: GetProtectionType = "Tipo desconhecido (" & doc.protectionType & ")"
    End Select
End Function

'================================================================================
' UTILITY: GET DOCUMENT SIZE
'================================================================================
Private Function GetDocumentSize(doc As Document) As String
    On Error Resume Next

    Dim size As Long
    size = doc.BuiltInDocumentProperties("Number of Characters").value * 2

    If Err.Number <> 0 Then
        GetDocumentSize = "Desconhecido"
        Exit Function
    End If

    If size < 1024 Then
        GetDocumentSize = size & " bytes"
    ElseIf size < 1048576 Then
        GetDocumentSize = Format(size / 1024, "0.0") & " KB"
    Else
        GetDocumentSize = Format(size / 1048576, "0.0") & " MB"
    End If
End Function

'================================================================================
' UTILITY: SANITIZE FILE NAME
'================================================================================
Private Function SanitizeFileName(fileName As String) As String
    On Error Resume Next

    Dim result As String
    Dim invalidChars As String
    Dim i As Long

    result = fileName
    invalidChars = "\/:*?""<>|"

    For i = 1 To Len(invalidChars)
        result = Replace(result, Mid(invalidChars, i, 1), "_")
    Next i

    ' Limita tamanho
    If Len(result) > 50 Then
        result = Left(result, 50)
    End If

    SanitizeFileName = result
End Function

'================================================================================
' UTILITY: GET WINDOWS VERSION
'================================================================================
Private Function GetWindowsVersion() As String
    On Error Resume Next

    Dim osVersion As String
    osVersion = Environ("OS")

    If osVersion = "" Then osVersion = "Windows"

    GetWindowsVersion = osVersion
End Function

'================================================================================
' UTILITY: GET WORD VERSION NAME
'================================================================================
Private Function GetWordVersionName() As String
    On Error Resume Next

    Dim ver As String
    ver = Application.version

    Select Case ver
        Case "16.0": GetWordVersionName = "Word 2016/2019/2021/365"
        Case "15.0": GetWordVersionName = "Word 2013"
        Case "14.0": GetWordVersionName = "Word 2010"
        Case "12.0": GetWordVersionName = "Word 2007"
        Case "11.0": GetWordVersionName = "Word 2003"
        Case Else: GetWordVersionName = "Word " & ver
    End Select
End Function

'================================================================================
' GERENCIAMENTO DE ESTADO DA APLICAÇÃO
'================================================================================
Private Function SetAppState(Optional ByVal enabled As Boolean = True, Optional ByVal statusMsg As String = "", Optional ByVal preserveStatusBar As Boolean = False) As Boolean
    On Error GoTo ErrorHandler

    Dim success As Boolean
    success = True

    With Application
        On Error Resume Next
        .ScreenUpdating = enabled
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler

        On Error Resume Next
        .DisplayAlerts = IIf(enabled, wdAlertsAll, wdAlertsNone)
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler

        ' Nao modifica StatusBar se preserveStatusBar = True
        If Not preserveStatusBar Then
            If statusMsg <> "" Then
                On Error Resume Next
                .StatusBar = statusMsg
                If Err.Number <> 0 Then success = False
                On Error GoTo ErrorHandler
            ElseIf enabled Then
                On Error Resume Next
                .StatusBar = False
                If Err.Number <> 0 Then success = False
                On Error GoTo ErrorHandler
            End If
        End If

        On Error Resume Next
        .EnableCancelKey = 0
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
    End With

    SetAppState = success
    Exit Function

ErrorHandler:
    SetAppState = False
End Function

'================================================================================
' VERIFICAÇÕES GLOBAIS ANTES DA FORMATAÇÃO
'================================================================================
Private Function PreviousChecking(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    LogSection "VERIFICAÇÕES INICIAIS"
    LogStepStart "Validação de documento"

    If doc Is Nothing Then
        Application.StatusBar = "Erro: Documento inacessível"
        LogMessage "Documento não acessível para verificação", LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If doc.Type <> wdTypeDocument Then
        Application.StatusBar = "Erro: Tipo nao suportado"
        LogMessage "Tipo de documento nao suportado: " & doc.Type, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    ' Verifica se a primeira palavra e um tipo valido de propositura
    If Not ValidateProposituraType(doc) Then
        LogMessage "Usuario cancelou processamento - tipo de propositura nao reconhecido", LOG_LEVEL_INFO
        PreviousChecking = False
        Exit Function
    End If

    If doc.protectionType <> wdNoProtection Then
        Dim protectionType As String
        protectionType = GetProtectionType(doc)
        Application.StatusBar = "Erro: Documento protegido"
        LogMessage "Documento protegido detectado: " & protectionType, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If doc.ReadOnly Then
        Application.StatusBar = "Erro: Somente leitura"
        LogMessage "Documento em modo somente leitura: " & doc.FullName, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If Not CheckDiskSpace(doc) Then
        Application.StatusBar = "Erro: Espaço insuficiente"
        LogMessage "Espaço em disco insuficiente para operação segura", LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If Not ValidateDocumentStructure(doc) Then
        LogMessage "Estrutura do documento validada com avisos", LOG_LEVEL_WARNING
    End If

    ' Verifica consistência de endereços entre 2º e 3º parágrafos
    If Not ValidateAddressConsistency(doc) Then
        LogMessage "Recomendação para verificar endereços foi exibida ao usuário", LOG_LEVEL_INFO
    End If

    ' Verifica presença de possíveis dados sensíveis
    If Not CheckSensitiveData(doc) Then
        LogMessage "Aviso de dados sensíveis foi exibido ao usuário", LOG_LEVEL_INFO
    End If

    LogStepComplete "Validação de documento", "Todas as verificações passaram"
    LogMessage "Verificações de segurança concluídas com sucesso", LOG_LEVEL_INFO
    PreviousChecking = True
    Exit Function

ErrorHandler:
    Application.StatusBar = "Erro na verificação"
    LogMessage "Erro durante verificações: " & Err.Description, LOG_LEVEL_ERROR
    PreviousChecking = False
End Function

'================================================================================
' VERIFICAÇÃO DE ESPAÇO EM DISCO
'================================================================================
Private Function CheckDiskSpace(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ' Verificação simplificada - assume espaço suficiente se não conseguir verificar
    Dim fso As Object
    Dim drive As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If doc.Path <> "" Then
        Set drive = fso.GetDrive(Left(doc.Path, 3))
    Else
        Set drive = fso.GetDrive(Left(Environ("TEMP"), 3))
    End If

    ' Verificação básica - 10MB mínimo
    If drive.AvailableSpace < 10485760 Then ' 10MB em bytes
        LogMessage "Espaço em disco muito baixo", LOG_LEVEL_WARNING
        CheckDiskSpace = False
    Else
        CheckDiskSpace = True
    End If

    Exit Function

ErrorHandler:
    ' Se não conseguir verificar, assume que há espaço suficiente
    CheckDiskSpace = True
End Function

'================================================================================
' ROTINA PRINCIPAL DE FORMATAÇÃO
'================================================================================
Private Function PreviousFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ' Formatações básicas de página e estrutura
    If Not ApplyPageSetup(doc) Then
        LogMessage "Falha na configuração de página", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If

    LogSection "LIMPEZA E FORMATAÇÃO"

    ' Limpeza e formatações otimizadas
    LogStepStart "Limpeza de formatação"
    ClearAllFormatting doc
    LogStepComplete "Limpeza de formatação"

    LogStepStart "Normalização de quebras"
    ReplaceLineBreaksWithParagraphBreaks doc
    RemovePageBreaks doc
    LogStepComplete "Normalização de quebras"

    LogStepStart "Limpeza estrutural"
    RemovePageNumberLines doc
    CleanDocumentStructure doc
    RemoveAllTabMarks doc
    LogStepComplete "Limpeza estrutural"

    LogStepStart "Formatação de título"
    FormatDocumentTitle doc
    LogStepComplete "Formatação de título"

    ' Formatações principais - Usa versão otimizada se cache disponível
    LogStepStart "Aplicação de fonte padrão"
    If cacheEnabled Then
        If Not ApplyStdFontOptimized(doc) Then
            LogMessage "Falha na formatação de fontes (otimizada) - tentando método tradicional", LOG_LEVEL_WARNING
            If Not ApplyStdFont(doc) Then
                LogMessage "Falha na formatação de fontes", LOG_LEVEL_ERROR
                PreviousFormatting = False
                Exit Function
            End If
        End If
    Else
        If Not ApplyStdFont(doc) Then
            LogMessage "Falha na formatação de fontes", LOG_LEVEL_ERROR
            PreviousFormatting = False
            Exit Function
        End If
    End If
    LogStepComplete "Aplicação de fonte padrão", doc.Paragraphs.count & " parágrafos"

    LogStepStart "Aplicação de formatação de parágrafos"
    If Not ApplyStdParagraphs(doc) Then
        LogMessage "Falha na formatação de parágrafos", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    LogStepComplete "Aplicação de formatação de parágrafos"

    LogSection "FORMATAÇÕES ESPECÍFICAS"

    LogStepStart "Formatação de parágrafos 1 e 2"
    FormatFirstParagraph doc
    FormatSecondParagraph doc
    LogStepComplete "Formatação de parágrafos 1 e 2"

    LogStepStart "Formatação de considerandos"
    FormatConsiderandoParagraphs doc
    LogStepComplete "Formatação de considerandos"

    LogStepStart "Aplicação de substituições de texto"
    ApplyTextReplacements doc
    LogStepComplete "Aplicação de substituições de texto"

    LogStepStart "Remoção de marca d'água e inserção de carimbo"
    RemoveWatermark doc
    InsertHeaderstamp doc
    LogStepComplete "Remoção de marca d'água e inserção de carimbo"

    LogSection "LIMPEZA FINAL"

    LogStepStart "Limpeza de espaços múltiplos"
    CleanMultipleSpaces doc
    LogStepComplete "Limpeza de espaços múltiplos"

    LogStepStart "Controle de linhas em branco"
    LimitSequentialEmptyLines doc
    LogStepComplete "Controle de linhas em branco"

    LogStepStart "Substituição de datas do plenário"
    ReplacePlenarioDateParagraph doc
    LogStepComplete "Substituição de datas do plenário"

    LogSection "FINALIZAÇÃO"

    LogStepStart "Configuração de visualização"
    ConfigureDocumentView doc
    LogStepComplete "Configuração de visualização"

    LogStepStart "Inserção de rodapé"
    If Not InsertFooterStamp(doc) Then
        LogMessage "Falha na inserção do rodapé", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    LogStepComplete "Inserção de rodapé"

    LogStepStart "Ajustes finais de negrito e formatação"
    ApplyBoldToSpecialParagraphs doc
    FormatVereadorParagraphs doc
    LogStepComplete "Ajustes finais de negrito e formatação"

    LogStepStart "Formatações especiais (diante do exposto, requeiro)"
    FormatDianteDoExposto doc
    FormatRequeiroParagraphs doc
    FormatPorTodasRazoesParagraphs doc
    LogStepComplete "Formatações especiais (diante do exposto, requeiro)"

    LogStepStart "Remoção de realces e bordas"
    RemoveAllHighlightsAndBorders doc
    LogStepComplete "Remoção de realces e bordas"

    LogStepStart "Remoção de páginas vazias no final"
    RemoveEmptyPagesAtEnd doc
    LogStepComplete "Remoção de páginas vazias no final"

    LogStepStart "Aplicação de formatação final universal"
    ApplyUniversalFinalFormatting doc
    LogStepComplete "Aplicação de formatação final universal"

    LogStepStart "Adicão de espaçamento especial (ementa, justificativa, data)"
    AddSpecialElementsSpacing doc
    LogStepComplete "Adicão de espaçamento especial (ementa, justificativa, data)"

    LogMessage "Formatação completa aplicada com sucesso", LOG_LEVEL_INFO
    LogMetric "Total de parágrafos", doc.Paragraphs.count
    PreviousFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "Erro durante formatação: " & Err.Description, LOG_LEVEL_ERROR
    PreviousFormatting = False
End Function

'================================================================================
' CONFIGURAÇÃO DE PÁGINA
'================================================================================
Private Function ApplyPageSetup(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    With doc.PageSetup
        .TopMargin = CentimetersToPoints(TOP_MARGIN_CM)
        .BottomMargin = CentimetersToPoints(BOTTOM_MARGIN_CM)
        .LeftMargin = CentimetersToPoints(LEFT_MARGIN_CM)
        .RightMargin = CentimetersToPoints(RIGHT_MARGIN_CM)
        .HeaderDistance = CentimetersToPoints(HEADER_DISTANCE_CM)
        .FooterDistance = CentimetersToPoints(FOOTER_DISTANCE_CM)
        .Gutter = 0
        .Orientation = wdOrientPortrait
    End With

    ' Configuração de página aplicada (sem log detalhado para performance)
    ApplyPageSetup = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na configuração de página: " & Err.Description, LOG_LEVEL_ERROR
    ApplyPageSetup = False
End Function

'================================================================================
' FORMATAÇÃO DE FONTE OTIMIZADA COM CACHE
'================================================================================
Private Function ApplyStdFontOptimized(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If Not cacheEnabled Then
        ' Fallback para método tradicional se cache não estiver disponível
        ApplyStdFontOptimized = ApplyStdFont(doc)
        Exit Function
    End If

    Dim i As Long
    Dim para As Paragraph
    Dim cache As paragraphCache
    Dim formattedCount As Long
    Dim startTime As Double

    startTime = Timer
    formattedCount = 0

    LogMessage "Aplicando fonte padrao (modo otimizado com cache)...", LOG_LEVEL_INFO

    ' Valida cache antes de processar
    If cacheSize < 1 Then
        LogMessage "Cache vazio - usando metodo tradicional", LOG_LEVEL_INFO
        ApplyStdFontOptimized = ApplyStdFont(doc)
        Exit Function
    End If

    ' Valida limites do array
    On Error Resume Next
    Dim cacheUpperBound As Long
    cacheUpperBound = UBound(paragraphCache)
    If Err.Number <> 0 Or cacheUpperBound < 1 Then
        Err.Clear
        On Error GoTo ErrorHandler
        LogMessage "Array de cache invalido - usando metodo tradicional", LOG_LEVEL_WARNING
        ApplyStdFontOptimized = ApplyStdFont(doc)
        Exit Function
    End If
    On Error GoTo ErrorHandler

    ' Ajusta cacheSize se necessario
    If cacheSize > cacheUpperBound Then
        cacheSize = cacheUpperBound
    End If

    ' SINGLE PASS - Processa todos os paragrafos em uma passagem usando cache
    For i = 1 To cacheSize
        cache = paragraphCache(i)

        ' Pula paragrafos vazios ou com imagens
        If Not cache.needsFormatting Then
            GoTo NextParagraph
        End If

        ' Validacao do indice do paragrafo no documento
        If cache.index < 1 Or cache.index > doc.Paragraphs.count Then
            LogMessage "Erro: Indice de paragrafo invalido (" & cache.index & ")", LOG_LEVEL_WARNING
            GoTo NextParagraph
        End If

        Set para = doc.Paragraphs(cache.index)

        ' Aplica fonte padrao
        On Error Resume Next
        With para.Range.Font
            .Name = STANDARD_FONT
            .size = STANDARD_FONT_SIZE
            .Color = wdColorAutomatic

            ' Remove sublinhado exceto para titulo (primeiro paragrafo com texto)
            If i > 3 Then
                .Underline = wdUnderlineNone
            End If

            ' Remove negrito exceto para paragrafos especiais
            If Not cache.isSpecial Or cache.specialType = "vereador" Then
                .Bold = False
            End If
        End With

        If Err.Number = 0 Then
            formattedCount = formattedCount + 1
        Else
            LogMessage "Erro ao formatar paragrafo " & i & ": " & Err.Description, LOG_LEVEL_WARNING
            Err.Clear
        End If
        On Error GoTo ErrorHandler

NextParagraph:
        ' Atualiza progresso a cada 500 paragrafos
        If i Mod 500 = 0 Then
            DoEvents ' Permite cancelamento
        End If
    Next i

    Dim elapsed As Single
    elapsed = Timer - startTime

    ' Marca documento como modificado se houve formatacao
    If formattedCount > 0 Then documentDirty = True

    LogMessage "Fonte padrao aplicada: " & formattedCount & " paragrafos em " & Format(elapsed, "0.00") & "s", LOG_LEVEL_INFO
    ApplyStdFontOptimized = True
    Exit Function

ErrorHandler:
    LogMessage "Erro em ApplyStdFontOptimized: " & Err.Description, LOG_LEVEL_ERROR
    ApplyStdFontOptimized = False
End Function

'================================================================================
' FORMATAÇÃO DE FONTE (MÉTODO TRADICIONAL - FALLBACK)
'================================================================================
Private Function ApplyStdFont(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim hasInlineImage As Boolean
    Dim i As Long
    Dim formattedCount As Long
    Dim skippedCount As Long
    Dim underlineRemovedCount As Long
    Dim isTitle As Boolean
    Dim hasConsiderando As Boolean
    Dim needsUnderlineRemoval As Boolean
    Dim needsBoldRemoval As Boolean
    Dim paraCount As Long

    ' Cache do count para performance
    paraCount = doc.Paragraphs.count

    For i = paraCount To 1 Step -1
        If i > doc.Paragraphs.count Then Exit For ' Proteção dinâmica
        Set para = doc.Paragraphs(i)

        ' Early exit se processou demais (proteção contra documentos gigantes)
        If formattedCount > 50000 Then
            LogMessage "Limite de processamento atingido em ApplyStdFont (50000 parágrafos)", LOG_LEVEL_WARNING
            Exit For
        End If
        hasInlineImage = False
        isTitle = False
        hasConsiderando = False
        needsUnderlineRemoval = False
        needsBoldRemoval = False

        ' SUPER OTIMIZADO: Verificação prévia consolidada - uma única leitura das propriedades
        Dim paraFont As Font
        Set paraFont = para.Range.Font
        Dim needsFontFormatting As Boolean
        needsFontFormatting = (paraFont.Name <> STANDARD_FONT) Or _
                             (paraFont.size <> STANDARD_FONT_SIZE) Or _
                             (paraFont.Color <> wdColorAutomatic)

        ' Cache das verificações de formatação especial
        needsUnderlineRemoval = (paraFont.Underline <> wdUnderlineNone)
        needsBoldRemoval = (paraFont.Bold = True)

        ' Cache da contagem de InlineShapes para evitar múltiplas chamadas
        Dim inlineShapesCount As Long
        inlineShapesCount = para.Range.InlineShapes.count

        ' OTIMIZAÇÃO MÁXIMA: Se não precisa de nenhuma formatação, pula imediatamente
        If Not needsFontFormatting And Not needsUnderlineRemoval And Not needsBoldRemoval And inlineShapesCount = 0 Then
            formattedCount = formattedCount + 1
            GoTo NextParagraph
        End If

        If inlineShapesCount > 0 Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If

        ' OTIMIZADO: Verificação de conteúdo visual só quando necessário
        If Not hasInlineImage And (needsFontFormatting Or needsUnderlineRemoval Or needsBoldRemoval) Then
            If HasVisualContent(para) Then
                hasInlineImage = True
                skippedCount = skippedCount + 1
            End If
        End If

        ' OTIMIZADO: Verificação consolidada de tipo de parágrafo - uma única leitura do texto
        Dim paraFullText As String
        Dim isSpecialParagraph As Boolean
        isSpecialParagraph = False

        ' Só faz verificação de texto se for necessário para formatação especial
        If needsUnderlineRemoval Or needsBoldRemoval Then
            paraFullText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

            ' Verifica se é o primeiro parágrafo com texto (título) - otimizado
            If i <= 3 And para.Format.alignment = wdAlignParagraphCenter And paraFullText <> "" Then
                isTitle = True
            End If

            ' Verifica se o parágrafo começa com "considerando" - otimizado
            If Len(paraFullText) >= CONSIDERANDO_MIN_LENGTH And LCase(Left(paraFullText, CONSIDERANDO_MIN_LENGTH)) = CONSIDERANDO_PREFIX Then
                hasConsiderando = True
            End If

            ' Verifica se é um parágrafo especial - otimizado
            Dim cleanParaText As String
            cleanParaText = paraFullText
            ' Remove pontuação final para análise com proteção
            Dim punctCounter As Long
            punctCounter = 0
            Do While Len(cleanParaText) > 0 And (Right(cleanParaText, 1) = "." Or Right(cleanParaText, 1) = "," Or Right(cleanParaText, 1) = ":" Or Right(cleanParaText, 1) = ";") And punctCounter < 50
                cleanParaText = Left(cleanParaText, Len(cleanParaText) - 1)
                punctCounter = punctCounter + 1
            Loop
            cleanParaText = Trim(LCase(cleanParaText))

            ' Vereador NÃO é mais tratado como parágrafo especial (negrito deve ser removido)
            If cleanParaText = "justificativa" Or IsAnexoPattern(cleanParaText) Then
                isSpecialParagraph = True
                LogMessage "Parágrafo especial detectado em ApplyStdFont (negrito preservado): " & cleanParaText, LOG_LEVEL_INFO
            End If

            ' O parágrafo ANTERIOR a "vereador" não precisa mais preservar negrito
            Dim isBeforeVereador As Boolean
            isBeforeVereador = False
        End If

        ' FORMATAÇÃO PRINCIPAL - Só executa se necessário
        If needsFontFormatting Then
            If Not hasInlineImage Then
                ' Formatação rápida para parágrafos sem imagens usando método seguro
                If SafeSetFont(para.Range, STANDARD_FONT, STANDARD_FONT_SIZE) Then
                    formattedCount = formattedCount + 1
                Else
                    ' Fallback para método tradicional em caso de erro
                    With paraFont
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                        .Color = wdColorAutomatic
                    End With
                    formattedCount = formattedCount + 1
                End If
            Else
                ' NOVO: Formatação protegida para parágrafos COM imagens
                If ProtectImagesInRange(para.Range) Then
                    formattedCount = formattedCount + 1
                Else
                    ' Fallback: formatação básica segura CONSOLIDADA
                    Call FormatCharacterByCharacter(para, STANDARD_FONT, STANDARD_FONT_SIZE, wdColorAutomatic, False, False)
                    formattedCount = formattedCount + 1
                End If
            End If
        End If

        ' FORMATAÇÃO ESPECIAL CONSOLIDADA - Remove sublinhado e negrito em uma única passada
        If needsUnderlineRemoval Or needsBoldRemoval Then
            ' Determina quais formatações remover
            Dim removeUnderline As Boolean
            Dim removeBold As Boolean
            removeUnderline = needsUnderlineRemoval And Not isTitle
            removeBold = needsBoldRemoval And Not isTitle And Not hasConsiderando And Not isSpecialParagraph And Not isBeforeVereador

            ' Se precisa remover alguma formatação
            If removeUnderline Or removeBold Then
                If Not hasInlineImage Then
                    ' Formatação rápida para parágrafos sem imagens
                    If removeUnderline Then paraFont.Underline = wdUnderlineNone
                    If removeBold Then paraFont.Bold = False
                Else
                    ' Formatação protegida CONSOLIDADA para parágrafos com imagens
                    Call FormatCharacterByCharacter(para, "", 0, 0, removeUnderline, removeBold)
                End If

                If removeUnderline Then underlineRemovedCount = underlineRemovedCount + 1
            End If
        End If

NextParagraph:
    Next i

    ' Marca documento como modificado se houve formatacao
    If formattedCount > 0 Then documentDirty = True

    ' Log otimizado
    If skippedCount > 0 Then
        LogMessage "Fontes formatadas: " & formattedCount & " paragrafos (incluindo " & skippedCount & " com protecao de imagens)"
    End If

    ApplyStdFont = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação de fonte: " & Err.Description, LOG_LEVEL_ERROR
    ApplyStdFont = False
End Function

'================================================================================
' FORMATAÇÃO CARACTERE POR CARACTERE CONSOLIDADA
'================================================================================
Private Sub FormatCharacterByCharacter(para As Paragraph, fontName As String, fontSize As Long, fontColor As Long, removeUnderline As Boolean, removeBold As Boolean)
    On Error Resume Next

    Dim j As Long
    Dim charCount As Long
    Dim charRange As Range

    charCount = SafeGetCharacterCount(para.Range) ' Cache da contagem segura

    If charCount > 0 Then ' Verificação de segurança
        For j = 1 To charCount
            Set charRange = para.Range.Characters(j)
            If charRange.InlineShapes.count = 0 Then
                With charRange.Font
                    ' Aplica formatação de fonte se especificada
                    If fontName <> "" Then .Name = fontName
                    If fontSize > 0 Then .size = fontSize
                    If fontColor >= 0 Then .Color = fontColor

                    ' Remove formatações especiais se solicitado
                    If removeUnderline Then .Underline = wdUnderlineNone
                    If removeBold Then .Bold = False
                End With
            End If
        Next j
    End If
End Sub

'================================================================================
' FORMATAÇÃO DE PARÁGRAFOS
'================================================================================
Private Function ApplyStdParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim hasInlineImage As Boolean
    Dim paragraphIndent As Single
    Dim firstIndent As Single
    Dim rightMarginPoints As Single
    Dim i As Long
    Dim formattedCount As Long
    Dim skippedCount As Long
    Dim paraText As String
    Dim prevPara As Paragraph

    rightMarginPoints = 0

    ' Cache do count para performance
    Dim paraCount As Long
    paraCount = doc.Paragraphs.count

    For i = paraCount To 1 Step -1
        If i > doc.Paragraphs.count Then Exit For ' Proteção dinâmica
        Set para = doc.Paragraphs(i)
        hasInlineImage = False

        ' Early exit se processou demais
        If formattedCount > 50000 Then
            LogMessage "Limite de processamento atingido em ApplyStdParagraphs (50000 parágrafos)", LOG_LEVEL_WARNING
            Exit For
        End If

        If para.Range.InlineShapes.count > 0 Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If

        ' Proteção adicional: verifica outros tipos de conteúdo visual
        If Not hasInlineImage And HasVisualContent(para) Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If

        ' Aplica formatação de parágrafo para TODOS os parágrafos
        ' (independente se contêm imagens ou não)

        ' Limpeza robusta de espaços múltiplos - SEMPRE aplicada
        Dim cleanText As String
        cleanText = para.Range.text

        ' OTIMIZADO: Combinação de múltiplas operações de limpeza em um bloco
        If InStr(cleanText, "  ") > 0 Or InStr(cleanText, vbTab) > 0 Then
            ' Remove múltiplos espaços consecutivos com proteção
            Dim cleanCounter As Long
            cleanCounter = 0
            Do While InStr(cleanText, "  ") > 0 And cleanCounter < MAX_LOOP_ITERATIONS
                cleanText = Replace(cleanText, "  ", " ")
                cleanCounter = cleanCounter + 1
            Loop

            ' Remove espaços antes/depois de quebras de linha
            cleanText = Replace(cleanText, " " & vbCr, vbCr)
            cleanText = Replace(cleanText, vbCr & " ", vbCr)
            cleanText = Replace(cleanText, " " & vbLf, vbLf)
            cleanText = Replace(cleanText, vbLf & " ", vbLf)

            ' Remove tabs extras e converte para espaços com proteção
            cleanCounter = 0
            Do While InStr(cleanText, vbTab & vbTab) > 0 And cleanCounter < MAX_LOOP_ITERATIONS
                cleanText = Replace(cleanText, vbTab & vbTab, vbTab)
                cleanCounter = cleanCounter + 1
            Loop
            cleanText = Replace(cleanText, vbTab, " ")

            ' Limpeza final de espaços múltiplos com proteção
            cleanCounter = 0
            Do While InStr(cleanText, "  ") > 0 And cleanCounter < MAX_LOOP_ITERATIONS
                cleanText = Replace(cleanText, "  ", " ")
                cleanCounter = cleanCounter + 1
            Loop
        End If

        ' Verifica se é um parágrafo especial ANTES de limpar o texto
        Dim isSpecialFormatParagraph As Boolean
        isSpecialFormatParagraph = False

        Dim checkText As String
        checkText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        ' Remove pontuação final para análise com proteção
        Dim checkCounter As Long
        checkCounter = 0
        Do While Len(checkText) > 0 And (Right(checkText, 1) = "." Or Right(checkText, 1) = "," Or Right(checkText, 1) = ":" Or Right(checkText, 1) = ";") And checkCounter < 50
            checkText = Left(checkText, Len(checkText) - 1)
            checkCounter = checkCounter + 1
        Loop
        checkText = Trim(LCase(checkText))

        ' Verifica se é "Justificativa", "Anexo", "Anexos" ou padrão de vereador
        If checkText = JUSTIFICATIVA_TEXT Or IsAnexoPattern(checkText) Or IsVereadorPattern(checkText) Then
            isSpecialFormatParagraph = True
        End If

        ' Aplica o texto limpo APENAS se não há imagens E não é parágrafo especial
        If cleanText <> para.Range.text And Not hasInlineImage And Not isSpecialFormatParagraph Then
            para.Range.text = cleanText
        End If

        ' Formatação de parágrafo - SEMPRE aplicada (exceto para parágrafos especiais)
        If Not isSpecialFormatParagraph Then
            With para.Format
                .LineSpacingRule = wdLineSpacingMultiple
                .LineSpacing = LINE_SPACING
                .RightIndent = rightMarginPoints
                .SpaceBefore = 0
                .SpaceAfter = 0

                If para.alignment = wdAlignParagraphCenter Then
                    .leftIndent = 0
                    .firstLineIndent = 0
                Else
                    firstIndent = .firstLineIndent
                    paragraphIndent = .leftIndent
                    If paragraphIndent >= CentimetersToPoints(5) Then
                        .leftIndent = CentimetersToPoints(9)
                    ElseIf firstIndent < CentimetersToPoints(5) Then
                        .leftIndent = CentimetersToPoints(0)
                        .firstLineIndent = CentimetersToPoints(2.5)
                    End If
                End If
            End With

            If para.alignment = wdAlignParagraphLeft Then
                para.alignment = wdAlignParagraphJustify
            End If
        End If

        formattedCount = formattedCount + 1
    Next i

    ' Marca documento como modificado se houve formatacao
    If formattedCount > 0 Then documentDirty = True

    ' Log atualizado para refletir que todos os paragrafos sao formatados
    If skippedCount > 0 Then
        LogMessage "Paragrafos formatados: " & formattedCount & " (incluindo " & skippedCount & " com protecao de imagens)"
    End If

    ApplyStdParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação de parágrafos: " & Err.Description, LOG_LEVEL_ERROR
    ApplyStdParagraphs = False
End Function

'================================================================================
' FORMAT SECOND PARAGRAPH - FORMATAÇÃO APENAS DO 2º PARÁGRAFO
'================================================================================
Private Function FormatSecondParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim secondParaIndex As Long

    ' Identifica apenas o 2º parágrafo (considerando apenas parágrafos com texto)
    actualParaIndex = 0
    secondParaIndex = 0

    ' Cache do count para performance
    Dim paraCount As Long
    paraCount = doc.Paragraphs.count

    ' Encontra o 2º parágrafo com conteúdo (pula vazios)
    For i = 1 To paraCount
        If i > paraCount Then Exit For ' Proteção dinâmica

        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        ' Se o parágrafo tem texto ou conteúdo visual, conta como parágrafo válido
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1

            ' Registra o índice do 2º parágrafo
            If actualParaIndex = 2 Then
                secondParaIndex = i
                Exit For ' Já encontramos o 2º parágrafo
            End If
        End If

        ' Proteção expandida: processa até 20 parágrafos para encontrar o 2º
        If i > 20 Then Exit For
    Next i

    ' Aplica formatação específica apenas ao 2º parágrafo
    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(secondParaIndex)

        ' Substitui palavras iniciais conforme regras específicas
        Dim paraFullText As String
        paraFullText = para.Range.text
        paraFullText = Trim(Replace(Replace(paraFullText, vbCr, ""), vbLf, ""))

        Dim lowerStart As String
        Dim wasReplaced As Boolean
        wasReplaced = False

        ' Verifica se inicia com "Solicita" (case insensitive)
        If Len(paraFullText) >= 8 Then
            lowerStart = LCase(Left(paraFullText, 8))
            If lowerStart = "solicita" Then
                para.Range.text = "Requer" & Mid(paraFullText, 9) & vbCr
                LogMessage "Palavra inicial 'Solicita' substituída por 'Requer' no 2º parágrafo", LOG_LEVEL_INFO
                wasReplaced = True
            End If
        End If

        ' Verifica se inicia com "Pede" (case insensitive)
        If Not wasReplaced And Len(paraFullText) >= 4 Then
            lowerStart = LCase(Left(paraFullText, 4))
            If lowerStart = "pede" Then
                para.Range.text = "Requer" & Mid(paraFullText, 5) & vbCr
                LogMessage "Palavra inicial 'Pede' substituída por 'Requer' no 2º parágrafo", LOG_LEVEL_INFO
                wasReplaced = True
            End If
        End If

        ' Verifica se inicia com "Sugere" (case insensitive)
        If Not wasReplaced And Len(paraFullText) >= 6 Then
            lowerStart = LCase(Left(paraFullText, 6))
            If lowerStart = "sugere" Then
                para.Range.text = "Indica" & Mid(paraFullText, 7) & vbCr
                LogMessage "Palavra inicial 'Sugere' substituída por 'Indica' no 2º parágrafo", LOG_LEVEL_INFO
                wasReplaced = True
            End If
        End If

        ' Atualiza o texto do parágrafo se houve substituição
        If wasReplaced Then
            paraFullText = para.Range.text
        End If

        ' Remove ", neste município" se estiver no final do parágrafo
        paraFullText = para.Range.text
        paraFullText = Trim(Replace(Replace(paraFullText, vbCr, ""), vbLf, ""))

        If Len(paraFullText) > 17 Then ' Tamanho mínimo para conter ", neste município"
            Dim lowerText As String
            lowerText = LCase(paraFullText)

            ' Verifica se termina com ", neste município"
            If Right(lowerText, 17) = ", neste município" Then
                ' Remove os últimos 17 caracteres
                para.Range.text = Left(paraFullText, Len(paraFullText) - 17) & vbCr
                LogMessage "String ', neste município' removida do 2º parágrafo", LOG_LEVEL_INFO
            End If
        End If

        ' PRIMEIRO: Adiciona 2 linhas em branco ANTES do 2º parágrafo
        Dim insertionPoint As Range
        Set insertionPoint = para.Range
        insertionPoint.Collapse wdCollapseStart

        ' Verifica se já existem linhas em branco antes
        Dim blankLinesBefore As Long
        blankLinesBefore = CountBlankLinesBefore(doc, secondParaIndex)

        ' Adiciona linhas em branco conforme necessário para chegar a 2
        If blankLinesBefore < 2 Then
            Dim linesToAdd As Long
            linesToAdd = 2 - blankLinesBefore

            Dim newLines As String
            newLines = String(linesToAdd, vbCrLf)
            insertionPoint.InsertBefore newLines

            ' Atualiza o índice do segundo parágrafo (foi deslocado)
            secondParaIndex = secondParaIndex + linesToAdd
            Set para = doc.Paragraphs(secondParaIndex)
        End If

        ' FORMATAÇÃO PRINCIPAL: Aplica formatação SEMPRE, protegendo apenas as imagens
        With para.Format
            .leftIndent = CentimetersToPoints(9)      ' Recuo à esquerda de 9 cm
            .firstLineIndent = 0                      ' Sem recuo da primeira linha
            .RightIndent = 0                          ' Sem recuo à direita
            .alignment = wdAlignParagraphJustify      ' Justificado
        End With

        ' SEGUNDO: Adiciona 2 linhas em branco DEPOIS do 2º parágrafo
        Dim insertionPointAfter As Range
        Set insertionPointAfter = para.Range
        insertionPointAfter.Collapse wdCollapseEnd

        ' Verifica se já existem linhas em branco depois
        Dim blankLinesAfter As Long
        blankLinesAfter = CountBlankLinesAfter(doc, secondParaIndex)

        ' Adiciona linhas em branco conforme necessário para chegar a 2
        If blankLinesAfter < 2 Then
            Dim linesToAddAfter As Long
            linesToAddAfter = 2 - blankLinesAfter

            Dim newLinesAfter As String
            newLinesAfter = String(linesToAddAfter, vbCrLf)
            insertionPointAfter.InsertAfter newLinesAfter
        End If

        ' Se tem imagens, apenas registra (mas não pula a formatação)
        If HasVisualContent(para) Then
            LogMessage "2º parágrafo formatado com proteção de imagem e linhas em branco (posição: " & secondParaIndex & ")", LOG_LEVEL_INFO
        Else
            LogMessage "2º parágrafo formatado com 2 linhas em branco antes e depois (posição: " & secondParaIndex & ")", LOG_LEVEL_INFO
        End If
    Else
        LogMessage "2º parágrafo não encontrado para formatação", LOG_LEVEL_WARNING
    End If

    FormatSecondParagraph = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação do 2º parágrafo: " & Err.Description, LOG_LEVEL_ERROR
    FormatSecondParagraph = False
End Function

'================================================================================
' HELPER FUNCTIONS FOR BLANK LINES - Funções auxiliares para linhas em branco
'================================================================================
' Nota: CountBlankLinesBefore já está definida nas linhas 918-958
' (seção de identificação de estrutura do documento)

Private Function CountBlankLinesAfter(doc As Document, paraIndex As Long) As Long
    On Error GoTo ErrorHandler

    Dim count As Long
    Dim i As Long
    Dim para As Paragraph
    Dim paraText As String

    count = 0

    ' Verifica parágrafos posteriores (máximo 5 para performance)
    For i = paraIndex + 1 To doc.Paragraphs.count
        If i > doc.Paragraphs.count Then Exit For

        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        ' Se o parágrafo está vazio, conta como linha em branco
        If paraText = "" And Not HasVisualContent(para) Then
            count = count + 1
        Else
            ' Se encontrou parágrafo com conteúdo, para de contar
            Exit For
        End If

        ' Limite de segurança
        If count >= 5 Then Exit For
    Next i

    CountBlankLinesAfter = count
    Exit Function

ErrorHandler:
    CountBlankLinesAfter = 0
End Function

'================================================================================
' SECOND PARAGRAPH LOCATION HELPER - Localiza o segundo parágrafo
'================================================================================
Private Function GetSecondParagraphIndex(doc As Document) As Long
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long

    actualParaIndex = 0

    ' Encontra o 2º parágrafo com conteúdo (pula vazios)
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        ' Se o parágrafo tem texto ou conteúdo visual, conta como parágrafo válido
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1

            ' Retorna o índice do 2º parágrafo
            If actualParaIndex = 2 Then
                GetSecondParagraphIndex = i
                Exit Function
            End If
        End If

        ' Proteção: processa até 20 parágrafos para encontrar o 2º
        If i > 20 Then Exit For
    Next i

    GetSecondParagraphIndex = 0  ' Não encontrado
    Exit Function

ErrorHandler:
    GetSecondParagraphIndex = 0
End Function

'================================================================================
' ENSURE SECOND PARAGRAPH BLANK LINES - Garante 2 linhas em branco no 2º parágrafo
'================================================================================
Private Function EnsureSecondParagraphBlankLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim secondParaIndex As Long
    Dim linesToAdd As Long
    Dim linesToAddAfter As Long

    secondParaIndex = GetSecondParagraphIndex(doc)
    linesToAdd = 0
    linesToAddAfter = 0

    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.count Then
        Dim para As Paragraph
        Set para = doc.Paragraphs(secondParaIndex)

        ' Verifica e corrige linhas em branco ANTES
        Dim blankLinesBefore As Long
        blankLinesBefore = CountBlankLinesBefore(doc, secondParaIndex)

        If blankLinesBefore < 2 Then
            Dim insertionPoint As Range
            Set insertionPoint = para.Range
            insertionPoint.Collapse wdCollapseStart

            linesToAdd = 2 - blankLinesBefore

            Dim newLines As String
            newLines = String(linesToAdd, vbCrLf)
            insertionPoint.InsertBefore newLines

            ' Atualiza o índice (foi deslocado)
            secondParaIndex = secondParaIndex + linesToAdd
            Set para = doc.Paragraphs(secondParaIndex)
        End If

        ' Verifica e corrige linhas em branco DEPOIS
        Dim blankLinesAfter As Long
        blankLinesAfter = CountBlankLinesAfter(doc, secondParaIndex)

        If blankLinesAfter < 2 Then
            Dim insertionPointAfter As Range
            Set insertionPointAfter = para.Range
            insertionPointAfter.Collapse wdCollapseEnd

            linesToAddAfter = 2 - blankLinesAfter

            Dim newLinesAfter As String
            newLinesAfter = String(linesToAddAfter, vbCrLf)
            insertionPointAfter.InsertAfter newLinesAfter
        End If

        LogMessage "Linhas em branco do 2º parágrafo reforçadas (antes: " & (blankLinesBefore + linesToAdd) & ", depois: " & (blankLinesAfter + linesToAddAfter) & ")", LOG_LEVEL_INFO
    End If

    EnsureSecondParagraphBlankLines = True
    Exit Function

ErrorHandler:
    EnsureSecondParagraphBlankLines = False
    LogMessage "Erro ao garantir linhas em branco do 2º parágrafo: " & Err.Description, LOG_LEVEL_WARNING
End Function

'================================================================================
' ENSURE PLENARIO BLANK LINES - Garante 2 linhas em branco antes e depois do Plenário
'================================================================================
Private Function EnsurePlenarioBlankLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim paraTextLower As String
    Dim i As Long
    Dim plenarioIndex As Long

    plenarioIndex = 0

    ' Localiza o parágrafo "Plenário Dr. Tancredo Neves"
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)

        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            paraTextLower = LCase(paraText)

            ' Procura por "Plenário" e "Tancredo Neves"
            If InStr(paraTextLower, "plenário") > 0 And _
               InStr(paraTextLower, "tancredo") > 0 And _
               InStr(paraTextLower, "neves") > 0 Then
                plenarioIndex = i
                Exit For
            End If
        End If
    Next i

    If plenarioIndex > 0 Then
        ' Remove linhas vazias ANTES
        i = plenarioIndex - 1
        Do While i >= 1
            Set para = doc.Paragraphs(i)
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

            If paraText = "" And Not HasVisualContent(para) Then
                para.Range.Delete
                plenarioIndex = plenarioIndex - 1
                i = i - 1
            Else
                Exit Do
            End If
        Loop

        ' Remove linhas vazias DEPOIS
        i = plenarioIndex + 1
        Do While i <= doc.Paragraphs.count
            Set para = doc.Paragraphs(i)
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

            If paraText = "" And Not HasVisualContent(para) Then
                para.Range.Delete
            Else
                Exit Do
            End If
        Loop

        ' Insere EXATAMENTE 2 linhas em branco ANTES
        Set para = doc.Paragraphs(plenarioIndex)
        para.Range.InsertParagraphBefore
        para.Range.InsertParagraphBefore

        ' Insere EXATAMENTE 2 linhas em branco DEPOIS
        Set para = doc.Paragraphs(plenarioIndex + 2) ' +2 porque inserimos 2 antes
        para.Range.InsertParagraphAfter
        para.Range.InsertParagraphAfter

        LogMessage "Linhas em branco do Plenário reforçadas: 2 antes e 2 depois", LOG_LEVEL_INFO
    End If

    EnsurePlenarioBlankLines = True
    Exit Function

ErrorHandler:
    EnsurePlenarioBlankLines = False
    LogMessage "Erro ao garantir linhas em branco do Plenário: " & Err.Description, LOG_LEVEL_WARNING
End Function

'================================================================================
' ENSURE SINGLE BLANK LINE BETWEEN PARAGRAPHS - Garante pelo menos 1 linha em branco entre parágrafos
'================================================================================
Private Function EnsureSingleBlankLineBetweenParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim para As Paragraph
    Dim NextPara As Paragraph
    Dim paraText As String
    Dim nextParaText As String
    Dim insertionPoint As Range
    Dim addedCount As Long

    addedCount = 0

    ' Percorre todos os parágrafos de trás para frente para não afetar os índices
    For i = doc.Paragraphs.count - 1 To 1 Step -1
        Set para = doc.Paragraphs(i)
        Set NextPara = doc.Paragraphs(i + 1)

        ' Obtém texto limpo dos parágrafos
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        nextParaText = Trim(Replace(Replace(NextPara.Range.text, vbCr, ""), vbLf, ""))

        ' Se ambos os parágrafos têm conteúdo (texto ou imagem)
        If (paraText <> "" Or HasVisualContent(para)) And _
           (nextParaText <> "" Or HasVisualContent(NextPara)) Then

            ' Verifica se há pelo menos uma linha em branco entre eles
            Dim hasBlankBetween As Boolean
            hasBlankBetween = False

            ' Verifica se o próximo parágrafo é imediatamente adjacente
            ' Isso seria indicado se não há parágrafo vazio entre eles
            If i + 1 <= doc.Paragraphs.count Then
                ' Se o índice do próximo parágrafo é i+1, eles são adjacentes
                ' e precisamos verificar se há linha em branco
                Dim checkIndex As Long
                For checkIndex = i + 1 To i + 1
                    If checkIndex <= doc.Paragraphs.count Then
                        Dim checkPara As Paragraph
                        Set checkPara = doc.Paragraphs(checkIndex)
                        Dim checkText As String
                        checkText = Trim(Replace(Replace(checkPara.Range.text, vbCr, ""), vbLf, ""))

                        ' Se o parágrafo entre eles está vazio, há linha em branco
                        If checkText = "" And Not HasVisualContent(checkPara) Then
                            hasBlankBetween = True
                        End If
                    End If
                Next checkIndex
            End If

            ' Se não há linha em branco, adiciona uma
            If Not hasBlankBetween Then
                Set insertionPoint = NextPara.Range
                insertionPoint.Collapse wdCollapseStart
                insertionPoint.InsertBefore vbCrLf
                addedCount = addedCount + 1
            End If
        End If
    Next i

    If addedCount > 0 Then
        LogMessage "Linhas em branco adicionadas entre parágrafos: " & addedCount, LOG_LEVEL_INFO
    End If

    EnsureSingleBlankLineBetweenParagraphs = True
    Exit Function

ErrorHandler:
    EnsureSingleBlankLineBetweenParagraphs = False
    LogMessage "Erro ao garantir linhas em branco entre parágrafos: " & Err.Description, LOG_LEVEL_WARNING
End Function

'================================================================================
' ENSURE BLANK LINES BETWEEN LONG PARAGRAPHS - Garante linha em branco entre parágrafos com mais de 10 palavras
'================================================================================
Private Function EnsureBlankLinesBetweenLongParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim para As Paragraph
    Dim NextPara As Paragraph
    Dim paraText As String
    Dim nextParaText As String
    Dim paraWordCount As Long
    Dim nextParaWordCount As Long
    Dim insertionPoint As Range
    Dim addedCount As Long

    addedCount = 0

    ' Percorre todos os parágrafos de trás para frente para não afetar os índices
    For i = doc.Paragraphs.count - 1 To 1 Step -1
        If i >= doc.Paragraphs.count Then Exit For ' Proteção dinâmica

        Set para = doc.Paragraphs(i)

        ' Verifica se há próximo parágrafo
        If i + 1 <= doc.Paragraphs.count Then
            Set NextPara = doc.Paragraphs(i + 1)
        Else
            Exit For
        End If

        ' Obtém texto limpo dos parágrafos
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        nextParaText = Trim(Replace(Replace(NextPara.Range.text, vbCr, ""), vbLf, ""))

        ' Conta palavras (divide por espaços)
        paraWordCount = 0
        nextParaWordCount = 0

        If paraText <> "" Then
            paraWordCount = UBound(Split(paraText, " ")) + 1
        End If

        If nextParaText <> "" Then
            nextParaWordCount = UBound(Split(nextParaText, " ")) + 1
        End If

        ' Se ambos os parágrafos têm mais de 10 palavras
        If paraWordCount > 10 And nextParaWordCount > 10 Then
            ' Verifica se há linha em branco entre eles
            Dim hasBlankBetween As Boolean
            hasBlankBetween = False

            ' Verifica se eles são adjacentes (sem linha em branco entre)
            ' Se i+1 é o próximo parágrafo e não está vazio, são adjacentes
            If nextParaText <> "" Then
                hasBlankBetween = False
            Else
                hasBlankBetween = True
            End If

            ' Se não há linha em branco, adiciona uma
            If Not hasBlankBetween Then
                Set insertionPoint = NextPara.Range
                insertionPoint.Collapse wdCollapseStart
                insertionPoint.InsertBefore vbCrLf
                addedCount = addedCount + 1
            End If
        End If
    Next i

    If addedCount > 0 Then
        LogMessage "Linhas em branco adicionadas entre parágrafos longos (>10 palavras): " & addedCount, LOG_LEVEL_INFO
    End If

    EnsureBlankLinesBetweenLongParagraphs = True
    Exit Function

ErrorHandler:
    EnsureBlankLinesBetweenLongParagraphs = False
    LogMessage "Erro ao garantir linhas em branco entre parágrafos longos: " & Err.Description, LOG_LEVEL_WARNING
End Function

'================================================================================
' FORMATAÇÃO DO PRIMEIRO PARÁGRAFO
'================================================================================
Private Function FormatFirstParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim firstParaIndex As Long

    ' Identifica o 1º parágrafo (considerando apenas parágrafos com texto)
    actualParaIndex = 0
    firstParaIndex = 0

    ' Encontra o 1º parágrafo com conteúdo (pula vazios)
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        ' Se o parágrafo tem texto ou conteúdo visual, conta como parágrafo válido
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1

            ' Registra o índice do 1º parágrafo
            If actualParaIndex = 1 Then
                firstParaIndex = i
                Exit For ' Já encontramos o 1º parágrafo
            End If
        End If

        ' Proteção expandida: processa até 20 parágrafos para encontrar o 1º
        If i > 20 Then Exit For
    Next i

    ' Aplica formatação específica apenas ao 1º parágrafo
    If firstParaIndex > 0 And firstParaIndex <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(firstParaIndex)

        ' NOVO: Aplica formatação SEMPRE, protegendo apenas as imagens
        ' Formatação do 1º parágrafo: caixa alta, negrito e sublinhado
        If HasVisualContent(para) Then
            ' Para parágrafos com imagens, aplica formatação caractere por caractere
            Dim n As Long
            Dim charCount4 As Long
            charCount4 = SafeGetCharacterCount(para.Range) ' Cache da contagem segura

            If charCount4 > 0 Then ' Verificação de segurança
                For n = 1 To charCount4
                    Dim charRange3 As Range
                    Set charRange3 = para.Range.Characters(n)
                    If charRange3.InlineShapes.count = 0 Then
                        With charRange3.Font
                            .AllCaps = True           ' Caixa alta (maiúsculas)
                            .Bold = True              ' Negrito
                            .Underline = wdUnderlineSingle ' Sublinhado
                        End With
                    End If
                Next n
            End If
            LogMessage "1º parágrafo formatado com proteção de imagem (posição: " & firstParaIndex & ")"
        Else
            ' Formatação normal para parágrafos sem imagens
            With para.Range.Font
                .AllCaps = True           ' Caixa alta (maiúsculas)
                .Bold = True              ' Negrito
                .Underline = wdUnderlineSingle ' Sublinhado
            End With
        End If

        ' Aplicar também formatação de parágrafo - SEMPRE
        With para.Format
            .alignment = wdAlignParagraphCenter       ' Centralizado
            .leftIndent = 0                           ' Sem recuo à esquerda
            .firstLineIndent = 0                      ' Sem recuo da primeira linha
            .RightIndent = 0                          ' Sem recuo à direita
        End With
    Else
        LogMessage "1º parágrafo não encontrado para formatação", LOG_LEVEL_WARNING
    End If

    FormatFirstParagraph = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação do 1º parágrafo: " & Err.Description, LOG_LEVEL_ERROR
    FormatFirstParagraph = False
End Function

'================================================================================
' REMOÇÃO DE MARCA D'ÁGUA
'================================================================================
Private Function RemoveWatermark(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim shp As shape
    Dim i As Long
    Dim removedCount As Long

    For Each sec In doc.Sections
        For Each header In sec.Headers
            If header.Exists And header.Shapes.count > 0 Then
                For i = header.Shapes.count To 1 Step -1
                    Set shp = header.Shapes(i)
                    If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
                        If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Or _
                           InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then
                            shp.Delete
                            removedCount = removedCount + 1
                        End If
                    End If
                Next i
            End If
        Next header

        For Each header In sec.Footers
            If header.Exists And header.Shapes.count > 0 Then
                For i = header.Shapes.count To 1 Step -1
                    Set shp = header.Shapes(i)
                    If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
                        If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Or _
                           InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then
                            shp.Delete
                            removedCount = removedCount + 1
                        End If
                    End If
                Next i
            End If
        Next header
    Next sec

    If removedCount > 0 Then
        LogMessage "Marcas d'água removidas: " & removedCount & " itens"
    End If
    ' Log de "nenhuma marca d'água" removido para performance

    RemoveWatermark = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover marcas d'água: " & Err.Description, LOG_LEVEL_ERROR
    RemoveWatermark = False
End Function

'================================================================================
' GERENCIAMENTO DE CAMINHO DA IMAGEM DE CABEÇALHO
'================================================================================
Private Function GetHeaderImagePath() As String
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim shell As Object
    Dim userProfilePath As String
    Dim headerImagePath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")

    ' Obtém pasta %USERPROFILE% do usuário atual (compatível com Windows)
    userProfilePath = shell.ExpandEnvironmentStrings("%USERPROFILE%")
    If Right(userProfilePath, 1) = "\" Then
        userProfilePath = Left(userProfilePath, Len(userProfilePath) - 1)
    End If

    ' Constrói caminho absoluto para a imagem desejada
    headerImagePath = userProfilePath & "\chainsaw\assets\stamp.png"

    ' Verifica se o arquivo existe
    If Not fso.FileExists(headerImagePath) Then
        LogMessage "Imagem de cabeçalho não encontrada em: " & headerImagePath, LOG_LEVEL_WARNING
        GetHeaderImagePath = ""
        Exit Function
    End If

    GetHeaderImagePath = headerImagePath
    Exit Function

ErrorHandler:
    LogMessage "Erro ao localizar imagem de cabeçalho: " & Err.Description, LOG_LEVEL_ERROR
    GetHeaderImagePath = ""
End Function

'================================================================================
' INSERÇÃO DE IMAGEM DE CABEÇALHO
'================================================================================
Private Function InsertHeaderstamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim imgFile As String
    Dim username As String
    Dim imgWidth As Single
    Dim imgHeight As Single
    Dim shp As shape
    Dim imgFound As Boolean
    Dim sectionsProcessed As Long

    ' Define o caminho da imagem do cabeçalho
    imgFile = Environ("USERPROFILE") & "\chainsaw\assets\stamp.png"

    If Dir(imgFile) = "" Then
        Application.StatusBar = "Aviso: Imagem não encontrada"
        LogMessage "Header image not found at: " & imgFile, LOG_LEVEL_WARNING
        InsertHeaderstamp = False
        Exit Function
    End If

    ' Size using standard constants
    imgWidth = CentimetersToPoints(HEADER_IMAGE_MAX_WIDTH_CM)
    imgHeight = imgWidth * HEADER_IMAGE_HEIGHT_RATIO

    For Each sec In doc.Sections
        Set header = sec.Headers(wdHeaderFooterPrimary)
        If header.Exists Then
            header.LinkToPrevious = False
            header.Range.Delete

            ' Define fonte padrão para o cabeçalho: Arial 12
            With header.Range.Font
                .Name = STANDARD_FONT  ' Arial
                .size = STANDARD_FONT_SIZE  ' 12
            End With

            Set shp = header.Shapes.AddPicture( _
                fileName:=imgFile, _
                LinkToFile:=False, _
                SaveWithDocument:=msoTrue)

            If shp Is Nothing Then
                LogMessage "Failed to insert header image at section " & sectionsProcessed + 1, LOG_LEVEL_WARNING
            Else
                With shp
                    .LockAspectRatio = msoTrue
                    .Width = imgWidth
                    .Height = imgHeight
                    .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                    .RelativeVerticalPosition = wdRelativeVerticalPositionPage
                    .Left = (doc.PageSetup.PageWidth - .Width) / 2
                    .Top = CentimetersToPoints(HEADER_IMAGE_TOP_MARGIN_CM)
                    .WrapFormat.Type = wdWrapTopBottom
                    .ZOrder msoSendToBack
                End With

                imgFound = True
                sectionsProcessed = sectionsProcessed + 1
            End If
        End If
    Next sec

    If imgFound Then
        ' Log detalhado removido para performance
        InsertHeaderstamp = True
    Else
    LogMessage "No header was inserted", LOG_LEVEL_WARNING
        InsertHeaderstamp = False
    End If

    Exit Function

ErrorHandler:
    LogMessage "Error inserting header: " & Err.Description, LOG_LEVEL_ERROR
    InsertHeaderstamp = False
End Function

'================================================================================
' INSERÇÃO DE NÚMEROS DE PÁGINA NO RODAPÉ
'================================================================================
Private Function InsertFooterStamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim footer As HeaderFooter
    Dim rng As Range
    Dim sectionsProcessed As Long

    For Each sec In doc.Sections
        Set footer = sec.Footers(wdHeaderFooterPrimary)

        If footer.Exists Then
            footer.LinkToPrevious = False
            Set rng = footer.Range

            rng.Delete

            Set rng = footer.Range
            rng.Collapse Direction:=wdCollapseEnd
            rng.Fields.Add Range:=rng, Type:=wdFieldPage

            Set rng = footer.Range
            rng.Collapse Direction:=wdCollapseEnd
            rng.text = "-"

            Set rng = footer.Range
            rng.Collapse Direction:=wdCollapseEnd
            rng.Fields.Add Range:=rng, Type:=wdFieldNumPages

            With footer.Range
                .Font.Name = STANDARD_FONT
                .Font.size = FOOTER_FONT_SIZE
                .ParagraphFormat.alignment = wdAlignParagraphCenter
                .Fields.Update
            End With

            sectionsProcessed = sectionsProcessed + 1
        End If
    Next sec

    ' Log detalhado removido para performance
    InsertFooterStamp = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao inserir rodapé: " & Err.Description, LOG_LEVEL_ERROR
    InsertFooterStamp = False
End Function

'================================================================================
' UTILITY: CM TO POINTS
'================================================================================
Private Function CentimetersToPoints(ByVal cm As Double) As Single
    On Error Resume Next
    CentimetersToPoints = Application.CentimetersToPoints(cm)
    If Err.Number <> 0 Then
        CentimetersToPoints = cm * 28.35
    End If
End Function

'================================================================================
' UTILITY: SAFE USERNAME
'================================================================================
Private Function GetSafeUserName() As String
    On Error GoTo ErrorHandler

    Dim rawName As String
    Dim safeName As String
    Dim i As Integer
    Dim c As String

    rawName = Environ("USERNAME")
    If rawName = "" Then rawName = Environ("USER")
    If rawName = "" Then
        On Error Resume Next
        rawName = CreateObject("WScript.Network").username
        On Error GoTo 0
    End If

    If rawName = "" Then
        rawName = "UsuarioDesconhecido"
    End If

    For i = 1 To Len(rawName)
        c = Mid(rawName, i, 1)
        If c Like "[A-Za-z0-9_\-]" Then
            safeName = safeName & c
        ElseIf c = " " Then
            safeName = safeName & "_"
        End If
    Next i

    If safeName = "" Then safeName = "Usuario"

    GetSafeUserName = safeName
    Exit Function

ErrorHandler:
    GetSafeUserName = "Usuario"
End Function

'================================================================================
' VALIDACAO DO TIPO DE PROPOSITURA
'================================================================================
' Verifica se a primeira palavra do documento e um tipo valido de propositura
' Tipos validos: indicacao, requerimento, mocao (com tolerancia a erros de grafia)
'================================================================================
Private Function ValidateProposituraType(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ValidateProposituraType = True

    ' Obtem a primeira palavra do documento
    Dim firstWord As String
    firstWord = GetFirstWord(doc)

    If Len(firstWord) = 0 Then
        LogMessage "Documento vazio ou sem texto no inicio", LOG_LEVEL_WARNING
        Exit Function
    End If

    ' Converte para minusculas e remove acentos para comparacao
    Dim normalizedWord As String
    normalizedWord = NormalizeForComparison(firstWord)

    ' Verifica se corresponde a um tipo valido (com tolerancia a erros)
    If IsValidProposituraWord(normalizedWord) Then
        LogMessage "Tipo de propositura identificado: " & firstWord, LOG_LEVEL_INFO
        ValidateProposituraType = True
        Exit Function
    End If

    ' Nao e um tipo reconhecido - pergunta ao usuario
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("A primeira palavra do titulo e: """ & firstWord & """" & vbCrLf & vbCrLf & _
                          "Nao parece ser uma propositura de Indicacao, Requerimento ou Mocao," & vbCrLf & _
                          "ou ha algum erro de grafia na primeira palavra do titulo." & vbCrLf & vbCrLf & _
                          "Deseja prosseguir com o processamento mesmo assim?", _
                          vbYesNo + vbQuestion, "CHAINSAW - Tipo de Propositura")

    If userResponse = vbYes Then
        LogMessage "Usuario optou por prosseguir com tipo nao reconhecido: " & firstWord, LOG_LEVEL_WARNING
        ValidateProposituraType = True
    Else
        LogMessage "Usuario cancelou - tipo nao reconhecido: " & firstWord, LOG_LEVEL_INFO
        ValidateProposituraType = False
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro ao validar tipo de propositura: " & Err.Description, LOG_LEVEL_WARNING
    ValidateProposituraType = True ' Em caso de erro, permite prosseguir
End Function

'================================================================================
' OBTEM A PRIMEIRA PALAVRA DO DOCUMENTO
'================================================================================
Private Function GetFirstWord(doc As Document) As String
    On Error GoTo ErrorHandler

    GetFirstWord = ""

    ' Percorre os primeiros paragrafos ate encontrar texto
    Dim i As Long
    Dim paraText As String

    For i = 1 To doc.Paragraphs.count
        If i > 10 Then Exit For ' Limite de seguranca

        paraText = Trim(Replace(Replace(doc.Paragraphs(i).Range.text, vbCr, ""), vbLf, ""))

        If Len(paraText) > 0 Then
            ' Extrai a primeira palavra (ate o primeiro espaco)
            Dim spacePos As Long
            spacePos = InStr(paraText, " ")

            If spacePos > 0 Then
                GetFirstWord = Left(paraText, spacePos - 1)
            Else
                GetFirstWord = paraText
            End If

            Exit For
        End If
    Next i

    Exit Function

ErrorHandler:
    GetFirstWord = ""
End Function

'================================================================================
' NORMALIZA TEXTO PARA COMPARACAO (remove acentos e converte para minusculas)
'================================================================================
Private Function NormalizeForComparison(text As String) As String
    Dim result As String
    result = LCase(text)

    ' Remove acentos comuns do portugues
    result = Replace(result, Chr(225), "a") ' a com acento agudo
    result = Replace(result, Chr(227), "a") ' a com til
    result = Replace(result, Chr(226), "a") ' a com circunflexo
    result = Replace(result, Chr(224), "a") ' a com acento grave
    result = Replace(result, Chr(233), "e") ' e com acento agudo
    result = Replace(result, Chr(234), "e") ' e com circunflexo
    result = Replace(result, Chr(237), "i") ' i com acento agudo
    result = Replace(result, Chr(243), "o") ' o com acento agudo
    result = Replace(result, Chr(245), "o") ' o com til
    result = Replace(result, Chr(244), "o") ' o com circunflexo
    result = Replace(result, Chr(250), "u") ' u com acento agudo
    result = Replace(result, Chr(231), "c") ' c cedilha

    NormalizeForComparison = result
End Function

'================================================================================
' VERIFICA SE A PALAVRA E UM TIPO VALIDO DE PROPOSITURA
'================================================================================
Private Function IsValidProposituraWord(normalizedWord As String) As Boolean
    IsValidProposituraWord = False

    ' Padroes validos (normalizados, sem acentos)
    ' indicacao, requerimento, mocao

    ' Verifica correspondencia exata primeiro
    If normalizedWord = "indicacao" Or _
       normalizedWord = "requerimento" Or _
       normalizedWord = "mocao" Then
        IsValidProposituraWord = True
        Exit Function
    End If

    ' Verifica com tolerancia a pequenos erros (distancia de Levenshtein <= 2)
    If LevenshteinDistance(normalizedWord, "indicacao") <= 2 Then
        IsValidProposituraWord = True
        Exit Function
    End If

    If LevenshteinDistance(normalizedWord, "requerimento") <= 2 Then
        IsValidProposituraWord = True
        Exit Function
    End If

    If LevenshteinDistance(normalizedWord, "mocao") <= 2 Then
        IsValidProposituraWord = True
        Exit Function
    End If
End Function

'================================================================================
' CALCULA A DISTANCIA DE LEVENSHTEIN ENTRE DUAS STRINGS
'================================================================================
Private Function LevenshteinDistance(s1 As String, s2 As String) As Long
    Dim len1 As Long, len2 As Long
    Dim i As Long, j As Long
    Dim cost As Long
    Dim d() As Long

    len1 = Len(s1)
    len2 = Len(s2)

    ' Casos triviais
    If len1 = 0 Then
        LevenshteinDistance = len2
        Exit Function
    End If

    If len2 = 0 Then
        LevenshteinDistance = len1
        Exit Function
    End If

    ' Matriz de distancias
    ReDim d(0 To len1, 0 To len2)

    ' Inicializa primeira coluna e linha
    For i = 0 To len1
        d(i, 0) = i
    Next i

    For j = 0 To len2
        d(0, j) = j
    Next j

    ' Calcula distancias
    For i = 1 To len1
        For j = 1 To len2
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If

            ' Minimo entre insercao, delecao e substituicao
            d(i, j) = MinOfThree(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cost)
        Next j
    Next i

    LevenshteinDistance = d(len1, len2)
End Function

'================================================================================
' RETORNA O MINIMO DE TRES VALORES
'================================================================================
Private Function MinOfThree(a As Long, b As Long, c As Long) As Long
    MinOfThree = a
    If b < MinOfThree Then MinOfThree = b
    If c < MinOfThree Then MinOfThree = c
End Function

'================================================================================
' VALIDACAO DE ESTRUTURA DO DOCUMENTO
'================================================================================
Private Function ValidateDocumentStructure(doc As Document) As Boolean
    On Error Resume Next

    ' Verificacao basica e rapida
    If doc.Range.End > 0 And doc.Sections.count > 0 Then
        ValidateDocumentStructure = True
    Else
        LogMessage "Documento com estrutura inconsistente", LOG_LEVEL_WARNING
        ValidateDocumentStructure = False
    End If
End Function

'================================================================================
' VALIDACAO DE CONSISTENCIA EMENTA x PROPOSICAO
' Compara elementos-chave entre ementa e texto da proposicao
'================================================================================
Private Function ValidateAddressConsistency(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim ementaText As String
    Dim proposicaoText As String
    Dim inconsistencies As String
    Dim inconsistencyCount As Long

    ' Obtem textos da ementa e proposicao usando o cache de estrutura
    ementaText = GetEmentaText(doc)
    proposicaoText = GetProposicaoText(doc)

    ' Se nao conseguiu identificar ementa ou proposicao, retorna True (sem verificacao)
    If Len(ementaText) < 10 Or Len(proposicaoText) < 20 Then
        LogMessage "Ementa ou proposicao nao identificadas para validacao", LOG_LEVEL_INFO
        ValidateAddressConsistency = True
        Exit Function
    End If

    inconsistencies = ""
    inconsistencyCount = 0

    ' 1. Verifica enderecos (Rua, Avenida, etc)
    Dim addressInconsistency As String
    addressInconsistency = CheckAddressConsistency(ementaText, proposicaoText)
    If Len(addressInconsistency) > 0 Then
        inconsistencies = inconsistencies & addressInconsistency & vbCrLf
        inconsistencyCount = inconsistencyCount + 1
    End If

    ' 2. Verifica valores monetarios (R$)
    Dim monetaryInconsistency As String
    monetaryInconsistency = CheckMonetaryConsistency(ementaText, proposicaoText)
    If Len(monetaryInconsistency) > 0 Then
        inconsistencies = inconsistencies & monetaryInconsistency & vbCrLf
        inconsistencyCount = inconsistencyCount + 1
    End If

    ' 3. Verifica numeros de referencia (n., n.o, numero)
    Dim numberInconsistency As String
    numberInconsistency = CheckNumberConsistency(ementaText, proposicaoText)
    If Len(numberInconsistency) > 0 Then
        inconsistencies = inconsistencies & numberInconsistency & vbCrLf
        inconsistencyCount = inconsistencyCount + 1
    End If

    ' 4. Verifica bairros mencionados
    Dim bairroInconsistency As String
    bairroInconsistency = CheckBairroConsistency(ementaText, proposicaoText)
    If Len(bairroInconsistency) > 0 Then
        inconsistencies = inconsistencies & bairroInconsistency & vbCrLf
        inconsistencyCount = inconsistencyCount + 1
    End If

    ' Se encontrou inconsistencias, exibe mensagem consolidada
    If inconsistencyCount > 0 Then
        Dim msg As String
        msg = "VERIFICAR CONSISTENCIA" & vbCrLf & vbCrLf
        msg = msg & "Foram encontradas " & inconsistencyCount & " possivel(is) inconsistencia(s) entre a ementa e o texto:" & vbCrLf & vbCrLf
        msg = msg & inconsistencies & vbCrLf
        msg = msg & "Recomenda-se revisar o documento antes de prosseguir."

        MsgBox msg, vbExclamation, "Verificacao de Consistencia"

        LogMessage "Inconsistencias detectadas: " & inconsistencyCount & " item(ns)", LOG_LEVEL_WARNING
        ValidateAddressConsistency = False
        Exit Function
    End If

    ' Tudo OK
    LogMessage "Validacao ementa x proposicao: sem inconsistencias detectadas", LOG_LEVEL_INFO
    ValidateAddressConsistency = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao validar consistencia ementa/proposicao: " & Err.Description, LOG_LEVEL_WARNING
    ValidateAddressConsistency = True
End Function

'================================================================================
' OBTEM TEXTO DA EMENTA
'================================================================================
Private Function GetEmentaText(doc As Document) As String
    On Error Resume Next
    GetEmentaText = ""

    If doc Is Nothing Then Exit Function

    Dim i As Long
    Dim para As Paragraph
    Dim paraText As String
    Dim textualCount As Long

    textualCount = 0

    ' Percorre paragrafos buscando o 2o paragrafo textual (ementa)
    For i = 1 To doc.Paragraphs.count
        If i > 20 Then Exit For ' Limite de seguranca

        Set para = doc.Paragraphs(i)
        paraText = Trim(para.Range.text)

        ' Ignora paragrafos vazios
        If Len(paraText) > 1 Then
            textualCount = textualCount + 1

            ' 2o paragrafo textual e a ementa
            If textualCount = 2 Then
                ' Verifica se tem recuo (caracteristica da ementa)
                If para.Format.leftIndent > 5 Then
                    GetEmentaText = paraText
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

'================================================================================
' OBTEM TEXTO DA PROPOSICAO (CORPO DO DOCUMENTO)
'================================================================================
Private Function GetProposicaoText(doc As Document) As String
    On Error Resume Next
    GetProposicaoText = ""

    If doc Is Nothing Then Exit Function

    Dim i As Long
    Dim para As Paragraph
    Dim paraText As String
    Dim textualCount As Long
    Dim result As String
    Dim paraCount As Long

    textualCount = 0
    result = ""
    paraCount = 0

    ' Percorre paragrafos coletando texto apos a ementa
    For i = 1 To doc.Paragraphs.count
        If i > 50 Then Exit For ' Limite de seguranca

        Set para = doc.Paragraphs(i)
        paraText = Trim(para.Range.text)

        ' Ignora paragrafos vazios
        If Len(paraText) > 1 Then
            textualCount = textualCount + 1

            ' Coleta paragrafos do 4o em diante (apos titulo, ementa, data)
            If textualCount >= 4 Then
                ' Para ao encontrar "Justificativa" ou assinatura
                Dim lowerText As String
                lowerText = LCase(paraText)
                If InStr(lowerText, "justificativa") > 0 Then Exit For
                If InStr(lowerText, "vereador") > 0 Then Exit For

                result = result & " " & paraText
                paraCount = paraCount + 1

                ' Limita a 10 paragrafos para nao sobrecarregar
                If paraCount >= 10 Then Exit For
            End If
        End If
    Next i

    GetProposicaoText = result
End Function

'================================================================================
' VERIFICA CONSISTENCIA DE ENDERECOS
'================================================================================
Private Function CheckAddressConsistency(ementaText As String, proposicaoText As String) As String
    On Error Resume Next
    CheckAddressConsistency = ""

    Dim addressKeywords() As Variant
    addressKeywords = Array("rua ", "avenida ", "av. ", "travessa ", "alameda ", "praca ", "estrada ")

    Dim kw As Variant
    Dim kwPos As Long
    Dim addressInEmenta As String
    Dim foundInProposicao As Boolean

    For Each kw In addressKeywords
        kwPos = InStr(1, LCase(ementaText), CStr(kw), vbTextCompare)

        If kwPos > 0 Then
            ' Extrai endereco da ementa (ate 50 caracteres apos a palavra-chave)
            addressInEmenta = ExtractAddressWords(ementaText, kwPos, CStr(kw))

            If Len(addressInEmenta) > 3 Then
                ' Verifica se endereco existe na proposicao
                foundInProposicao = CheckAddressInText(addressInEmenta, proposicaoText)

                If Not foundInProposicao Then
                    CheckAddressConsistency = "ENDERECO: '" & UCase(kw) & addressInEmenta & "' da ementa nao encontrado no texto."
                    Exit Function
                End If
            End If
        End If
    Next kw
End Function

'================================================================================
' EXTRAI PALAVRAS DO ENDERECO
'================================================================================
Private Function ExtractAddressWords(text As String, startPos As Long, keyword As String) As String
    On Error Resume Next
    ExtractAddressWords = ""

    Dim afterKeyword As String
    Dim words() As String
    Dim result As String
    Dim i As Long

    ' Pega texto apos a palavra-chave
    afterKeyword = Mid(text, startPos + Len(keyword), 60)
    afterKeyword = CleanTextForComparison(afterKeyword)

    ' Divide em palavras
    words = Split(afterKeyword, " ")

    result = ""
    For i = 0 To UBound(words)
        If i > 2 Then Exit For ' Maximo 3 palavras

        Dim word As String
        word = Trim(words(i))

        ' Ignora artigos e preposicoes
        If Len(word) > 2 Then
            If result <> "" Then result = result & " "
            result = result & word
        End If
    Next i

    ExtractAddressWords = result
End Function

'================================================================================
' VERIFICA SE ENDERECO EXISTE NO TEXTO
'================================================================================
Private Function CheckAddressInText(address As String, text As String) As Boolean
    On Error Resume Next
    CheckAddressInText = False

    Dim normalizedAddress As String
    Dim normalizedText As String
    Dim words() As String
    Dim word As Variant
    Dim foundCount As Long
    Dim totalWords As Long

    normalizedAddress = CleanTextForComparison(address)
    normalizedText = CleanTextForComparison(text)

    words = Split(normalizedAddress, " ")
    foundCount = 0
    totalWords = 0

    For Each word In words
        If Len(Trim(CStr(word))) > 2 Then
            totalWords = totalWords + 1
            If InStr(1, normalizedText, CStr(word), vbTextCompare) > 0 Then
                foundCount = foundCount + 1
            End If
        End If
    Next word

    ' Considera consistente se encontrou pelo menos 70% das palavras
    If totalWords > 0 Then
        CheckAddressInText = (foundCount / totalWords) >= 0.7
    End If
End Function

'================================================================================
' VERIFICA CONSISTENCIA DE VALORES MONETARIOS
'================================================================================
Private Function CheckMonetaryConsistency(ementaText As String, proposicaoText As String) As String
    On Error Resume Next
    CheckMonetaryConsistency = ""

    Dim rsPos As Long
    Dim valueInEmenta As String
    Dim normalizedProposicao As String

    ' Procura por R$ na ementa
    rsPos = InStr(1, ementaText, "R$", vbTextCompare)

    If rsPos > 0 Then
        ' Extrai valor (R$ seguido de numeros)
        valueInEmenta = ExtractMonetaryValue(ementaText, rsPos)

        If Len(valueInEmenta) > 0 Then
            ' Normaliza proposicao para comparacao
            normalizedProposicao = CleanTextForComparison(proposicaoText)
            normalizedProposicao = Replace(normalizedProposicao, ".", "")
            normalizedProposicao = Replace(normalizedProposicao, ",", "")

            ' Remove pontuacao do valor para comparacao
            Dim normalizedValue As String
            normalizedValue = Replace(valueInEmenta, ".", "")
            normalizedValue = Replace(normalizedValue, ",", "")
            normalizedValue = Replace(normalizedValue, " ", "")

            ' Verifica se valor numerico existe na proposicao
            If InStr(1, normalizedProposicao, normalizedValue, vbTextCompare) = 0 Then
                CheckMonetaryConsistency = "VALOR: 'R$ " & valueInEmenta & "' da ementa nao encontrado no texto."
            End If
        End If
    End If
End Function

'================================================================================
' EXTRAI VALOR MONETARIO
'================================================================================
Private Function ExtractMonetaryValue(text As String, rsPos As Long) As String
    On Error Resume Next
    ExtractMonetaryValue = ""

    Dim afterRS As String
    Dim i As Long
    Dim c As String
    Dim result As String
    Dim foundDigit As Boolean

    afterRS = Mid(text, rsPos + 2, 30) ' Pega ate 30 caracteres apos R$
    afterRS = Trim(afterRS)

    result = ""
    foundDigit = False

    For i = 1 To Len(afterRS)
        c = Mid(afterRS, i, 1)

        ' Aceita digitos, ponto, virgula e espaco
        If c Like "[0-9]" Then
            result = result & c
            foundDigit = True
        ElseIf (c = "." Or c = "," Or c = " ") And foundDigit Then
            result = result & c
        ElseIf foundDigit Then
            Exit For ' Terminou o numero
        End If
    Next i

    ExtractMonetaryValue = Trim(result)
End Function

'================================================================================
' VERIFICA CONSISTENCIA DE NUMEROS DE REFERENCIA
'================================================================================
Private Function CheckNumberConsistency(ementaText As String, proposicaoText As String) As String
    On Error Resume Next
    CheckNumberConsistency = ""

    Dim numberPatterns() As Variant
    numberPatterns = Array("n. ", "n.o ", "no ", "numero ")

    Dim pattern As Variant
    Dim patternPos As Long
    Dim numberInEmenta As String
    Dim normalizedProposicao As String

    normalizedProposicao = CleanTextForComparison(proposicaoText)

    For Each pattern In numberPatterns
        patternPos = InStr(1, LCase(ementaText), CStr(pattern), vbTextCompare)

        If patternPos > 0 Then
            ' Extrai numero apos o padrao
            numberInEmenta = ExtractReferenceNumber(ementaText, patternPos + Len(pattern))

            If Len(numberInEmenta) > 0 Then
                ' Verifica se numero existe na proposicao
                If InStr(1, normalizedProposicao, numberInEmenta, vbTextCompare) = 0 Then
                    CheckNumberConsistency = "NUMERO: '" & numberInEmenta & "' da ementa nao encontrado no texto."
                    Exit Function
                End If
            End If
        End If
    Next pattern
End Function

'================================================================================
' EXTRAI NUMERO DE REFERENCIA
'================================================================================
Private Function ExtractReferenceNumber(text As String, startPos As Long) As String
    On Error Resume Next
    ExtractReferenceNumber = ""

    Dim afterPattern As String
    Dim i As Long
    Dim c As String
    Dim result As String

    afterPattern = Mid(text, startPos, 20)
    afterPattern = Trim(afterPattern)

    result = ""

    For i = 1 To Len(afterPattern)
        c = Mid(afterPattern, i, 1)

        If c Like "[0-9]" Then
            result = result & c
        ElseIf c = "." Or c = "/" Or c = "-" Then
            ' Aceita separadores comuns em numeros de referencia
            If Len(result) > 0 Then result = result & c
        ElseIf Len(result) > 0 Then
            Exit For ' Terminou o numero
        End If
    Next i

    ' Remove separadores no final
    Do While Right(result, 1) = "." Or Right(result, 1) = "/" Or Right(result, 1) = "-"
        result = Left(result, Len(result) - 1)
    Loop

    ExtractReferenceNumber = result
End Function

'================================================================================
' VERIFICA CONSISTENCIA DE BAIRROS
'================================================================================
Private Function CheckBairroConsistency(ementaText As String, proposicaoText As String) As String
    On Error Resume Next
    CheckBairroConsistency = ""

    Dim bairroPatterns() As Variant
    bairroPatterns = Array("bairro ", "no bairro ", "do bairro ")

    Dim pattern As Variant
    Dim patternPos As Long
    Dim bairroInEmenta As String
    Dim normalizedProposicao As String

    normalizedProposicao = CleanTextForComparison(proposicaoText)

    For Each pattern In bairroPatterns
        patternPos = InStr(1, LCase(ementaText), CStr(pattern), vbTextCompare)

        If patternPos > 0 Then
            ' Extrai nome do bairro (ate 30 caracteres)
            bairroInEmenta = ExtractBairroName(ementaText, patternPos + Len(pattern))

            If Len(bairroInEmenta) > 2 Then
                ' Verifica se bairro existe na proposicao (com tolerancia)
                If Not CheckBairroInText(bairroInEmenta, normalizedProposicao) Then
                    CheckBairroConsistency = "BAIRRO: '" & bairroInEmenta & "' da ementa nao encontrado no texto."
                    Exit Function
                End If
            End If
        End If
    Next pattern
End Function

'================================================================================
' EXTRAI NOME DO BAIRRO
'================================================================================
Private Function ExtractBairroName(text As String, startPos As Long) As String
    On Error Resume Next
    ExtractBairroName = ""

    Dim afterPattern As String
    Dim words() As String
    Dim result As String
    Dim i As Long

    afterPattern = Mid(text, startPos, 40)
    afterPattern = CleanTextForComparison(afterPattern)

    words = Split(afterPattern, " ")
    result = ""

    For i = 0 To UBound(words)
        If i > 2 Then Exit For ' Maximo 3 palavras

        Dim word As String
        word = Trim(words(i))

        ' Para se encontrar pontuacao ou palavras-chave que indicam fim
        If InStr(word, ",") > 0 Or InStr(word, ".") > 0 Then Exit For
        If LCase(word) = "neste" Or LCase(word) = "desta" Then Exit For

        If Len(word) > 1 Then
            If result <> "" Then result = result & " "
            result = result & word
        End If
    Next i

    ExtractBairroName = result
End Function

'================================================================================
' VERIFICA SE BAIRRO EXISTE NO TEXTO
'================================================================================
Private Function CheckBairroInText(bairro As String, text As String) As Boolean
    On Error Resume Next
    CheckBairroInText = False

    ' Busca exata primeiro
    If InStr(1, text, bairro, vbTextCompare) > 0 Then
        CheckBairroInText = True
        Exit Function
    End If

    ' Busca por palavras individuais
    Dim words() As String
    Dim word As Variant
    Dim foundCount As Long

    words = Split(bairro, " ")
    foundCount = 0

    For Each word In words
        If Len(Trim(CStr(word))) > 2 Then
            If InStr(1, text, CStr(word), vbTextCompare) > 0 Then
                foundCount = foundCount + 1
            End If
        End If
    Next word

    ' Considera encontrado se achou pelo menos metade das palavras
    CheckBairroInText = (foundCount >= (UBound(words) + 1) / 2)
End Function

'================================================================================
' LIMPA TEXTO PARA COMPARACAO
'================================================================================
Private Function CleanTextForComparison(text As String) As String
    On Error Resume Next
    CleanTextForComparison = text

    Dim result As String
    result = text

    ' Remove quebras de linha
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, vbTab, " ")

    ' Normaliza variacoes de caracteres
    result = Replace(result, Chr(160), " ")  ' Non-breaking space

    ' Remove multiplos espacos
    Dim counter As Long
    counter = 0
    Do While InStr(result, "  ") > 0 And counter < 100
        result = Replace(result, "  ", " ")
        counter = counter + 1
    Loop

    CleanTextForComparison = Trim(result)
End Function

'================================================================================
' VERIFICAÇÃO DE DADOS SENSÍVEIS
'================================================================================
'================================================================================
' VERIFICACAO DE DADOS SENSIVEIS (LGPD)
' Detecta dados pessoais que requerem cuidado especial conforme Lei 13.709/2018
' Art. 5 - Dados pessoais e dados pessoais sensiveis
' Art. 11 - Tratamento de dados pessoais sensiveis
'================================================================================
Private Function CheckSensitiveData(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim docText As String
    Dim findings As String
    Dim categoryCount As Long
    Dim sensitiveSpecialCount As Long

    ' Obtem texto do documento
    docText = doc.Range.text

    If Len(docText) < 10 Then
        CheckSensitiveData = True
        Exit Function
    End If

    findings = ""
    categoryCount = 0
    sensitiveSpecialCount = 0

    ' 1. Verifica documentos de identificacao (CPF, RG, CNH, etc)
    Dim docIdFindings As String
    docIdFindings = CheckDocumentIdentifiers(docText)
    If Len(docIdFindings) > 0 Then
        findings = findings & "[1] DOCUMENTOS DE IDENTIFICACAO:" & vbCrLf & docIdFindings & vbCrLf
        categoryCount = categoryCount + 1
    End If

    ' 2. Verifica dados pessoais (filiacao, nascimento, etc)
    Dim personalFindings As String
    personalFindings = CheckPersonalData(docText)
    If Len(personalFindings) > 0 Then
        findings = findings & "[2] DADOS PESSOAIS:" & vbCrLf & personalFindings & vbCrLf
        categoryCount = categoryCount + 1
    End If

    ' 3. Verifica dados de contato (email, telefone)
    Dim contactFindings As String
    contactFindings = CheckContactData(docText)
    If Len(contactFindings) > 0 Then
        findings = findings & "[3] DADOS DE CONTATO:" & vbCrLf & contactFindings & vbCrLf
        categoryCount = categoryCount + 1
    End If

    ' 4. Verifica dados de veiculos (placa, renavam)
    Dim vehicleFindings As String
    vehicleFindings = CheckVehicleData(docText)
    If Len(vehicleFindings) > 0 Then
        findings = findings & "[4] DADOS DE VEICULOS:" & vbCrLf & vehicleFindings & vbCrLf
        categoryCount = categoryCount + 1
    End If

    ' 5. Verifica dados financeiros (conta, PIX, renda)
    Dim financialFindings As String
    financialFindings = CheckFinancialData(docText)
    If Len(financialFindings) > 0 Then
        findings = findings & "[5] DADOS FINANCEIROS:" & vbCrLf & financialFindings & vbCrLf
        categoryCount = categoryCount + 1
    End If

    ' 6. Verifica dados de saude (Art. 5, II - dado sensivel especial)
    Dim healthFindings As String
    healthFindings = CheckHealthData(docText)
    If Len(healthFindings) > 0 Then
        findings = findings & "[6] DADOS DE SAUDE (SENSIVEL ESPECIAL - Art.5,II):" & vbCrLf & healthFindings & vbCrLf
        categoryCount = categoryCount + 1
        sensitiveSpecialCount = sensitiveSpecialCount + 1
    End If

    ' 7. Verifica dados sensiveis especiais LGPD (Art. 5, II)
    Dim sensitiveSpecialFindings As String
    sensitiveSpecialFindings = CheckSensitiveSpecialData(docText)
    If Len(sensitiveSpecialFindings) > 0 Then
        findings = findings & "[7] DADOS SENSIVEIS ESPECIAIS (Art.5,II LGPD):" & vbCrLf & sensitiveSpecialFindings & vbCrLf
        categoryCount = categoryCount + 1
        sensitiveSpecialCount = sensitiveSpecialCount + 1
    End If

    ' 8. Verifica dados de menores de idade
    Dim minorFindings As String
    minorFindings = CheckMinorData(docText)
    If Len(minorFindings) > 0 Then
        findings = findings & "[8] DADOS DE MENORES (Art.14 LGPD):" & vbCrLf & minorFindings & vbCrLf
        categoryCount = categoryCount + 1
    End If

    ' 9. Verifica dados judiciais/criminais
    Dim judicialFindings As String
    judicialFindings = CheckJudicialData(docText)
    If Len(judicialFindings) > 0 Then
        findings = findings & "[9] DADOS JUDICIAIS/CRIMINAIS:" & vbCrLf & judicialFindings & vbCrLf
        categoryCount = categoryCount + 1
    End If

    ' Se encontrou dados sensiveis, exibe mensagem consolidada
    If categoryCount > 0 Then
        Dim msg As String
        msg = "DADOS SENSIVEIS DETECTADOS (Lei 13.709/2018 - LGPD)" & vbCrLf & vbCrLf

        If sensitiveSpecialCount > 0 Then
            msg = msg & "ATENCAO: " & sensitiveSpecialCount & " categoria(s) de DADOS SENSIVEIS ESPECIAIS!" & vbCrLf
            msg = msg & "(Art. 5, II - Requerem consentimento explicito)" & vbCrLf & vbCrLf
        End If

        msg = msg & "Total: " & categoryCount & " categoria(s) detectada(s):" & vbCrLf & vbCrLf
        msg = msg & findings & vbCrLf
        msg = msg & "FUNDAMENTACAO LEGAL:" & vbCrLf
        msg = msg & "  - Art. 5: Define dados pessoais e sensiveis" & vbCrLf
        msg = msg & "  - Art. 7: Bases legais para tratamento" & vbCrLf
        msg = msg & "  - Art. 11: Tratamento de dados sensiveis" & vbCrLf
        msg = msg & "  - Art. 14: Dados de criancas e adolescentes" & vbCrLf & vbCrLf
        msg = msg & "RECOMENDACOES:" & vbCrLf
        msg = msg & "  - Verifique a necessidade de cada dado" & vbCrLf
        msg = msg & "  - Anonimize ou pseudonimize quando possivel" & vbCrLf
        msg = msg & "  - Obtenha consentimento para dados sensiveis"

        MsgBox msg, vbExclamation, "Verificacao LGPD - Dados Sensiveis"

        LogMessage "LGPD: " & categoryCount & " categoria(s), " & sensitiveSpecialCount & " sensivel(is) especial(is)", LOG_LEVEL_WARNING
        CheckSensitiveData = False
        Exit Function
    End If

    LogMessage "Verificacao LGPD concluida - nenhum dado sensivel detectado", LOG_LEVEL_INFO
    CheckSensitiveData = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na verificacao LGPD: " & Err.Description, LOG_LEVEL_WARNING
    CheckSensitiveData = True
End Function

'================================================================================
' VERIFICA DADOS SENSIVEIS ESPECIAIS (Art. 5, II LGPD)
' Origem racial/etnica, conviccao religiosa, opiniao politica, filiacao sindical,
' dados de saude, vida sexual, dados geneticos ou biometricos
'================================================================================
Private Function CheckSensitiveSpecialData(docText As String) As String
    On Error Resume Next
    CheckSensitiveSpecialData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Origem racial ou etnica
    If InStr(lowerText, "raca:") > 0 Or InStr(lowerText, "etnia:") > 0 Or _
       InStr(lowerText, "cor da pele") > 0 Or InStr(lowerText, "origem etnica") > 0 Or _
       InStr(lowerText, "afrodescendente") > 0 Or InStr(lowerText, "indigena") > 0 Then
        findings = findings & "  - Origem racial/etnica detectada" & vbCrLf
    End If

    ' Conviccao religiosa
    If InStr(lowerText, "religiao:") > 0 Or InStr(lowerText, "crenca:") > 0 Or _
       InStr(lowerText, "conviccao religiosa") > 0 Or InStr(lowerText, "fe:") > 0 Or _
       InStr(lowerText, "praticante de") > 0 Then
        findings = findings & "  - Conviccao religiosa detectada" & vbCrLf
    End If

    ' Opiniao politica
    If InStr(lowerText, "opiniao politica") > 0 Or InStr(lowerText, "filiacao partidaria") > 0 Or _
       InStr(lowerText, "partido politico:") > 0 Or InStr(lowerText, "ideologia:") > 0 Then
        findings = findings & "  - Opiniao politica detectada" & vbCrLf
    End If

    ' Filiacao sindical
    If InStr(lowerText, "sindicato:") > 0 Or InStr(lowerText, "filiacao sindical") > 0 Or _
       InStr(lowerText, "sindicalizado") > 0 Or InStr(lowerText, "membro do sindicato") > 0 Then
        findings = findings & "  - Filiacao sindical detectada" & vbCrLf
    End If

    ' Vida sexual
    If InStr(lowerText, "orientacao sexual") > 0 Or InStr(lowerText, "identidade de genero") > 0 Or _
       InStr(lowerText, "vida sexual") > 0 Or InStr(lowerText, "preferencia sexual") > 0 Then
        findings = findings & "  - Dado sobre vida sexual detectado" & vbCrLf
    End If

    ' Dados geneticos
    If InStr(lowerText, "dna") > 0 Or InStr(lowerText, "genetico") > 0 Or _
       InStr(lowerText, "exame genetico") > 0 Or InStr(lowerText, "teste de paternidade") > 0 Then
        findings = findings & "  - Dado genetico detectado" & vbCrLf
    End If

    ' Dados biometricos
    If InStr(lowerText, "biometria") > 0 Or InStr(lowerText, "biometrico") > 0 Or _
       InStr(lowerText, "impressao digital") > 0 Or InStr(lowerText, "reconhecimento facial") > 0 Or _
       InStr(lowerText, "iris") > 0 Then
        findings = findings & "  - Dado biometrico detectado" & vbCrLf
    End If

    CheckSensitiveSpecialData = findings
End Function

'================================================================================
' VERIFICA DADOS DE MENORES DE IDADE (Art. 14 LGPD)
'================================================================================
Private Function CheckMinorData(docText As String) As String
    On Error Resume Next
    CheckMinorData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Mencoes a menores
    If InStr(lowerText, "menor de idade") > 0 Or InStr(lowerText, "crianca") > 0 Or _
       InStr(lowerText, "adolescente") > 0 Then
        findings = findings & "  - Referencia a menor de idade detectada" & vbCrLf
    End If

    ' Dados escolares de menores
    If (InStr(lowerText, "aluno") > 0 Or InStr(lowerText, "estudante") > 0) And _
       (InStr(lowerText, "escola") > 0 Or InStr(lowerText, "colegio") > 0) Then
        If InStr(lowerText, "fundamental") > 0 Or InStr(lowerText, "infantil") > 0 Then
            findings = findings & "  - Dados escolares de menor detectados" & vbCrLf
        End If
    End If

    ' Responsavel legal
    If InStr(lowerText, "responsavel legal") > 0 Or InStr(lowerText, "representante legal") > 0 Or _
       InStr(lowerText, "tutor:") > 0 Or InStr(lowerText, "curador:") > 0 Then
        findings = findings & "  - Mencao a responsavel legal (possivel menor)" & vbCrLf
    End If

    ' ECA - Estatuto da Crianca e Adolescente
    If InStr(lowerText, "eca") > 0 Or InStr(lowerText, "estatuto da crianca") > 0 Or _
       InStr(lowerText, "conselho tutelar") > 0 Then
        findings = findings & "  - Referencia ao ECA detectada" & vbCrLf
    End If

    CheckMinorData = findings
End Function

'================================================================================
' VERIFICA DADOS JUDICIAIS E CRIMINAIS
'================================================================================
Private Function CheckJudicialData(docText As String) As String
    On Error Resume Next
    CheckJudicialData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Antecedentes criminais
    If InStr(lowerText, "antecedentes criminais") > 0 Or InStr(lowerText, "folha corrida") > 0 Or _
       InStr(lowerText, "certidao criminal") > 0 Then
        findings = findings & "  - Antecedentes criminais detectados" & vbCrLf
    End If

    ' Processos judiciais
    If InStr(lowerText, "processo n") > 0 And (InStr(lowerText, "vara") > 0 Or _
       InStr(lowerText, "tribunal") > 0 Or InStr(lowerText, "juizo") > 0) Then
        findings = findings & "  - Numero de processo judicial detectado" & vbCrLf
    End If

    ' Inquerito policial
    If InStr(lowerText, "inquerito policial") > 0 Or InStr(lowerText, "boletim de ocorrencia") > 0 Or _
       InStr(lowerText, "b.o.") > 0 Then
        findings = findings & "  - Inquerito/BO detectado" & vbCrLf
    End If

    ' Condenacao
    If InStr(lowerText, "condenado") > 0 Or InStr(lowerText, "sentenciado") > 0 Or _
       InStr(lowerText, "apenado") > 0 Or InStr(lowerText, "reeducando") > 0 Then
        findings = findings & "  - Informacao de condenacao detectada" & vbCrLf
    End If

    ' Medida protetiva
    If InStr(lowerText, "medida protetiva") > 0 Or InStr(lowerText, "lei maria da penha") > 0 Then
        findings = findings & "  - Medida protetiva detectada" & vbCrLf
    End If

    CheckJudicialData = findings
End Function

'================================================================================
' VERIFICA DOCUMENTOS DE IDENTIFICACAO
'================================================================================
Private Function CheckDocumentIdentifiers(docText As String) As String
    On Error Resume Next
    CheckDocumentIdentifiers = ""

    Dim lowerText As String
    Dim findings As String
    Dim cpfCount As Long
    Dim rgCount As Long

    lowerText = LCase(docText)
    findings = ""

    ' Verifica mencoes a CPF
    cpfCount = 0
    If InStr(lowerText, "cpf:") > 0 Then cpfCount = cpfCount + 1
    If InStr(lowerText, "cpf n") > 0 Then cpfCount = cpfCount + 1
    If InStr(lowerText, "cpf/mf") > 0 Then cpfCount = cpfCount + 1
    If InStr(lowerText, "inscrito no cpf") > 0 Then cpfCount = cpfCount + 1

    ' Detecta padrao numerico de CPF (XXX.XXX.XXX-XX)
    If ContainsCPFPattern(docText) Then cpfCount = cpfCount + 1

    If cpfCount > 0 Then
        findings = findings & "  - CPF detectado" & vbCrLf
    End If

    ' Verifica mencoes a RG
    rgCount = 0
    If InStr(lowerText, "rg:") > 0 Then rgCount = rgCount + 1
    If InStr(lowerText, "rg n") > 0 Then rgCount = rgCount + 1
    If InStr(lowerText, "identidade n") > 0 Then rgCount = rgCount + 1
    If InStr(lowerText, "carteira de identidade") > 0 Then rgCount = rgCount + 1

    ' Detecta padrao numerico de RG
    If ContainsRGPattern(docText) Then rgCount = rgCount + 1

    If rgCount > 0 Then
        findings = findings & "  - RG/Identidade detectado" & vbCrLf
    End If

    ' CNH
    If InStr(lowerText, "cnh:") > 0 Or InStr(lowerText, "cnh n") > 0 Or _
       InStr(lowerText, "habilitacao n") > 0 Then
        findings = findings & "  - CNH detectada" & vbCrLf
    End If

    ' CTPS
    If InStr(lowerText, "ctps") > 0 Or InStr(lowerText, "carteira de trabalho") > 0 Then
        findings = findings & "  - CTPS detectada" & vbCrLf
    End If

    ' Titulo de eleitor
    If InStr(lowerText, "titulo de eleitor") > 0 Or InStr(lowerText, "titulo eleitoral") > 0 Then
        findings = findings & "  - Titulo de eleitor detectado" & vbCrLf
    End If

    ' PIS/PASEP
    If InStr(lowerText, "pis:") > 0 Or InStr(lowerText, "pis/pasep") > 0 Or _
       InStr(lowerText, "pasep:") > 0 Then
        findings = findings & "  - PIS/PASEP detectado" & vbCrLf
    End If

    CheckDocumentIdentifiers = findings
End Function

'================================================================================
' DETECTA PADRAO NUMERICO DE CPF (XXX.XXX.XXX-XX)
'================================================================================
Private Function ContainsCPFPattern(text As String) As Boolean
    On Error Resume Next
    ContainsCPFPattern = False

    Dim i As Long
    Dim segment As String
    Dim digitCount As Long
    Dim hasSeparator As Boolean

    ' Busca sequencia de 11 digitos com separadores tipicos de CPF
    For i = 1 To Len(text) - 13
        segment = Mid(text, i, 14)

        ' Verifica padrao XXX.XXX.XXX-XX
        If Mid(segment, 4, 1) = "." And Mid(segment, 8, 1) = "." And Mid(segment, 12, 1) = "-" Then
            digitCount = CountDigitsInString(segment)
            If digitCount = 11 Then
                ContainsCPFPattern = True
                Exit Function
            End If
        End If
    Next i

    ' Busca sequencia de 11 digitos consecutivos
    digitCount = 0
    For i = 1 To Len(text)
        If Mid(text, i, 1) Like "[0-9]" Then
            digitCount = digitCount + 1
            If digitCount = 11 Then
                ' Verifica se nao e parte de um numero maior
                If i < Len(text) Then
                    If Not Mid(text, i + 1, 1) Like "[0-9]" Then
                        ContainsCPFPattern = True
                        Exit Function
                    End If
                End If
            End If
        Else
            digitCount = 0
        End If
    Next i
End Function

'================================================================================
' DETECTA PADRAO NUMERICO DE RG
'================================================================================
Private Function ContainsRGPattern(text As String) As Boolean
    On Error Resume Next
    ContainsRGPattern = False

    Dim i As Long
    Dim segment As String
    Dim digitCount As Long

    ' RG geralmente tem 7-9 digitos com separadores
    ' Padrao comum: XX.XXX.XXX-X ou similar
    For i = 1 To Len(text) - 11
        segment = Mid(text, i, 12)

        ' Verifica padrao XX.XXX.XXX-X
        If Mid(segment, 3, 1) = "." And Mid(segment, 7, 1) = "." And Mid(segment, 11, 1) = "-" Then
            digitCount = CountDigitsInString(segment)
            If digitCount >= 8 And digitCount <= 10 Then
                ContainsRGPattern = True
                Exit Function
            End If
        End If
    Next i
End Function

'================================================================================
' CONTA DIGITOS EM UMA STRING
'================================================================================
Private Function CountDigitsInString(text As String) As Long
    On Error Resume Next
    CountDigitsInString = 0

    Dim i As Long
    Dim count As Long

    count = 0
    For i = 1 To Len(text)
        If Mid(text, i, 1) Like "[0-9]" Then
            count = count + 1
        End If
    Next i

    CountDigitsInString = count
End Function

'================================================================================
' VERIFICA DADOS PESSOAIS
'================================================================================
Private Function CheckPersonalData(docText As String) As String
    On Error Resume Next
    CheckPersonalData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Filiacao
    If InStr(lowerText, "nome da mae") > 0 Or InStr(lowerText, "mae:") > 0 Or _
       InStr(lowerText, "filiacao:") > 0 Or InStr(lowerText, "filho de") > 0 Or _
       InStr(lowerText, "filha de") > 0 Then
        findings = findings & "  - Filiacao detectada" & vbCrLf
    End If

    ' Data de nascimento
    If InStr(lowerText, "nascimento:") > 0 Or InStr(lowerText, "nascido em") > 0 Or _
       InStr(lowerText, "nascida em") > 0 Or InStr(lowerText, "data de nascimento") > 0 Then
        findings = findings & "  - Data de nascimento detectada" & vbCrLf
    End If

    ' Naturalidade
    If InStr(lowerText, "naturalidade:") > 0 Or InStr(lowerText, "natural de") > 0 Then
        findings = findings & "  - Naturalidade detectada" & vbCrLf
    End If

    ' Estado civil
    If InStr(lowerText, "estado civil:") > 0 Then
        findings = findings & "  - Estado civil detectado" & vbCrLf
    End If

    ' Nacionalidade
    If InStr(lowerText, "nacionalidade:") > 0 Then
        findings = findings & "  - Nacionalidade detectada" & vbCrLf
    End If

    ' Profissao/Ocupacao
    If InStr(lowerText, "profissao:") > 0 Or InStr(lowerText, "ocupacao:") > 0 Then
        findings = findings & "  - Profissao/Ocupacao detectada" & vbCrLf
    End If

    ' Endereco residencial
    If InStr(lowerText, "residente") > 0 And (InStr(lowerText, "rua ") > 0 Or _
       InStr(lowerText, "avenida ") > 0) Then
        findings = findings & "  - Endereco residencial detectado" & vbCrLf
    End If

    ' Sexo/Genero
    If InStr(lowerText, "sexo:") > 0 Or InStr(lowerText, "genero:") > 0 Then
        findings = findings & "  - Sexo/Genero detectado" & vbCrLf
    End If

    ' Escolaridade
    If InStr(lowerText, "escolaridade:") > 0 Or InStr(lowerText, "grau de instrucao") > 0 Then
        findings = findings & "  - Escolaridade detectada" & vbCrLf
    End If

    CheckPersonalData = findings
End Function

'================================================================================
' VERIFICA DADOS DE CONTATO
'================================================================================
Private Function CheckContactData(docText As String) As String
    On Error Resume Next
    CheckContactData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Email
    If ContainsEmailPattern(docText) Then
        findings = findings & "  - Email detectado" & vbCrLf
    End If

    ' Telefone
    If InStr(lowerText, "telefone:") > 0 Or InStr(lowerText, "tel:") > 0 Or _
       InStr(lowerText, "celular:") > 0 Or InStr(lowerText, "fone:") > 0 Or _
       ContainsPhonePattern(docText) Then
        findings = findings & "  - Telefone detectado" & vbCrLf
    End If

    ' WhatsApp
    If InStr(lowerText, "whatsapp") > 0 Or InStr(lowerText, "zap:") > 0 Then
        findings = findings & "  - WhatsApp detectado" & vbCrLf
    End If

    CheckContactData = findings
End Function

'================================================================================
' DETECTA PADRAO DE EMAIL
'================================================================================
Private Function ContainsEmailPattern(text As String) As Boolean
    On Error Resume Next
    ContainsEmailPattern = False

    ' Busca por @ seguido de dominio
    Dim atPos As Long
    atPos = InStr(text, "@")

    If atPos > 1 Then
        ' Verifica se tem caracteres antes e depois do @
        Dim beforeAt As String
        Dim afterAt As String

        beforeAt = Mid(text, atPos - 1, 1)
        If atPos < Len(text) - 3 Then
            afterAt = Mid(text, atPos + 1, 4)
            ' Verifica se parece um dominio (letras seguidas de ponto)
            If InStr(afterAt, ".") > 0 Then
                ContainsEmailPattern = True
            End If
        End If
    End If
End Function

'================================================================================
' DETECTA PADRAO DE TELEFONE
'================================================================================
Private Function ContainsPhonePattern(text As String) As Boolean
    On Error Resume Next
    ContainsPhonePattern = False

    Dim i As Long
    Dim segment As String
    Dim digitCount As Long

    ' Busca padrao (XX) XXXXX-XXXX ou similar
    For i = 1 To Len(text) - 13
        segment = Mid(text, i, 15)

        ' Verifica se comeca com (
        If Mid(segment, 1, 1) = "(" Then
            digitCount = CountDigitsInString(segment)
            ' Telefone brasileiro tem 10-11 digitos
            If digitCount >= 10 And digitCount <= 11 Then
                ContainsPhonePattern = True
                Exit Function
            End If
        End If
    Next i
End Function

'================================================================================
' VERIFICA DADOS DE VEICULOS
'================================================================================
Private Function CheckVehicleData(docText As String) As String
    On Error Resume Next
    CheckVehicleData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Placa
    If InStr(lowerText, "placa:") > 0 Or InStr(lowerText, "placa n") > 0 Or _
       ContainsPlacaPattern(docText) Then
        findings = findings & "  - Placa de veiculo detectada" & vbCrLf
    End If

    ' Renavam
    If InStr(lowerText, "renavam") > 0 Then
        findings = findings & "  - RENAVAM detectado" & vbCrLf
    End If

    ' Chassi
    If InStr(lowerText, "chassi") > 0 Then
        findings = findings & "  - Chassi detectado" & vbCrLf
    End If

    CheckVehicleData = findings
End Function

'================================================================================
' DETECTA PADRAO DE PLACA (ABC-1234 ou ABC1D23)
'================================================================================
Private Function ContainsPlacaPattern(text As String) As Boolean
    On Error Resume Next
    ContainsPlacaPattern = False

    Dim i As Long
    Dim segment As String
    Dim c As String

    ' Busca padrao antigo: ABC-1234 ou ABC1234
    For i = 1 To Len(text) - 6
        segment = UCase(Mid(text, i, 8))

        ' Verifica 3 letras + hifen ou digito + 4 digitos
        If Mid(segment, 1, 1) Like "[A-Z]" And _
           Mid(segment, 2, 1) Like "[A-Z]" And _
           Mid(segment, 3, 1) Like "[A-Z]" Then

            ' Padrao com hifen: ABC-1234
            If Mid(segment, 4, 1) = "-" Then
                If Mid(segment, 5, 1) Like "[0-9]" And _
                   Mid(segment, 6, 1) Like "[0-9]" And _
                   Mid(segment, 7, 1) Like "[0-9]" And _
                   Mid(segment, 8, 1) Like "[0-9]" Then
                    ContainsPlacaPattern = True
                    Exit Function
                End If
            End If

            ' Padrao Mercosul: ABC1D23
            If Mid(segment, 4, 1) Like "[0-9]" And _
               Mid(segment, 5, 1) Like "[A-Z]" And _
               Mid(segment, 6, 1) Like "[0-9]" And _
               Mid(segment, 7, 1) Like "[0-9]" Then
                ContainsPlacaPattern = True
                Exit Function
            End If
        End If
    Next i
End Function

'================================================================================
' VERIFICA DADOS FINANCEIROS
'================================================================================
Private Function CheckFinancialData(docText As String) As String
    On Error Resume Next
    CheckFinancialData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Conta bancaria
    If InStr(lowerText, "conta:") > 0 Or InStr(lowerText, "conta corrente") > 0 Or _
       InStr(lowerText, "conta poupanca") > 0 Or InStr(lowerText, "n. da conta") > 0 Then
        findings = findings & "  - Conta bancaria detectada" & vbCrLf
    End If

    ' Agencia
    If InStr(lowerText, "agencia:") > 0 Or InStr(lowerText, "ag:") > 0 Then
        findings = findings & "  - Agencia bancaria detectada" & vbCrLf
    End If

    ' PIX
    If InStr(lowerText, "pix:") > 0 Or InStr(lowerText, "chave pix") > 0 Then
        findings = findings & "  - Chave PIX detectada" & vbCrLf
    End If

    ' Salario/Renda
    If InStr(lowerText, "salario:") > 0 Or InStr(lowerText, "renda:") > 0 Or _
       InStr(lowerText, "remuneracao:") > 0 Then
        findings = findings & "  - Informacao de renda detectada" & vbCrLf
    End If

    CheckFinancialData = findings
End Function

'================================================================================
' VERIFICA DADOS DE SAUDE
'================================================================================
Private Function CheckHealthData(docText As String) As String
    On Error Resume Next
    CheckHealthData = ""

    Dim lowerText As String
    Dim findings As String

    lowerText = LCase(docText)
    findings = ""

    ' Cartao SUS
    If InStr(lowerText, "cartao sus") > 0 Or InStr(lowerText, "cns:") > 0 Or _
       InStr(lowerText, "cartao nacional de saude") > 0 Then
        findings = findings & "  - Cartao SUS detectado" & vbCrLf
    End If

    ' CID (Classificacao Internacional de Doencas)
    If InStr(lowerText, "cid:") > 0 Or InStr(lowerText, "cid-10") > 0 Or _
       InStr(lowerText, "cid 10") > 0 Then
        findings = findings & "  - Codigo CID detectado (DADO SENSIVEL ESPECIAL)" & vbCrLf
    End If

    ' Laudo medico
    If InStr(lowerText, "laudo medico") > 0 Or InStr(lowerText, "atestado medico") > 0 Then
        findings = findings & "  - Laudo/Atestado medico detectado (DADO SENSIVEL ESPECIAL)" & vbCrLf
    End If

    ' Deficiencia (dado sensivel especial)
    If InStr(lowerText, "deficiencia:") > 0 Or InStr(lowerText, "pcd") > 0 Or _
       InStr(lowerText, "pessoa com deficiencia") > 0 Then
        findings = findings & "  - Informacao de deficiencia detectada (DADO SENSIVEL ESPECIAL)" & vbCrLf
    End If

    ' Tipo sanguineo
    If InStr(lowerText, "tipo sanguineo") > 0 Or InStr(lowerText, "fator rh") > 0 Then
        findings = findings & "  - Tipo sanguineo detectado" & vbCrLf
    End If

    ' Alergia
    If InStr(lowerText, "alergia:") > 0 Or InStr(lowerText, "alergico a") > 0 Then
        findings = findings & "  - Informacao de alergia detectada" & vbCrLf
    End If

    CheckHealthData = findings
End Function

'================================================================================
'================================================================================
' SALVAMENTO INICIAL DO DOCUMENTO
'================================================================================
Private Function SaveDocumentFirst(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Salvando documento..."
    ' Log de início removido para performance

    Dim saveDialog As Object
    Set saveDialog = Application.Dialogs(wdDialogFileSaveAs)

    If saveDialog.Show <> -1 Then
        LogMessage "Operação de salvamento cancelada pelo usuário", LOG_LEVEL_INFO
        Application.StatusBar = "Cancelado"
        SaveDocumentFirst = False
        Exit Function
    End If

    ' Aguarda confirmação do salvamento com timeout de segurança
    Dim waitCount As Integer
    Dim maxWait As Integer
    maxWait = 10

    For waitCount = 1 To maxWait
        DoEvents
        If doc.Path <> "" Then Exit For
        Dim startTime As Double
        startTime = Timer
        Do While Timer < startTime + 1
            DoEvents
        Loop
        Application.StatusBar = "Salvando... (" & waitCount & "/" & maxWait & ")"
    Next waitCount

    If doc.Path = "" Then
        LogMessage "Falha ao salvar documento após " & maxWait & " tentativas", LOG_LEVEL_ERROR
        Application.StatusBar = "Falha ao salvar"
        SaveDocumentFirst = False
    Else
        ' Log de sucesso removido para performance
        Application.StatusBar = "Salvo"
        SaveDocumentFirst = True
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro durante salvamento: " & Err.Description & " (Erro #" & Err.Number & ")", LOG_LEVEL_ERROR
    Application.StatusBar = "Erro ao salvar"
    SaveDocumentFirst = False
End Function

'================================================================================
' LIMPEZA DE FORMATAÇÃO
'================================================================================
Private Function ClearAllFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Limpando formatação..."

    ' SUPER OTIMIZADO: Verificação única de conteúdo visual no documento
    Dim hasImages As Boolean
    Dim hasShapes As Boolean
    hasImages = (doc.InlineShapes.count > 0)
    hasShapes = (doc.Shapes.count > 0)
    Dim hasAnyVisualContent As Boolean
    hasAnyVisualContent = hasImages Or hasShapes

    Dim paraCount As Long
    Dim styleResetCount As Long

    If hasAnyVisualContent Then
        ' MODO SEGURO OTIMIZADO: Cache de verificações visuais por parágrafo
        Dim para As Paragraph
        Dim visualContentCache As Object ' Cache para evitar recálculos
        Set visualContentCache = CreateObject("Scripting.Dictionary")

        Dim clearCounter As Long
        clearCounter = 0
        For Each para In doc.Paragraphs
            clearCounter = clearCounter + 1
            ' DoEvents a cada 15 paragrafos para manter responsividade
            If clearCounter Mod 15 = 0 Then DoEvents

            On Error Resume Next

            ' Cache da verificação de conteúdo visual
            Dim paraKey As String
            paraKey = CStr(para.Range.Start) & "-" & CStr(para.Range.End)

            Dim hasVisualInPara As Boolean
            If visualContentCache.Exists(paraKey) Then
                hasVisualInPara = visualContentCache(paraKey)
            Else
                hasVisualInPara = HasVisualContent(para)
                visualContentCache.Add paraKey, hasVisualInPara
            End If

            If Not hasVisualInPara Then
                ' FORMATACAO CONSOLIDADA: Aplica todas as configuracoes em uma unica operacao
                With para.Range
                    ' Reset completo de fonte em uma unica operacao
                    With .Font
                        .Reset
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                        .Color = wdColorAutomatic
                        .Bold = False
                        .Italic = False
                        .Underline = wdUnderlineNone
                    End With

                    ' Reset completo de paragrafo em uma unica operacao
                    With .ParagraphFormat
                        .Reset
                        .alignment = wdAlignParagraphLeft
                        .LineSpacing = 12
                        .SpaceBefore = 0
                        .SpaceAfter = 0
                        .leftIndent = 0
                        .RightIndent = 0
                        .firstLineIndent = 0
                    End With

                    ' Reset de bordas e sombreamento
                    .Borders.Enable = False
                    .Shading.Texture = wdTextureNone
                End With
                paraCount = paraCount + 1
            Else
                ' OTIMIZADO: Para paragrafos com imagens, formatacao protegida mais rapida
                Call FormatCharacterByCharacter(para, STANDARD_FONT, STANDARD_FONT_SIZE, wdColorAutomatic, True, True)
                paraCount = paraCount + 1
            End If

            ' Protecao contra loops infinitos
            If paraCount > 1000 Then Exit For
            On Error GoTo ErrorHandler
        Next para

    Else
        ' MODO ULTRA-RÁPIDO: Sem conteúdo visual - formatação global em uma única operação
        With doc.Range
            ' Reset completo de fonte
            With .Font
                .Reset
                .Name = STANDARD_FONT
                .size = STANDARD_FONT_SIZE
                .Color = wdColorAutomatic
                .Bold = False
                .Italic = False
                .Underline = wdUnderlineNone
            End With

            ' Reset completo de parágrafo
            With .ParagraphFormat
                .Reset
                .alignment = wdAlignParagraphLeft
                .LineSpacing = 12
                .SpaceBefore = 0
                .SpaceAfter = 0
                .leftIndent = 0
                .RightIndent = 0
                .firstLineIndent = 0
            End With

            On Error Resume Next
            .Borders.Enable = False
            .Shading.Texture = wdTextureNone
            On Error GoTo ErrorHandler
        End With

        paraCount = doc.Paragraphs.count
    End If

    ' OTIMIZADO: Reset de estilos em uma unica passada
    Dim styleCounter As Long
    styleCounter = 0
    For Each para In doc.Paragraphs
        styleCounter = styleCounter + 1
        ' DoEvents a cada 20 paragrafos para manter responsividade
        If styleCounter Mod 20 = 0 Then DoEvents

        On Error Resume Next
        para.Style = "Normal"
        styleResetCount = styleResetCount + 1
        ' Protecao contra loops infinitos
        If styleResetCount > 1000 Then Exit For
        On Error GoTo ErrorHandler
    Next para

    LogMessage "Formatacao limpa: " & paraCount & " paragrafos resetados", LOG_LEVEL_INFO

    ' Cleanup do cache de conteúdo visual para evitar memory leak
    If Not visualContentCache Is Nothing Then
        visualContentCache.RemoveAll
        Set visualContentCache = Nothing
    End If

    ClearAllFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao limpar formatação: " & Err.Description, LOG_LEVEL_WARNING
    ClearAllFormatting = False ' Não falha o processo por isso
End Function

'================================================================================
' REMOVE PAGE NUMBER LINES - Remove linhas com padrão $NUMERO$/$ANO$/Página N
'================================================================================
Private Function RemovePageNumberLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim NextPara As Paragraph
    Dim paraText As String
    Dim cleanText As String
    Dim removedCount As Long
    Dim i As Long

    removedCount = 0

    ' Percorre de trás para frente para não afetar índices ao deletar
    For i = doc.Paragraphs.count To 1 Step -1
        If i > doc.Paragraphs.count Then Exit For ' Proteção dinâmica

        Set para = doc.Paragraphs(i)
        paraText = para.Range.text
        cleanText = Trim(Replace(Replace(paraText, vbCr, ""), vbLf, ""))

        ' Verifica se a linha termina com o padrão desejado
        If IsPageNumberLine(cleanText) Then
            ' Verifica se existe uma próxima linha
            Dim hasNextLine As Boolean
            Dim nextLineIsEmpty As Boolean
            hasNextLine = False
            nextLineIsEmpty = False

            If i < doc.Paragraphs.count Then
                hasNextLine = True
                Set NextPara = doc.Paragraphs(i + 1)
                Dim nextText As String
                nextText = Trim(Replace(Replace(NextPara.Range.text, vbCr, ""), vbLf, ""))

                ' Verifica se a próxima linha está em branco
                If nextText = "" And Not HasVisualContent(NextPara) Then
                    nextLineIsEmpty = True
                End If
            End If

            ' Remove a linha com padrão de paginação
            para.Range.Delete
            removedCount = removedCount + 1

            ' Se a próxima linha estava em branco, remove também
            If hasNextLine And nextLineIsEmpty Then
                ' Atualiza a referência pois os índices mudaram
                If i <= doc.Paragraphs.count Then
                    Set NextPara = doc.Paragraphs(i)
                    nextText = Trim(Replace(Replace(NextPara.Range.text, vbCr, ""), vbLf, ""))

                    ' Confirma que ainda está vazia antes de deletar
                    If nextText = "" And Not HasVisualContent(NextPara) Then
                        NextPara.Range.Delete
                        removedCount = removedCount + 1
                    End If
                End If
            End If
        End If

        ' Proteção contra processamento excessivo
        If removedCount > 500 Then Exit For
    Next i

    If removedCount > 0 Then
        LogMessage "Linhas de paginação removidas: " & removedCount & " linhas", LOG_LEVEL_INFO
    End If

    RemovePageNumberLines = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover linhas de paginação: " & Err.Description, LOG_LEVEL_WARNING
    RemovePageNumberLines = False
End Function

'================================================================================
' IS PAGE NUMBER LINE - Verifica se texto termina com padrão de paginação
'================================================================================
Private Function IsPageNumberLine(text As String) As Boolean
    On Error GoTo ErrorHandler

    IsPageNumberLine = False

    ' Verifica se está vazio
    If Len(text) < 10 Then Exit Function

    ' Converte para minúsculas para comparação case-insensitive
    Dim lowerText As String
    lowerText = LCase(text)

    ' Verifica se contém o padrão base
    If InStr(lowerText, "$numero$/$ano$/p") = 0 Then Exit Function

    ' Procura pelos padrões possíveis no final
    Dim patterns() As String
    ReDim patterns(0 To 1)
    patterns(0) = "$numero$/$ano$/página"
    patterns(1) = "$numero$/$ano$/pagina"

    Dim pattern As String
    Dim i As Long

    For i = 0 To UBound(patterns)
        pattern = patterns(i)

        ' Verifica se o padrão está presente
        Dim patternPos As Long
        patternPos = InStr(lowerText, pattern)

        If patternPos > 0 Then
            ' Extrai o texto após o padrão
            Dim afterPattern As String
            afterPattern = Trim(Mid(text, patternPos + Len(pattern)))

            ' Remove espaços
            afterPattern = Trim(afterPattern)

            ' Verifica se o que sobrou é apenas 1 ou 2 dígitos
            If Len(afterPattern) >= 1 And Len(afterPattern) <= 2 Then
                If IsNumeric(afterPattern) Then
                    IsPageNumberLine = True
                    Exit Function
                End If
            End If
        End If
    Next i

    Exit Function

ErrorHandler:
    IsPageNumberLine = False
End Function

'================================================================================
' LIMPEZA DA ESTRUTURA DO DOCUMENTO
'================================================================================
Private Function CleanDocumentStructure(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim i As Long
    Dim firstTextParaIndex As Long
    Dim emptyLinesRemoved As Long
    Dim leadingSpacesRemoved As Long
    Dim paraCount As Long

    ' Cache da contagem total de parágrafos
    paraCount = doc.Paragraphs.count

    ' Busca otimizada do primeiro parágrafo com texto
    firstTextParaIndex = -1
    For i = 1 To paraCount
        If i > doc.Paragraphs.count Then Exit For ' Proteção dinâmica

        Set para = doc.Paragraphs(i)
        Dim paraTextCheck As String
        paraTextCheck = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        ' Encontra o primeiro parágrafo com texto real
        If paraTextCheck <> "" Then
            firstTextParaIndex = i
            Exit For
        End If

        ' Proteção contra documentos muito grandes
        If i > MAX_INITIAL_PARAGRAPHS_TO_SCAN Then Exit For
    Next i

    ' OTIMIZADO: Remove linhas vazias ANTES do primeiro texto em uma única passada
    If firstTextParaIndex > 1 Then
        ' Processa de trás para frente para evitar problemas com índices
        For i = firstTextParaIndex - 1 To 1 Step -1
            If i > doc.Paragraphs.count Or i < 1 Then Exit For ' Proteção dinâmica

            Set para = doc.Paragraphs(i)
            Dim paraTextEmpty As String
            paraTextEmpty = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

            ' OTIMIZADO: Verificação visual só se necessário
            If paraTextEmpty = "" Then
                If Not HasVisualContent(para) Then
                    para.Range.Delete
                    emptyLinesRemoved = emptyLinesRemoved + 1
                    ' Atualiza cache após remoção
                    paraCount = paraCount - 1
                End If
            End If
        Next i
    End If

    ' Usa Find/Replace que é muito mais rápido que loop por parágrafo
    Dim rng As Range
    Set rng = doc.Range

    ' Remove espaços no início de linhas usando Find/Replace
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False

        ' Remove espaços/tabs no início de linhas usando Find/Replace simples
        .text = "^p "  ' Quebra seguida de espaço
        .Replacement.text = "^p"

        Do While .Execute(Replace:=True)
            leadingSpacesRemoved = leadingSpacesRemoved + 1
            ' Proteção contra loop infinito
            If leadingSpacesRemoved > MAX_LOOP_ITERATIONS Then Exit Do
        Loop

        ' Remove tabs no início de linhas
        .text = "^p^t"  ' Quebra seguida de tab
        .Replacement.text = "^p"

        Do While .Execute(Replace:=True)
            leadingSpacesRemoved = leadingSpacesRemoved + 1
            If leadingSpacesRemoved > MAX_LOOP_ITERATIONS Then Exit Do
        Loop
    End With

    ' Segunda passada para espaços no início do documento (sem ^p precedente)
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False  ' Não usa wildcards nesta seção

        ' Posiciona no início do documento
        rng.Start = 0
        rng.End = 1

        ' Remove espaços/tabs no início absoluto do documento
        If rng.text = " " Or rng.text = vbTab Then
            ' Expande o range para pegar todos os espaços iniciais usando método seguro
            Do While rng.End <= doc.Range.End And (SafeGetLastCharacter(rng) = " " Or SafeGetLastCharacter(rng) = vbTab)
                rng.End = rng.End + 1
                leadingSpacesRemoved = leadingSpacesRemoved + 1
                If leadingSpacesRemoved > 100 Then Exit Do ' Proteção
            Loop

            If rng.Start < rng.End - 1 Then
                rng.Delete
            End If
        End If
    End With

    ' Log simplificado apenas se houve limpeza significativa
    If emptyLinesRemoved > 0 Then
        LogMessage "Estrutura limpa: " & emptyLinesRemoved & " linhas vazias removidas"
    End If

    CleanDocumentStructure = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na limpeza da estrutura: " & Err.Description, LOG_LEVEL_ERROR
    CleanDocumentStructure = False
End Function

'================================================================================
' REMOVE ALL TAB MARKS - Remove todas as marcas de tabulação do documento
'================================================================================
Private Function RemoveAllTabMarks(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim rng As Range
    Dim tabsRemoved As Long
    tabsRemoved = 0

    Set rng = doc.Range

    ' Remove todas as tabulações substituindo por espaço simples
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "^t"  ' ^t representa tabulação
        .Replacement.text = " "  ' Substitui por espaço simples
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        Do While .Execute(Replace:=True)
            tabsRemoved = tabsRemoved + 1
            ' Proteção contra loop infinito
            If tabsRemoved > 10000 Then
                LogMessage "Limite de remoção de tabulações atingido", LOG_LEVEL_WARNING
                Exit Do
            End If
        Loop
    End With

    If tabsRemoved > 0 Then
        LogMessage "Marcas de tabulação removidas: " & tabsRemoved & " ocorrências", LOG_LEVEL_INFO
    End If

    RemoveAllTabMarks = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover marcas de tabulação: " & Err.Description, LOG_LEVEL_ERROR
    RemoveAllTabMarks = False
End Function

'================================================================================
' REPLACE LINE BREAKS WITH PARAGRAPH BREAKS - Substitui quebras de linha por quebras de parágrafo
'================================================================================
Private Function ReplaceLineBreaksWithParagraphBreaks(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim rng As Range
    Dim breaksReplaced As Long
    breaksReplaced = 0

    Set rng = doc.Range

    ' Substitui todas as quebras de linha manuais (^l) por quebras de parágrafo (^p)
    ' ^l = Shift+Enter (quebra de linha manual/soft return)
    ' ^p = Enter (quebra de parágrafo/hard return)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "^l"  ' ^l representa quebra de linha manual (Shift+Enter)
        .Replacement.text = "^p"  ' ^p representa quebra de parágrafo (Enter)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        Do While .Execute(Replace:=True)
            breaksReplaced = breaksReplaced + 1
            ' Proteção contra loop infinito
            If breaksReplaced > 10000 Then
                LogMessage "Limite de substituição de quebras de linha atingido", LOG_LEVEL_WARNING
                Exit Do
            End If
        Loop
    End With

    If breaksReplaced > 0 Then
        LogMessage "Quebras de linha substituídas por quebras de parágrafo: " & breaksReplaced & " ocorrências", LOG_LEVEL_INFO
    End If

    ReplaceLineBreaksWithParagraphBreaks = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao substituir quebras de linha: " & Err.Description, LOG_LEVEL_ERROR
    ReplaceLineBreaksWithParagraphBreaks = False
End Function

'================================================================================
' REMOVE PAGE BREAKS - Remove todas as quebras de página do documento
'================================================================================
Private Function RemovePageBreaks(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim rng As Range
    Dim breaksRemoved As Long
    breaksRemoved = 0

    Set rng = doc.Range

    ' Remove quebras de página manuais (^m)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "^m"  ' ^m representa quebra de página manual
        .Replacement.text = ""  ' Substitui por nada (remove)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        Do While .Execute(Replace:=True)
            breaksRemoved = breaksRemoved + 1
            ' Proteção contra loop infinito
            If breaksRemoved > 1000 Then
                LogMessage "Limite de remoção de quebras de página atingido", LOG_LEVEL_WARNING
                Exit Do
            End If
        Loop
    End With

    If breaksRemoved > 0 Then
        LogMessage "Quebras de página removidas: " & breaksRemoved & " ocorrências", LOG_LEVEL_INFO
    End If

    RemovePageBreaks = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover quebras de página: " & Err.Description, LOG_LEVEL_ERROR
    RemovePageBreaks = False
End Function

'================================================================================
' SAFE CHECK FOR VISUAL CONTENT - VERIFICAÇÃO SEGURA DE CONTEÚDO VISUAL
'================================================================================
Private Function HasVisualContent(para As Paragraph) As Boolean
    ' Usa a função segura implementada para compatibilidade total
    HasVisualContent = SafeHasVisualContent(para)
End Function

'================================================================================
' FORMATAÇÃO DO TÍTULO DO DOCUMENTO
'================================================================================
Private Function FormatDocumentTitle(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim firstPara As Paragraph
    Dim paraText As String
    Dim words() As String
    Dim i As Long
    Dim newText As String
    Dim testRange As Range

    ' Verifica se o documento esta protegido
    If doc.ProtectionType <> wdNoProtection Then
        LogMessage "Documento protegido - formatacao de titulo ignorada", LOG_LEVEL_INFO
        FormatDocumentTitle = True
        Exit Function
    End If

    ' Testa se e possivel editar o primeiro paragrafo
    On Error Resume Next
    Set testRange = doc.Paragraphs(1).Range
    If testRange Is Nothing Then
        Err.Clear
        On Error GoTo ErrorHandler
        LogMessage "Range invalido - formatacao de titulo ignorada", LOG_LEVEL_INFO
        FormatDocumentTitle = True
        Exit Function
    End If
    ' Tenta modificar uma propriedade para verificar acesso de escrita
    Dim originalBold As Boolean
    originalBold = testRange.Font.Bold
    testRange.Font.Bold = originalBold
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ErrorHandler
        LogMessage "Selecao protegida - formatacao de titulo ignorada", LOG_LEVEL_INFO
        FormatDocumentTitle = True
        Exit Function
    End If
    On Error GoTo ErrorHandler

    ' Encontra o primeiro parágrafo com texto (após exclusão de linhas em branco)
    For i = 1 To doc.Paragraphs.count
        Set firstPara = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(firstPara.Range.text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then
            Exit For
        End If
    Next i

    If paraText = "" Then
        LogMessage "Nenhum texto encontrado para formatação do título", LOG_LEVEL_WARNING
        FormatDocumentTitle = True
        Exit Function
    End If

    ' Remove ponto final se existir
    If Right(paraText, 1) = "." Then
        paraText = Left(paraText, Len(paraText) - 1)
    End If

    ' Verifica se é uma proposição (para aplicar substituição $NUMERO$/$ANO$)
    Dim isProposition As Boolean
    Dim firstWord As String

    words = Split(paraText, " ")
    If UBound(words) >= 0 Then
        firstWord = LCase(Trim(words(0)))
        If firstWord = "indicação" Or firstWord = "requerimento" Or firstWord = "moção" Then
            isProposition = True
        End If
    End If

    ' Se for proposição, substitui a última palavra por $NUMERO$/$ANO$
    If isProposition And UBound(words) > 0 Then ' FIX: Changed >= 0 to > 0
        ' Reconstrói o texto substituindo a última palavra com validação
        newText = ""
        For i = 0 To UBound(words) - 1
            If i > 0 Then newText = newText & " "
            newText = newText & words(i)
        Next i

        ' Adiciona $NUMERO$/$ANO$ no lugar da última palavra
        If newText <> "" Then newText = newText & " "
        newText = newText & "$NUMERO$/$ANO$"
    Else
        ' Se não for proposição, mantém o texto original
        newText = paraText
    End If

    ' SEMPRE aplica formatação de título: caixa alta, negrito, sublinhado
    firstPara.Range.text = UCase(newText)

    ' Formatação completa do título (primeira linha)
    With firstPara.Range.Font
        .Bold = True
        .Underline = wdUnderlineSingle
    End With

    With firstPara.Format
        .alignment = wdAlignParagraphCenter
        .leftIndent = 0
        .firstLineIndent = 0
        .RightIndent = 0
        .SpaceBefore = 0
        .SpaceAfter = 6  ' Pequeno espaço após o título
    End With

    If isProposition Then
        LogMessage "Título de proposição formatado: " & newText & " (centralizado, caixa alta, negrito, sublinhado)", LOG_LEVEL_INFO
    Else
        LogMessage "Primeira linha formatada como título: " & newText & " (centralizado, caixa alta, negrito, sublinhado)", LOG_LEVEL_INFO
    End If

    FormatDocumentTitle = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação do título: " & Err.Description, LOG_LEVEL_ERROR
    FormatDocumentTitle = False
End Function

'================================================================================
' FORMATACAO DE PARAGRAFOS "CONSIDERANDO" E "ANTE O EXPOSTO"
'================================================================================
Private Function FormatConsiderandoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim rng As Range
    Dim totalFormatted As Long
    Dim anteExpostoFormatted As Long
    Dim i As Long
    Dim nextChar As String

    ' Percorre todos os paragrafos procurando por "considerando" ou "ante o exposto" no inicio
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        ' Verifica se o paragrafo comeca com "considerando" (ignorando maiusculas/minusculas)
        If Len(paraText) >= 12 And LCase(Left(paraText, 12)) = "considerando" Then
            ' Verifica se apos "considerando" vem espaco, virgula, ponto-e-virgula ou fim da linha
            If Len(paraText) > 12 Then
                nextChar = Mid(paraText, 13, 1)
                If nextChar = " " Or nextChar = "," Or nextChar = ";" Or nextChar = ":" Then
                    ' E realmente "considerando" no inicio do paragrafo
                    Set rng = para.Range

                    ' CORRECAO: Usa Find/Replace para preservar espacamento
                    With rng.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .text = "considerando"
                        .Replacement.text = "CONSIDERANDO"
                        .Replacement.Font.Bold = True
                        .MatchCase = False
                        .MatchWholeWord = False
                        .Forward = True
                        .Wrap = wdFindStop

                        ' Limita a busca ao inicio do paragrafo
                        rng.End = rng.Start + 15

                        If .Execute(Replace:=True) Then
                            totalFormatted = totalFormatted + 1
                        End If
                    End With
                End If
            Else
                ' Paragrafo contem apenas "considerando"
                Set rng = para.Range
                rng.End = rng.Start + 12

                With rng
                    .text = "CONSIDERANDO"
                    .Font.Bold = True
                End With

                totalFormatted = totalFormatted + 1
            End If

        ' Verifica se o paragrafo comeca com "ante o exposto" (14 caracteres)
        ElseIf Len(paraText) >= 14 And LCase(Left(paraText, 14)) = "ante o exposto" Then
            ' Verifica se apos "ante o exposto" vem espaco, virgula, ponto-e-virgula ou fim
            If Len(paraText) > 14 Then
                nextChar = Mid(paraText, 15, 1)
                If nextChar = " " Or nextChar = "," Or nextChar = ";" Or nextChar = ":" Then
                    Set rng = para.Range

                    With rng.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .text = "ante o exposto"
                        .Replacement.text = "ANTE O EXPOSTO"
                        .Replacement.Font.Bold = True
                        .MatchCase = False
                        .MatchWholeWord = False
                        .Forward = True
                        .Wrap = wdFindStop

                        rng.End = rng.Start + 17

                        If .Execute(Replace:=True) Then
                            anteExpostoFormatted = anteExpostoFormatted + 1
                        End If
                    End With
                End If
            Else
                ' Paragrafo contem apenas "ante o exposto"
                Set rng = para.Range
                rng.End = rng.Start + 14

                With rng
                    .text = "ANTE O EXPOSTO"
                    .Font.Bold = True
                End With

                anteExpostoFormatted = anteExpostoFormatted + 1
            End If
        End If
    Next i

    If totalFormatted > 0 Then
        LogMessage "Formatacao 'CONSIDERANDO' aplicada: " & totalFormatted & " ocorrencia(s)", LOG_LEVEL_INFO
    End If
    If anteExpostoFormatted > 0 Then
        LogMessage "Formatacao 'ANTE O EXPOSTO' aplicada: " & anteExpostoFormatted & " ocorrencia(s)", LOG_LEVEL_INFO
    End If

    FormatConsiderandoParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatacao CONSIDERANDO/ANTE O EXPOSTO: " & Err.Description, LOG_LEVEL_ERROR
    FormatConsiderandoParagraphs = False
End Function

'================================================================================
' FUNÇÃO AUXILIAR DE FIND/REPLACE - Elimina código repetitivo
'================================================================================
Private Function ExecuteFindReplace(doc As Document, _
                                    searchText As String, _
                                    replaceText As String, _
                                    Optional matchCase As Boolean = False, _
                                    Optional maxIterations As Long = 500) As Long
    ' Retorna quantidade de substituicoes realizadas
    On Error Resume Next
    ExecuteFindReplace = 0

    If doc Is Nothing Then Exit Function
    If searchText = "" Then Exit Function

    Dim rng As Range
    Set rng = doc.Range
    If rng Is Nothing Then Exit Function

    Dim iterCount As Long
    iterCount = 0

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = searchText
        .Replacement.text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = matchCase
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        Do While .Execute(Replace:=True) And iterCount < maxIterations
            iterCount = iterCount + 1
            ExecuteFindReplace = ExecuteFindReplace + 1
        Loop
    End With

    Err.Clear
End Function

'================================================================================
' FORMATACAO DE "IN LOCO" EM ITALICO (REMOVE ASPAS)
'================================================================================
Private Sub FormatInLocoItalic(doc As Document)
    On Error GoTo ErrorHandler

    If doc Is Nothing Then Exit Sub

    Dim rng As Range
    Dim inLocoCount As Long
    inLocoCount = 0

    ' Procura por "in loco" (com aspas) e substitui por in loco em italico
    Set rng = doc.Range

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = Chr(34) & "in loco" & Chr(34)
        .Replacement.text = "in loco"
        .Replacement.Font.Italic = True
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False

        Do While .Execute(Replace:=True)
            inLocoCount = inLocoCount + 1
            If inLocoCount > 100 Then Exit Do  ' Limite de seguranca
        Loop
    End With

    If inLocoCount > 0 Then
        LogMessage "Formatacao 'in loco' aplicada: " & inLocoCount & " ocorrencia(s) em italico", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao formatar 'in loco': " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' APLICACAO DE SUBSTITUICOES DE TEXTO
'================================================================================
Private Function ApplyTextReplacements(doc As Document) As Boolean
    Dim errorContext As String
    Dim i As Long  ' Movida para escopo de função
    On Error GoTo ErrorHandler

    ' Validação de documento
    If Not ValidateDocument(doc) Then
        ApplyTextReplacements = False
        Exit Function
    End If

    ' Verifica se há conteúdo suficiente
    If doc.Range.text = "" Or Len(Trim(doc.Range.text)) <= 1 Then
        LogMessage "Documento vazio - substituições de texto ignoradas", LOG_LEVEL_INFO
        ApplyTextReplacements = True
        Exit Function
    End If

    Dim rng As Range
    Dim replacementCount As Long
    Dim wasReplaced As Boolean
    Dim totalReplacements As Long
    totalReplacements = 0

    ' Funcionalidade 10: Substitui variantes de "d'Oeste"
    Dim dOesteVariants() As String

    ' Define as variantes possiveis dos 3 primeiros caracteres de "d'Oeste"
    ReDim dOesteVariants(0 To 15)
    dOesteVariants(0) = "d'O"   ' Original
    dOesteVariants(1) = "d" & Chr(180) & "O"   ' Acento agudo (Chr 180)
    dOesteVariants(2) = "d`O"   ' Acento grave
    dOesteVariants(3) = "d" & ChrW(8220) & "O"   ' Aspas curvas esquerda (Unicode)
    dOesteVariants(4) = "d'o"   ' Minuscula
    dOesteVariants(5) = "d" & Chr(180) & "o"
    dOesteVariants(6) = "d`o"
    dOesteVariants(7) = "d" & ChrW(8220) & "o"
    dOesteVariants(8) = "D'O"   ' Maiuscula no D
    dOesteVariants(9) = "D" & Chr(180) & "O"
    dOesteVariants(10) = "D`O"
    dOesteVariants(11) = "D" & ChrW(8220) & "O"
    dOesteVariants(12) = "D'o"
    dOesteVariants(13) = "D" & Chr(180) & "o"
    dOesteVariants(14) = "D`o"
    dOesteVariants(15) = "D" & ChrW(8220) & "o"

    ' Valida o array antes de processar
    On Error Resume Next
    Dim arraySize As Long
    arraySize = UBound(dOesteVariants)
    If Err.Number <> 0 Or arraySize < 0 Then
        LogMessage "Erro ao inicializar array de variantes - substituições de texto ignoradas", LOG_LEVEL_WARNING
        Err.Clear
        ApplyTextReplacements = True
        Exit Function
    End If
    On Error GoTo ErrorHandler

    ' Processa cada variante de forma segura
    For i = 0 To arraySize
        On Error Resume Next
        errorContext = "dOesteVariants(" & i & ")"
        ' Valida a variante antes de usar
        If IsEmpty(dOesteVariants(i)) Or dOesteVariants(i) = "" Then
            GoTo NextVariant
        End If
        ' Cria novo range para cada busca
        Set rng = Nothing
        Set rng = doc.Range
        ' Verifica se o range foi criado com sucesso
        If rng Is Nothing Then GoTo NextVariant
        ' Configura os parâmetros de busca e substituição
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = dOesteVariants(i) & "este"
            .Replacement.text = "d'Oeste"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            ' Executa a substituição e armazena resultado booleano
            wasReplaced = .Execute(Replace:=wdReplaceAll)
            ' Verifica se houve erro
            If Err.Number = 0 Then
                If wasReplaced Then
                    totalReplacements = totalReplacements + 1
                End If
            Else
                If Err.Number <> 0 Then
                    LogMessage "Aviso ao substituir variante #" & i & " ('" & dOesteVariants(i) & "este'): " & Err.Description, LOG_LEVEL_WARNING
                End If
                Err.Clear
            End If
        End With
NextVariant:
        On Error GoTo ErrorHandler
        Err.Clear
    Next i

    If totalReplacements > 0 Then
        LogMessage "Substituições de texto aplicadas: " & totalReplacements & " variante(s) substituída(s)", LOG_LEVEL_INFO
    Else
        LogMessage "Substituições de texto: nenhuma ocorrência encontrada", LOG_LEVEL_INFO
    End If

    ' Funcionalidade 11: Substitui " ao Setor, " por " ao setor competente"
    Dim setorCount As Long
    setorCount = ExecuteFindReplace(doc, " ao Setor, ", " ao setor competente", True)
    If setorCount > 0 Then
        LogMessage "Substituicao aplicada: ' ao Setor, ' -> ' ao setor competente' (" & setorCount & "x)", LOG_LEVEL_INFO
    End If

    ' Funcionalidade 12: Substitui " Setor Competente " por " setor competente " (case insensitive)
    Dim competenteCount As Long
    competenteCount = ExecuteFindReplace(doc, " Setor Competente ", " setor competente ", False)
    If competenteCount > 0 Then
        LogMessage "Substituicao aplicada: ' Setor Competente ' -> ' setor competente ' (" & competenteCount & "x)", LOG_LEVEL_INFO
    End If

    ' Funcionalidade 13: Normaliza variantes de "tapa-buracos"
    Dim tapaBuracosCount As Long
    tapaBuracosCount = 0
    ' Com aspas
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(34) & "tapa buraco" & Chr(34), "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(34) & "tapa buracos" & Chr(34), "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(34) & "tapa-buraco" & Chr(34), "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, Chr(34) & "tapa-buracos" & Chr(34), "tapa-buracos", False)
    ' Sem aspas (ordem importa: primeiro os com hifen para evitar duplicacao)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, "tapa-buraco ", "tapa-buracos ", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, "tapa buracos", "tapa-buracos", False)
    tapaBuracosCount = tapaBuracosCount + ExecuteFindReplace(doc, "tapa buraco", "tapa-buracos", False)
    If tapaBuracosCount > 0 Then
        LogMessage "Substituicao aplicada: variantes de 'tapa-buracos' normalizadas (" & tapaBuracosCount & "x)", LOG_LEVEL_INFO
    End If

    ' Funcionalidade 14: Substitui "in loco" (com aspas) por in loco (italico, sem aspas)
    FormatInLocoItalic doc

    ApplyTextReplacements = True
    Exit Function

ErrorHandler:
    Dim errMsg As String
    errMsg = Err.Description
    If Len(errorContext) > 0 Then
        LogMessage "Erro nas substituicoes de texto (contexto: " & errorContext & "): " & errMsg, LOG_LEVEL_WARNING
    ElseIf i >= 0 And i <= 15 Then
        LogMessage "Erro nas substituicoes de texto (variante: " & CStr(i) & "): " & errMsg, LOG_LEVEL_WARNING
    Else
        LogMessage "Erro nas substituicoes de texto: " & errMsg, LOG_LEVEL_WARNING
    End If
    ' Continua execucao - erros de substituicao nao sao criticos
    ApplyTextReplacements = True
End Function

'================================================================================
' APPLY BOLD TO SPECIAL PARAGRAPHS - SIMPLIFIED & OPTIMIZED
'================================================================================
Private Sub ApplyBoldToSpecialParagraphs(doc As Document)
    On Error GoTo ErrorHandler

    If Not ValidateDocument(doc) Then Exit Sub

    Dim para As Paragraph
    Dim cleanText As String
    Dim specialParagraphs As Collection
    Set specialParagraphs = New Collection

    ' FASE 1: Identificar parágrafos especiais (uma única passada)
    Dim paraCounter As Long
    paraCounter = 0
    For Each para In doc.Paragraphs
        paraCounter = paraCounter + 1
        If paraCounter Mod 25 = 0 Then DoEvents ' Responsividade

        If Not HasVisualContent(para) Then
            cleanText = GetCleanParagraphText(para)

            ' Adiciona apenas Justificativa e Anexo (Vereador não recebe negrito)
            If cleanText = JUSTIFICATIVA_TEXT Or _
               IsAnexoPattern(cleanText) Then
                specialParagraphs.Add para
            End If
        End If
    Next para

    ' FASE 2: Aplicar negrito E reforçar alinhamento atomicamente
    ' Não controla ScreenUpdating aqui - deixa a função principal controlar

    Dim p As Variant
    Dim pCleanText As String
    For Each p In specialParagraphs
        Set para = p ' Converte Variant para Paragraph

        ' Aplica negrito
        With para.Range.Font
            .Bold = True
            .Name = STANDARD_FONT
            .size = STANDARD_FONT_SIZE
        End With

        ' REFORÇO: Garante alinhamento correto baseado no tipo
        pCleanText = GetCleanParagraphText(para)
        If pCleanText = JUSTIFICATIVA_TEXT Then
            ' Justificativa: centralizado (linhas em branco serão inseridas depois)
            para.Format.alignment = wdAlignParagraphCenter
            para.Format.leftIndent = 0
            para.Format.firstLineIndent = 0
            para.Format.RightIndent = 0
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0
        ElseIf IsAnexoPattern(pCleanText) Then
            ' Anexo/Anexos: alinhado à esquerda
            para.Format.alignment = wdAlignParagraphLeft
            para.Format.leftIndent = 0
            para.Format.firstLineIndent = 0
            para.Format.RightIndent = 0
        End If
    Next p

    LogMessage "Negrito e alinhamento aplicados a " & specialParagraphs.count & " parágrafos especiais", LOG_LEVEL_INFO
    Exit Sub

ErrorHandler:
    LogMessage "Erro ao aplicar negrito a parágrafos especiais: " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' FORMAT VEREADOR PARAGRAPHS - Formata parágrafo com "vereador" e adjacentes
'================================================================================
Private Sub FormatVereadorParagraphs(doc As Document)
    On Error GoTo ErrorHandler

    If Not ValidateDocument(doc) Then Exit Sub

    Dim para As Paragraph
    Dim prevPara As Paragraph
    Dim NextPara As Paragraph
    Dim cleanText As String
    Dim i As Long
    Dim formattedCount As Long

    formattedCount = 0

    ' Procura por parágrafos com "vereador"
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)

        If Not HasVisualContent(para) Then
            cleanText = GetCleanParagraphText(para)

            If IsVereadorPattern(cleanText) Then
                ' Remove negrito do parágrafo "vereador"
                With para.Range.Font
                    .Bold = False
                    .Name = STANDARD_FONT
                    .size = STANDARD_FONT_SIZE
                End With

                ' Centraliza e zera recuo do próprio parágrafo "vereador"
                With para.Format
                    .alignment = wdAlignParagraphCenter
                    .leftIndent = 0
                    .firstLineIndent = 0
                    .RightIndent = 0
                End With

                ' Formata linha ACIMA (se existir): centraliza, zera recuo, aplica caixa alta e negrito
                If i > 1 Then
                    Set prevPara = doc.Paragraphs(i - 1)
                    If Not HasVisualContent(prevPara) Then
                        ' Aplica caixa alta e negrito na fonte
                        With prevPara.Range.Font
                            .AllCaps = True
                            .Bold = True
                            .Name = STANDARD_FONT
                            .size = STANDARD_FONT_SIZE
                        End With

                        ' Centraliza e zera recuos
                        With prevPara.Format
                            .alignment = wdAlignParagraphCenter
                            .leftIndent = 0
                            .firstLineIndent = 0
                            .RightIndent = 0
                        End With
                    End If
                End If

                ' Formata linha ABAIXO (se existir)
                If i < doc.Paragraphs.count Then
                    Set NextPara = doc.Paragraphs(i + 1)
                    If Not HasVisualContent(NextPara) Then
                        With NextPara.Format
                            .alignment = wdAlignParagraphCenter
                            .leftIndent = 0
                            .firstLineIndent = 0
                            .RightIndent = 0
                        End With
                    End If
                End If

                formattedCount = formattedCount + 1
                LogMessage "Parágrafo 'Vereador' formatado (sem negrito) com linhas adjacentes centralizadas (posição: " & i & ")", LOG_LEVEL_INFO
            End If
        End If
    Next i

    If formattedCount > 0 Then
        LogMessage "Formatação 'Vereador': " & formattedCount & " ocorrências formatadas", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao formatar parágrafos 'Vereador': " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' FUNÇÕES AUXILIARES PARA MANIPULAÇÃO DE LINHAS EM BRANCO
'================================================================================

' Remove linhas vazias ANTES de um paragrafo especifico
' Retorna o novo indice do paragrafo apos remocoes
Private Function RemoveBlankLinesBefore(doc As Document, ByVal targetIndex As Long) As Long
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long

    i = targetIndex - 1
    Do While i >= 1
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        If paraText = "" And Not HasVisualContent(para) Then
            para.Range.Delete
            targetIndex = targetIndex - 1
            i = i - 1
        Else
            Exit Do
        End If
    Loop

    RemoveBlankLinesBefore = targetIndex
    Exit Function

ErrorHandler:
    RemoveBlankLinesBefore = targetIndex
End Function

' Remove linhas vazias DEPOIS de um paragrafo especifico
Private Sub RemoveBlankLinesAfter(doc As Document, ByVal targetIndex As Long)
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long

    i = targetIndex + 1
    Do While i <= doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        If paraText = "" And Not HasVisualContent(para) Then
            para.Range.Delete
        Else
            Exit Do
        End If
    Loop

    Exit Sub

ErrorHandler:
    ' Silently continue
End Sub

' Insere N linhas em branco ANTES de um paragrafo
Private Sub InsertBlankLinesBefore(doc As Document, ByVal targetIndex As Long, ByVal lineCount As Long)
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim j As Long

    Set para = doc.Paragraphs(targetIndex)
    For j = 1 To lineCount
        para.Range.InsertParagraphBefore
    Next j

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao inserir linhas antes: " & Err.Description, LOG_LEVEL_WARNING
End Sub

' Insere N linhas em branco DEPOIS de um paragrafo
Private Sub InsertBlankLinesAfter(doc As Document, ByVal targetIndex As Long, ByVal lineCount As Long)
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim j As Long

    Set para = doc.Paragraphs(targetIndex)
    For j = 1 To lineCount
        para.Range.InsertParagraphAfter
    Next j

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao inserir linhas depois: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' INSERÇÃO DE LINHAS EM BRANCO NA JUSTIFICATIVA
'================================================================================
Private Sub InsertJustificativaBlankLines(doc As Document)
    On Error GoTo ErrorHandler

    If Not ValidateDocument(doc) Then Exit Sub

    Dim para As Paragraph
    Dim cleanText As String
    Dim i As Long
    Dim justificativaIndex As Long
    Dim paraText As String

    ' Não controla ScreenUpdating aqui - deixa a função principal controlar

    ' FASE 1: Localiza o parágrafo "Justificativa"
    justificativaIndex = 0
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)

        If Not HasVisualContent(para) Then
            cleanText = GetCleanParagraphText(para)

            If cleanText = JUSTIFICATIVA_TEXT Then
                justificativaIndex = i
                Exit For
            End If
        End If
    Next i

    If justificativaIndex = 0 Then
        Exit Sub ' Nao encontrou "Justificativa"
    End If

    ' FASE 2-5: Remove linhas vazias e insere exatamente 2 antes e 2 depois
    justificativaIndex = RemoveBlankLinesBefore(doc, justificativaIndex)
    RemoveBlankLinesAfter doc, justificativaIndex
    InsertBlankLinesBefore doc, justificativaIndex, 2
    InsertBlankLinesAfter doc, justificativaIndex + 2, 2  ' +2 por causa das insercoes anteriores

    LogMessage "Linhas em branco ajustadas: 2 antes e 2 depois de 'Justificativa'", LOG_LEVEL_INFO

    ' FASE 6: Processa "Plenário Dr. Tancredo Neves"
    Dim plenarioIndex As Long
    Dim paraTextLower As String

    plenarioIndex = 0
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)

        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            paraTextLower = LCase(paraText)

            ' Procura por "Plenário" e "Tancredo Neves" (case insensitive)
            If InStr(paraTextLower, "plenário") > 0 And _
               InStr(paraTextLower, "tancredo") > 0 And _
               InStr(paraTextLower, "neves") > 0 Then
                plenarioIndex = i
                Exit For
            End If
        End If
    Next i

    If plenarioIndex > 0 Then
        ' Remove linhas vazias e insere exatamente 2 antes e 2 depois
        plenarioIndex = RemoveBlankLinesBefore(doc, plenarioIndex)
        RemoveBlankLinesAfter doc, plenarioIndex
        InsertBlankLinesBefore doc, plenarioIndex, 2
        InsertBlankLinesAfter doc, plenarioIndex + 2, 2

        LogMessage "2 linhas em branco inseridas antes e depois de 'Plenario Dr. Tancredo Neves'", LOG_LEVEL_INFO
    End If

    ' FASE 7: Processa "Excelentíssimo Senhor Prefeito Municipal,"
    Dim prefeitoIndex As Long

    prefeitoIndex = 0
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)

        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            paraTextLower = LCase(paraText)

            ' Procura por "Excelentíssimo Senhor Prefeito Municipal" (case insensitive)
            If InStr(paraTextLower, "excelentíssimo") > 0 And _
               InStr(paraTextLower, "senhor") > 0 And _
               InStr(paraTextLower, "prefeito") > 0 And _
               InStr(paraTextLower, "municipal") > 0 Then
                prefeitoIndex = i
                Exit For
            End If
        End If
    Next i

    If prefeitoIndex > 0 Then
        ' Remove linhas vazias depois e insere exatamente 2
        RemoveBlankLinesAfter doc, prefeitoIndex
        InsertBlankLinesAfter doc, prefeitoIndex, 2

        LogMessage "2 linhas em branco inseridas apos 'Excelentissimo Senhor Prefeito Municipal,'", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao inserir linhas em branco: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' FUNÇÕES AUXILIARES PARA DETECÇÃO DE PADRÕES
'================================================================================
Private Function IsVereadorPattern(text As String) As Boolean
    ' Remove espaços extras para análise
    Dim cleanText As String
    cleanText = Trim(text)

    ' Remove hifens/travessões do início e fim e espaços adjacentes
    cleanText = Trim(cleanText)
    If Left(cleanText, 1) = "-" Or Left(cleanText, 1) = "–" Or Left(cleanText, 1) = "—" Then
        cleanText = Trim(Mid(cleanText, 2))
    End If
    If Right(cleanText, 1) = "-" Or Right(cleanText, 1) = "–" Or Right(cleanText, 1) = "—" Then
        cleanText = Trim(Left(cleanText, Len(cleanText) - 1))
    End If

    ' Verifica se o que sobrou é alguma variação de "vereador"
    cleanText = LCase(Trim(cleanText))
    IsVereadorPattern = (cleanText = "vereador" Or cleanText = "vereadora")
End Function

Private Function IsAnexoPattern(text As String) As Boolean
    Dim cleanText As String
    cleanText = LCase(Trim(text))
    IsAnexoPattern = (cleanText = "anexo" Or cleanText = "anexos")
End Function

'================================================================================
' FORMAT DIANTE DO EXPOSTO - Formata "Diante do exposto" no início de parágrafos
'================================================================================
Private Sub FormatDianteDoExposto(doc As Document)
    On Error GoTo ErrorHandler

    If Not ValidateDocument(doc) Then Exit Sub

    Dim para As Paragraph
    Dim paraText As String
    Dim cleanText As String
    Dim formattedCount As Long
    formattedCount = 0

    ' Procura por parágrafos que começam com "Diante do exposto"
    Dim iterCounter As Long
    iterCounter = 0
    For Each para In doc.Paragraphs
        iterCounter = iterCounter + 1
        If iterCounter Mod 25 = 0 Then DoEvents ' Responsividade

        If Not HasVisualContent(para) Then
            ' Obtém o texto do parágrafo
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            cleanText = LCase(paraText)

            ' Verifica se começa com "diante do exposto"
            If Left(cleanText, 17) = "diante do exposto" Then
                ' Encontra a posição exata da frase (primeiros 17 caracteres)
                Dim targetRange As Range
                Set targetRange = para.Range
                targetRange.End = targetRange.Start + 17

                ' Aplica formatação: negrito e caixa alta
                With targetRange.Font
                    .Bold = True
                    .AllCaps = True
                    .Name = STANDARD_FONT
                    .size = STANDARD_FONT_SIZE
                End With

                formattedCount = formattedCount + 1
            End If
        End If
    Next para

    If formattedCount > 0 Then
        LogMessage "Formatação 'Diante do exposto': " & formattedCount & " ocorrências formatadas em negrito e caixa alta", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao formatar 'Diante do exposto': " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' FORMAT REQUEIRO PARAGRAPHS - Formata parágrafos que começam com "requeiro"
'================================================================================
Private Sub FormatRequeiroParagraphs(doc As Document)
    On Error GoTo ErrorHandler

    If Not ValidateDocument(doc) Then Exit Sub

    Dim para As Paragraph
    Dim paraText As String
    Dim cleanText As String
    Dim formattedCount As Long
    formattedCount = 0

    ' Procura por parágrafos que começam com "requeiro" (case insensitive)
    Dim reqCounter As Long
    reqCounter = 0
    For Each para In doc.Paragraphs
        reqCounter = reqCounter + 1
        If reqCounter Mod 25 = 0 Then DoEvents ' Responsividade

        If Not HasVisualContent(para) Then
            ' Obtém o texto do parágrafo (sem marca de parágrafo)
            paraText = para.Range.text
            If Right(paraText, 1) = vbCr Then
                paraText = Left(paraText, Len(paraText) - 1)
            End If
            paraText = Trim(paraText)
            cleanText = LCase(paraText)

            ' Verifica se começa com "requeiro" (8 caracteres)
            If Len(paraText) >= 8 Then
                If Left(cleanText, 8) = "requeiro" Then
                    ' Aplica formatação APENAS à palavra "requeiro": negrito e caixa alta
                    Dim wordRange As Range
                    Dim startPos As Long

                    ' Encontra a posição inicial do texto (após espaços/tabs)
                    Set wordRange = para.Range
                    startPos = wordRange.Start

                    ' Move para o início do texto visível
                    Do While startPos < wordRange.End
                        wordRange.Start = startPos
                        If Trim(Left(wordRange.text, 1)) <> "" Then Exit Do
                        startPos = startPos + 1
                    Loop

                    ' Seleciona apenas os 8 caracteres de "requeiro"
                    wordRange.End = wordRange.Start + 8

                    ' Aplica formatação apenas à palavra
                    With wordRange.Font
                        .Bold = True
                        .AllCaps = True
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                    End With

                    formattedCount = formattedCount + 1
                End If
            End If
        End If
    Next para

    If formattedCount > 0 Then
        LogMessage "Formatação 'Requeiro': " & formattedCount & " palavras formatadas em negrito e caixa alta", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao formatar parágrafos 'Requeiro': " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' FORMAT "POR TODAS AS RAZÕES" PARAGRAPHS - Formata "Por todas as razões aqui expostas" e "Pelas razões aqui expostas"
'================================================================================
Private Sub FormatPorTodasRazoesParagraphs(doc As Document)
    On Error GoTo ErrorHandler

    If Not ValidateDocument(doc) Then Exit Sub

    Dim para As Paragraph
    Dim paraText As String
    Dim cleanText As String
    Dim formattedCount As Long
    Dim wordRange As Range
    Dim phrase1Len As Long
    Dim phrase2Len As Long

    formattedCount = 0
    phrase1Len = 33 ' "por todas as razões aqui expostas"
    phrase2Len = 28 ' "pelas razões aqui expostas"

    ' Procura por parágrafos que começam com as frases (case insensitive)
    For Each para In doc.Paragraphs
        If Not HasVisualContent(para) Then
            ' Obtém o texto do parágrafo
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            cleanText = LCase(paraText)

            ' Verifica "por todas as razões aqui expostas"
            If Len(paraText) >= phrase1Len Then
                If Left(cleanText, phrase1Len) = "por todas as razões aqui expostas" Or _
                   Left(cleanText, phrase1Len) = "por todas as razoes aqui expostas" Then
                    Set wordRange = para.Range.Duplicate
                    wordRange.Collapse wdCollapseStart
                    wordRange.MoveEnd wdCharacter, phrase1Len

                    With wordRange.Font
                        .Bold = True
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                    End With

                    formattedCount = formattedCount + 1
                    GoTo NextPara
                End If
            End If

            ' Verifica "pelas razões aqui expostas"
            If Len(paraText) >= phrase2Len Then
                If Left(cleanText, phrase2Len) = "pelas razões aqui expostas" Or _
                   Left(cleanText, phrase2Len) = "pelas razoes aqui expostas" Then
                    Set wordRange = para.Range.Duplicate
                    wordRange.Collapse wdCollapseStart
                    wordRange.MoveEnd wdCharacter, phrase2Len

                    With wordRange.Font
                        .Bold = True
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                    End With

                    formattedCount = formattedCount + 1
                End If
            End If
        End If
NextPara:
    Next para

    If formattedCount > 0 Then
        LogMessage "Formatação 'Por todas as razões': " & formattedCount & " frases formatadas em negrito", LOG_LEVEL_INFO
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao formatar frases 'Por todas as razões': " & Err.Description, LOG_LEVEL_WARNING
End Sub



'================================================================================
' SUBROTINA PÚBLICA: ABRIR REPOSITÓRIO DO GITHUB
'================================================================================
Public Sub AbrirReadme()
    On Error GoTo ErrorHandler

    Const GITHUB_REPO_URL As String = "https://github.com/chrmsantos/chainsaw"

    ' Abre o repositório do GitHub no navegador padrão
    Application.StatusBar = "Abrindo repositório do GitHub..."

    ' Usa o comando Shell com o protocolo http:// para abrir no navegador padrão
    CreateObject("WScript.Shell").Run GITHUB_REPO_URL, 1, False

    ' Log da operação se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Repositório do GitHub aberto pelo usuário: " & GITHUB_REPO_URL, LOG_LEVEL_INFO
    End If

    Application.StatusBar = "Repositório aberto no navegador"

    Exit Sub

ErrorHandler:
    Application.StatusBar = "Erro ao abrir repositório"
    LogMessage "Erro ao abrir repositório do GitHub: " & Err.Description, LOG_LEVEL_ERROR

    ' Tenta método alternativo
    On Error Resume Next
    shell "explorer.exe """ & GITHUB_REPO_URL & """", vbNormalFocus
End Sub

'================================================================================
' SUBROTINA PÚBLICA: CONFIRMAR DESFAZIMENTO DA PADRONIZAÇÃO
'================================================================================
Public Sub ConfirmarDesfazerPadronizacao()
    On Error GoTo ErrorHandler

    ' Verifica se há um documento ativo
    Dim doc As Document
    Set doc = Nothing

    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        Exit Sub
    End If

    ' Verifica o número de ações disponíveis para desfazer
    Dim canUndo As Boolean
    canUndo = False

    On Error Resume Next
    canUndo = Application.CommandBars.ActionControl.enabled
    If Err.Number <> 0 Then canUndo = False
    On Error GoTo ErrorHandler

    ' Armazena informações antes do desfazer
    Dim beforeUndoCount As Long
    Dim docName As String
    Dim docPath As String

    beforeUndoCount = doc.Paragraphs.count
    docName = doc.Name
    docPath = doc.Path

    ' Executa o comando Desfazer (Undo)
    Application.StatusBar = "Desfazendo padronização..."
    On Error Resume Next
    doc.Undo
    On Error GoTo ErrorHandler

    ' Aguarda o Word processar o desfazer
    DoEvents

    ' Verifica se houve mudança no documento
    Dim afterUndoCount As Long
    afterUndoCount = doc.Paragraphs.count

    ' Calcula a diferença
    Dim changeCount As Long
    changeCount = Abs(beforeUndoCount - afterUndoCount)

    ' Cria mensagem informativa
    Dim undoMsg As String

    If changeCount > 0 Then
        undoMsg = "[<<] Padronização desfeita com sucesso!" & vbCrLf & vbCrLf & _
                  "[CHART] Alterações revertidas:" & vbCrLf & _
                  "   • Parágrafos afetados: " & changeCount & vbCrLf & vbCrLf & _
                  "[DIR] Documento:" & vbCrLf & _
                  "   " & docName & vbCrLf & vbCrLf & _
                  "[i] DICA: O backup da padronização permanece disponível." & vbCrLf & _
                  "   Use 'Abrir Pasta de Logs e Backups' para acessá-lo."
    Else
        undoMsg = "[<<] Desfazer executado!" & vbCrLf & vbCrLf & _
                  "[i] O documento foi revertido para o estado anterior." & vbCrLf & vbCrLf & _
                  "[DIR] Documento:" & vbCrLf & _
                  "   " & docName & vbCrLf & vbCrLf & _
                  "[i] DICA: O backup da padronização permanece disponível." & vbCrLf & _
                  "   Use 'Abrir Pasta de Logs e Backups' para acessá-lo."
    End If

    ' Exibe mensagem de confirmação
    MsgBox undoMsg, vbInformation, "CHAINSAW - Desfazer Padronização"

    ' Registra no log se estiver ativo
    If loggingEnabled Then
        LogMessage "Padronização desfeita pelo usuário - documento: " & docName, LOG_LEVEL_INFO
    End If

    Application.StatusBar = "Padronização desfeita"

    Exit Sub

ErrorHandler:
    Application.StatusBar = "Erro ao desfazer"

    ' Mensagem de erro genérica
    MsgBox "Não foi possível desfazer a operação." & vbCrLf & vbCrLf & _
           "[!] Possíveis causas:" & vbCrLf & _
           "   • Não há operações para desfazer" & vbCrLf & _
           "   • O documento foi fechado e reaberto" & vbCrLf & _
           "   • Limite de desfazer atingido" & vbCrLf & vbCrLf & _
           "[i] SOLUÇÃO: Restaure manualmente a partir do backup." & vbCrLf & _
           "   Use 'Abrir Pasta de Logs e Backups' para acessar os backups.", _
           vbExclamation, "CHAINSAW - Erro ao Desfazer"

    If loggingEnabled Then
        LogMessage "Erro ao desfazer padronização: " & Err.Description, LOG_LEVEL_WARNING
    End If
End Sub

'================================================================================
' SUBROTINA PÚBLICA: DESFAZER COM CONFIRMAÇÃO AUTOMÁTICA
' Esta sub pode ser chamada diretamente ou após o usuário usar Ctrl+Z
'================================================================================
Public Sub NotificarDesfazerPadronizacao()
    On Error Resume Next

    ' Verifica se há um documento ativo
    Dim doc As Document
    Set doc = ActiveDocument

    If doc Is Nothing Then Exit Sub

    ' Cria mensagem de confirmação simplificada
    Dim msg As String
    msg = "[<<] Padronização desfeita!" & vbCrLf & vbCrLf & _
          "[OK] Todas as alterações da última padronização foram revertidas." & vbCrLf & vbCrLf & _
          "[DIR] Documento: " & doc.Name & vbCrLf & vbCrLf & _
          "[SAVE] O backup continua disponível na pasta de backups." & vbCrLf & _
          "   Use 'Abrir Pasta de Logs e Backups' para acessá-lo."

    ' Exibe notificação
    MsgBox msg, vbInformation, "CHAINSAW - Operação Desfeita"

    ' Log se disponível
    If loggingEnabled Then
        LogMessage "Notificação de desfazer exibida para: " & doc.Name, LOG_LEVEL_INFO
    End If
End Sub

'================================================================================
' SISTEMA DE BACKUP
'================================================================================
Private Function CreateDocumentBackup(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Não faz backup se documento não foi realmente salvo (não existe no disco)
    If doc.Path = "" Or Not fso.FileExists(doc.FullName) Then
        LogMessage "Backup ignorado - documento não salvo", LOG_LEVEL_INFO
        CreateDocumentBackup = True
        Exit Function
    End If

    Dim backupFolder As String
    Dim docName As String
    Dim docExtension As String
    Dim timeStamp As String
    Dim backupFileName As String

    ' Usa a funcao que garante o diretorio de backup
    backupFolder = EnsureBackupDirectory(doc)

    ' Extrai nome e extensão do documento
    docName = fso.GetBaseName(doc.Name)
    docExtension = fso.GetExtensionName(doc.Name)

    ' Cria timestamp para o backup
    timeStamp = Format(Now, "yyyy-mm-dd_HHmmss")

    ' Nome do arquivo de backup
    backupFileName = docName & "_backup_" & timeStamp & "." & docExtension
    backupFilePath = backupFolder & "\" & backupFileName

    ' Protege contra conflito: exclui arquivo pre-existente com mesmo nome
    If fso.FileExists(backupFilePath) Then
        fso.DeleteFile backupFilePath, True
        LogMessage "Backup anterior com mesmo nome excluido: " & backupFileName, LOG_LEVEL_INFO
    End If

    ' Salva uma cópia do documento como backup
    Application.StatusBar = "Criando backup..."

    ' Salva o documento atual primeiro para garantir que está atualizado
    doc.Save

    ' Cria uma cópia do arquivo usando FileSystemObject
    fso.CopyFile doc.FullName, backupFilePath, True

    ' Limpa backups antigos se necessário
    CleanOldBackups backupFolder, docName

    LogMessage "Backup criado com sucesso: " & backupFileName, LOG_LEVEL_INFO
    Application.StatusBar = "Backup criado"

    CreateDocumentBackup = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao criar backup: " & Err.Description, LOG_LEVEL_ERROR
    CreateDocumentBackup = False
End Function

'================================================================================
' RESTAURAR BACKUP - Descarta documento atual e restaura backup
'================================================================================
Public Sub RestaurarBackup()
    On Error GoTo ErrorHandler

    Dim doc As Document
    Set doc = ActiveDocument

    If doc Is Nothing Then
        MsgBox "Nenhum documento ativo para restaurar.", vbExclamation, "CHAINSAW - Restaurar Backup"
        Exit Sub
    End If

    ' Verifica se existe backup para este documento
    If backupFilePath = "" Or Not CreateObject("Scripting.FileSystemObject").FileExists(backupFilePath) Then
        MsgBox "Nenhum backup disponivel para este documento." & vbCrLf & vbCrLf & _
               "[i] O backup e criado apenas apos a primeira execucao de PadronizarDocumentoMain.", _
               vbExclamation, "CHAINSAW - Restaurar Backup"
        Exit Sub
    End If

    ' Confirma com usuario
    Dim confirmMsg As String
    confirmMsg = "[?] Deseja restaurar o backup do documento?" & vbCrLf & vbCrLf & _
                 "[!] ATENCAO: O documento atual sera descartado!" & vbCrLf & vbCrLf & _
                 "[DIR] Documento atual: " & doc.Name & vbCrLf & _
                 "[DIR] Backup: " & CreateObject("Scripting.FileSystemObject").GetFileName(backupFilePath)

    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "CHAINSAW - Confirmar Restauracao") <> vbYes Then
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim originalPath As String
    Dim originalName As String
    Dim discardedPath As String
    Dim timeStamp As String

    originalPath = doc.FullName
    originalName = doc.Name

    ' Cria timestamp para arquivo descartado
    timeStamp = Format(Now, "yyyy-mm-dd_HHmmss")

    ' Nome do arquivo descartado: nome_discarded_timestamp.ext
    Dim baseName As String
    Dim extension As String
    baseName = fso.GetBaseName(originalName)
    extension = fso.GetExtensionName(originalName)

    discardedPath = fso.GetParentFolderName(originalPath) & "\" & _
                    baseName & "_discarded_" & timeStamp & "." & extension

    ' Protege contra conflito: exclui arquivo pre-existente
    If fso.FileExists(discardedPath) Then
        fso.DeleteFile discardedPath, True
    End If

    ' Salva documento atual como _discarded
    Application.StatusBar = "Salvando documento descartado..."
    doc.SaveAs2 discardedPath

    ' Fecha o documento descartado
    doc.Close SaveChanges:=False

    ' Protege contra conflito no caminho original
    If fso.FileExists(originalPath) Then
        fso.DeleteFile originalPath, True
    End If

    ' Copia backup para o local original
    Application.StatusBar = "Restaurando backup..."
    fso.CopyFile backupFilePath, originalPath, True

    ' Abre o backup restaurado
    Application.Documents.Open originalPath

    Application.StatusBar = "Backup restaurado com sucesso"

    MsgBox "[OK] Backup restaurado com sucesso!" & vbCrLf & vbCrLf & _
           "[DIR] Documento descartado salvo em:" & vbCrLf & _
           "   " & discardedPath & vbCrLf & vbCrLf & _
           "[DIR] Backup restaurado:" & vbCrLf & _
           "   " & originalPath, _
           vbInformation, "CHAINSAW - Backup Restaurado"

    Exit Sub

ErrorHandler:
    Application.StatusBar = "Erro ao restaurar backup"
    MsgBox "[ERRO] Falha ao restaurar backup:" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "[i] O documento pode estar em estado inconsistente." & vbCrLf & _
           "   Verifique manualmente a pasta de backups.", _
           vbCritical, "CHAINSAW - Erro na Restauracao"
End Sub

'================================================================================
' LIMPEZA DE BACKUPS ANTIGOS
'================================================================================
Private Sub CleanOldBackups(backupFolder As String, docBaseName As String)
    On Error GoTo CleanExit

    If MAX_BACKUP_FILES < 1 Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(backupFolder) Then GoTo CleanExit

    Dim folder As Object
    Set folder = fso.GetFolder(backupFolder)

    Dim items As Object
    Set items = CreateObject("System.Collections.ArrayList")

    Dim fileItem As Object
    Dim prefix As String
    prefix = LCase(docBaseName & "_backup_")

    For Each fileItem In folder.Files
        If Left(LCase(fileItem.Name), Len(prefix)) = prefix Then
            items.Add Format(fileItem.DateLastModified, "yyyymmddHHMMSS") & "|" & fileItem.Path
        End If
    Next fileItem

    If items.count <= MAX_BACKUP_FILES Then GoTo CleanExit

    items.Sort
    items.Reverse

    Dim idx As Long
    For idx = MAX_BACKUP_FILES To items.count - 1
        Dim parts() As String
        parts = Split(items(idx), "|")
        On Error Resume Next
        fso.DeleteFile parts(1), True
        If Err.Number <> 0 Then
            If loggingEnabled Then
                LogMessage "Failed to delete old backup: " & parts(1) & " - " & Err.Description, LOG_LEVEL_WARNING
            End If
            Err.Clear
        Else
            If loggingEnabled Then
                LogMessage "Old backup removed: " & parts(1), LOG_LEVEL_INFO
            End If
        End If
        On Error GoTo CleanExit
    Next idx

CleanExit:
    On Error Resume Next
    Set items = Nothing
    Set folder = Nothing
    Set fso = Nothing
End Sub

'================================================================================
' LIMPEZA DE ESPAÇOS MÚLTIPLOS
'================================================================================
Private Function CleanMultipleSpaces(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Limpando espaços..."

    Dim rng As Range
    Dim spacesRemoved As Long
    Dim totalOperations As Long

    ' SUPER OTIMIZADO: Operações consolidadas em uma única configuração Find
    Set rng = doc.Range

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        ' OTIMIZAÇÃO 1: Remove espaços múltiplos (2 ou mais) em uma única operação
        ' Usa um loop otimizado que reduz progressivamente os espaços
        Do
            .text = "  "  ' Dois espaços
            .Replacement.text = " "  ' Um espaço

            Dim currentReplaceCount As Long
            currentReplaceCount = 0

            ' Executa até não encontrar mais duplos
            Do While .Execute(Replace:=True)
                currentReplaceCount = currentReplaceCount + 1
                spacesRemoved = spacesRemoved + 1
                ' Proteção otimizada - verifica a cada 200 operações
                If currentReplaceCount Mod 200 = 0 Then
                    DoEvents
                    If spacesRemoved > 2000 Then Exit Do
                End If
            Loop

            totalOperations = totalOperations + 1
            ' Se não encontrou mais duplos ou atingiu limite, para
            If currentReplaceCount = 0 Or totalOperations > 10 Then Exit Do
        Loop
    End With

    ' OTIMIZAÇÃO 2: Operações de limpeza de quebras de linha consolidadas
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False  ' Usar Find/Replace simples para compatibilidade

        ' Remove múltiplos espaços antes de quebras - método iterativo
        .text = "  ^p"  ' 2 espaços seguidos de quebra
        .Replacement.text = " ^p"  ' 1 espaço seguido de quebra
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop

        ' Segunda passada para garantir limpeza completa
        .text = " ^p"  ' Espaço antes de quebra
        .Replacement.text = "^p"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop

        ' Remove múltiplos espaços depois de quebras - método iterativo
        .text = "^p  "  ' Quebra seguida de 2 espaços
        .Replacement.text = "^p "  ' Quebra seguida de 1 espaço
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
    End With

    ' OTIMIZAÇÃO 3: Limpeza de tabs consolidada e otimizada
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False  ' Usar Find/Replace simples

        ' Remove múltiplos tabs iterativamente
        .text = "^t^t"  ' 2 tabs
        .Replacement.text = "^t"  ' 1 tab
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop

        ' Converte tabs para espaços
        .text = "^t"
        .Replacement.text = " "
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
    End With

    ' OTIMIZAÇÃO 4: Verificação final ultra-rápida de espaços duplos remanescentes
    Set rng = doc.Range
    With rng.Find
        .text = "  "
        .Replacement.text = " "
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop  ' Mais rápido que wdFindContinue

        Dim finalCleanCount As Long
        Do While .Execute(Replace:=True) And finalCleanCount < 100
            finalCleanCount = finalCleanCount + 1
            spacesRemoved = spacesRemoved + 1
        Loop
    End With

    ' PROTEÇÃO ESPECÍFICA: Garante espaço após CONSIDERANDO
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchCase = False
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False

        ' Corrige CONSIDERANDO grudado com a próxima palavra
        .text = "CONSIDERANDOa"
        .Replacement.text = "CONSIDERANDO a"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop

        .text = "CONSIDERANDOe"
        .Replacement.text = "CONSIDERANDO e"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop

        .text = "CONSIDERANDOo"
        .Replacement.text = "CONSIDERANDO o"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop

        .text = "CONSIDERANDOq"
        .Replacement.text = "CONSIDERANDO q"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop
    End With

    ' Marca documento como modificado se houve limpeza
    If spacesRemoved > 0 Then documentDirty = True

    LogMessage "Limpeza de espacos concluida: " & spacesRemoved & " correcoes aplicadas (com protecao CONSIDERANDO)", LOG_LEVEL_INFO
    CleanMultipleSpaces = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na limpeza de espaços múltiplos: " & Err.Description, LOG_LEVEL_WARNING
    CleanMultipleSpaces = False ' Não falha o processo por isso
End Function

'================================================================================
' LIMITAÇÃO DE LINHAS VAZIAS SEQUENCIAIS
'================================================================================
Private Function LimitSequentialEmptyLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Controlando linhas..."

    ' IDENTIFICAÇÃO DO SEGUNDO PARÁGRAFO PARA PROTEÇÃO
    Dim secondParaIndex As Long
    secondParaIndex = GetSecondParagraphIndex(doc)

    ' SUPER OTIMIZADO: Usa Find/Replace com wildcard para operação muito mais rápida
    Dim rng As Range
    Dim linesRemoved As Long
    Dim totalReplaces As Long
    Dim passCount As Long

    passCount = 1 ' Inicializa contador de passadas

    Set rng = doc.Range

    ' MÉTODO ULTRA-RÁPIDO: Remove múltiplas quebras consecutivas usando wildcard
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False  ' Usar Find/Replace simples para compatibilidade

        ' Remove múltiplas quebras consecutivas iterativamente
        .text = "^p^p^p^p"  ' 4 quebras
        .Replacement.text = "^p^p"  ' 2 quebras

        Do While .Execute(Replace:=True)
            linesRemoved = linesRemoved + 1
            totalReplaces = totalReplaces + 1
            If totalReplaces > 500 Then Exit Do
            If linesRemoved Mod 50 = 0 Then DoEvents
        Loop

        ' Remove 3 quebras -> 2 quebras
        .text = "^p^p^p"  ' 3 quebras
        .Replacement.text = "^p^p"  ' 2 quebras

        Do While .Execute(Replace:=True)
            linesRemoved = linesRemoved + 1
            totalReplaces = totalReplaces + 1
            If totalReplaces > 500 Then Exit Do
            If linesRemoved Mod 50 = 0 Then DoEvents
        Loop
    End With

    ' SEGUNDA PASSADA: Remove quebras duplas restantes (2 quebras -> 1 quebra)
    If totalReplaces > 0 Then passCount = passCount + 1

    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindContinue

        ' Converte quebras duplas em quebras simples
        .text = "^p^p^p"  ' 3 quebras
        .Replacement.text = "^p^p"  ' 2 quebras

        Dim secondPassCount As Long
        Do While .Execute(Replace:=True) And secondPassCount < 200
            secondPassCount = secondPassCount + 1
            linesRemoved = linesRemoved + 1
        Loop
    End With

    ' VERIFICAÇÃO FINAL: Garantir que não há mais de 1 linha vazia consecutiva
    If secondPassCount > 0 Then passCount = passCount + 1

    ' Método híbrido: Find/Replace para casos simples + loop apenas se necessário
    Set rng = doc.Range
    With rng.Find
        .text = "^p^p^p"  ' 3 quebras (2 linhas vazias + conteúdo)
        .Replacement.text = "^p^p"  ' 2 quebras (1 linha vazia + conteúdo)
        .MatchWildcards = False

        Dim finalPassCount As Long
        Do While .Execute(Replace:=True) And finalPassCount < 100
            finalPassCount = finalPassCount + 1
            linesRemoved = linesRemoved + 1
        Loop
    End With

    If finalPassCount > 0 Then passCount = passCount + 1

    ' FALLBACK OTIMIZADO: Se ainda há problemas, usa método tradicional limitado
    If finalPassCount >= 100 Then
        passCount = passCount + 1 ' Incrementa para o fallback

        Dim para As Paragraph
        Dim i As Long
        Dim emptyLineCount As Long
        Dim paraText As String
        Dim fallbackRemoved As Long

        i = 1
        emptyLineCount = 0

        Do While i <= doc.Paragraphs.count And fallbackRemoved < 50
            Set para = doc.Paragraphs(i)
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

            ' Verifica se o parágrafo está vazio
            If paraText = "" And Not HasVisualContent(para) Then
                emptyLineCount = emptyLineCount + 1

                ' Se já temos mais de 1 linha vazia consecutiva, remove esta
                If emptyLineCount > 1 Then
                    para.Range.Delete
                    fallbackRemoved = fallbackRemoved + 1
                    linesRemoved = linesRemoved + 1
                    ' Não incrementa i pois removemos um parágrafo
                Else
                    i = i + 1
                End If
            Else
                ' Se encontrou conteúdo, reseta o contador
                emptyLineCount = 0
                i = i + 1
            End If

            ' Responsividade e proteção otimizadas
            If fallbackRemoved Mod 10 = 0 Then DoEvents
            If i > 500 Then Exit Do ' Proteção adicional
        Loop
    End If

    LogMessage "Controle de linhas vazias concluído em " & passCount & " passada(s): " & linesRemoved & " linhas excedentes removidas (máximo 1 sequencial)", LOG_LEVEL_INFO
    LimitSequentialEmptyLines = True
    Exit Function

ErrorHandler:
    LogMessage "Erro no controle de linhas vazias: " & Err.Description, LOG_LEVEL_WARNING
    LimitSequentialEmptyLines = False ' Não falha o processo por isso
End Function

'================================================================================
' CONFIGURE DOCUMENT VIEW - CONFIGURAÇÃO DE VISUALIZAÇÃO
'================================================================================
Private Function ConfigureDocumentView(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Configurando visualização..."

    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow

    ' Configura APENAS o zoom para 110% - todas as outras configurações são preservadas
    With docWindow.View
        .Zoom.Percentage = 110
        ' NÃO altera mais o tipo de visualização - preserva o original
    End With

    ' Remove configurações que alteravam configurações globais do Word
    ' Estas configurações são agora preservadas do estado original

    LogMessage "Visualização configurada: zoom definido para 110%, demais configurações preservadas"
    ConfigureDocumentView = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao configurar visualização: " & Err.Description, LOG_LEVEL_WARNING
    ConfigureDocumentView = False ' Não falha o processo por isso
End Function

'================================================================================
' REMOÇÃO DE REALCES E BORDAS - REMOVE HIGHLIGHTING AND BORDERS
'================================================================================
Private Function RemoveAllHighlightsAndBorders(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Removendo realces e bordas..."

    Dim para As Paragraph
    Dim highlightCount As Long
    Dim borderCount As Long
    Dim processedCount As Long

    highlightCount = 0
    borderCount = 0
    processedCount = 0

    ' Remove realce de todo o documento primeiro (mais rápido)
    On Error Resume Next
    doc.Range.HighlightColorIndex = 0 ' Remove realce
    If Err.Number = 0 Then
        highlightCount = 1
        LogMessage "Realce removido do documento completo", LOG_LEVEL_INFO
    End If
    Err.Clear
    On Error GoTo ErrorHandler

    ' Remove bordas de todos os parágrafos
    For Each para In doc.Paragraphs
        On Error Resume Next

        ' Remove bordas do parágrafo
        With para.Borders
            .Enable = False
        End With

        If Err.Number = 0 Then
            borderCount = borderCount + 1
        End If
        Err.Clear

        processedCount = processedCount + 1

        ' Responsividade
        If processedCount Mod 50 = 0 Then
            DoEvents
            Application.StatusBar = "Removendo bordas: " & processedCount & " de " & doc.Paragraphs.count
        End If

        On Error GoTo ErrorHandler
    Next para

    LogMessage "Realces e bordas removidos: " & highlightCount & " realces, " & borderCount & " parágrafos com bordas", LOG_LEVEL_INFO
    RemoveAllHighlightsAndBorders = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover realces e bordas: " & Err.Description, LOG_LEVEL_WARNING
    RemoveAllHighlightsAndBorders = False ' Não falha o processo por isso
End Function

'================================================================================
' REMOÇÃO DE PÁGINAS VAZIAS NO FINAL - REMOVE EMPTY PAGES AT END
'================================================================================
Private Function RemoveEmptyPagesAtEnd(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Verificando páginas vazias no final..."

    ' Verifica se há páginas vazias no final
    Dim totalPages As Long
    Dim lastPageRange As Range
    Dim lastPageText As String
    Dim pagesRemoved As Long
    Dim maxAttempts As Long
    Dim attemptCount As Long

    pagesRemoved = 0
    maxAttempts = 5 ' Máximo de tentativas para evitar loop infinito
    attemptCount = 0

    Do
        attemptCount = attemptCount + 1

        ' Obtém número total de páginas
        On Error Resume Next
        totalPages = doc.ComputeStatistics(wdStatisticPages)
        If Err.Number <> 0 Then
            LogMessage "Não foi possível obter estatísticas de páginas: " & Err.Description, LOG_LEVEL_WARNING
            Err.Clear
            Exit Do
        End If
        Err.Clear
        On Error GoTo ErrorHandler

        ' Se há apenas 1 página, não remove nada
        If totalPages <= 1 Then
            Exit Do
        End If

        ' Obtém o range da última página
        Set lastPageRange = doc.Range
        lastPageRange.Start = doc.Range.End - 1
        lastPageRange.End = doc.Range.End

        ' Expande para incluir toda a última página
        lastPageRange.Expand wdParagraph

        ' Obtém texto da última página (últimos parágrafos)
        Dim lastParaIndex As Long
        Dim para As Paragraph
        Dim hasContent As Boolean

        hasContent = False
        lastParaIndex = doc.Paragraphs.count

        ' Verifica os últimos parágrafos em busca de conteúdo
        Dim checkCount As Long
        checkCount = 0

        Do While lastParaIndex > 0 And checkCount < 20
            Set para = doc.Paragraphs(lastParaIndex)
            lastPageText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

            ' Se encontrou conteúdo de texto
            If Len(lastPageText) > 0 Then
                hasContent = True
                Exit Do
            End If

            ' Se encontrou imagem ou objeto
            If para.Range.InlineShapes.count > 0 Then
                hasContent = True
                Exit Do
            End If

            lastParaIndex = lastParaIndex - 1
            checkCount = checkCount + 1
        Loop

        ' Se a última página NÃO tem conteúdo, remove parágrafos vazios do final
        If Not hasContent Then
            Dim removedInThisPass As Long
            removedInThisPass = 0

            ' Remove parágrafos vazios do final (mínimo necessário)
            lastParaIndex = doc.Paragraphs.count
            Do While lastParaIndex > 0 And removedInThisPass < 10
                Set para = doc.Paragraphs(lastParaIndex)
                lastPageText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

                ' Se é parágrafo vazio sem conteúdo visual
                If Len(lastPageText) = 0 And para.Range.InlineShapes.count = 0 Then
                    para.Range.Delete
                    removedInThisPass = removedInThisPass + 1
                    pagesRemoved = pagesRemoved + 1
                    lastParaIndex = lastParaIndex - 1
                Else
                    ' Encontrou conteúdo, para de remover
                    Exit Do
                End If

                ' Proteção contra loop infinito
                If removedInThisPass Mod 3 = 0 Then DoEvents
            Loop

            ' Se não removeu nada nesta passada, termina
            If removedInThisPass = 0 Then
                Exit Do
            End If
        Else
            ' Última página tem conteúdo, não remove
            Exit Do
        End If

        ' Proteção contra tentativas excessivas
        If attemptCount >= maxAttempts Then
            LogMessage "Atingido número máximo de tentativas de remoção de páginas vazias", LOG_LEVEL_WARNING
            Exit Do
        End If
    Loop

    If pagesRemoved > 0 Then
        LogMessage "Páginas vazias removidas do final: " & pagesRemoved & " parágrafo(s) vazio(s) removido(s)", LOG_LEVEL_INFO
    Else
        LogMessage "Nenhuma página vazia no final do documento", LOG_LEVEL_INFO
    End If

    RemoveEmptyPagesAtEnd = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover páginas vazias: " & Err.Description, LOG_LEVEL_WARNING
    RemoveEmptyPagesAtEnd = False ' Não falha o processo por isso
End Function

'================================================================================
' IMAGE PROTECTION SYSTEM - SISTEMA DE PROTEÇÃO DE IMAGENS
'================================================================================

'================================================================================
' BACKUP DE IMAGENS
'================================================================================
Private Function BackupAllImages(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Protegendo imagens..."

    imageCount = 0
    ReDim savedImages(0)

    Dim para As Paragraph
    Dim i As Long
    Dim j As Long
    Dim shape As InlineShape
    Dim tempImageInfo As ImageInfo

    ' Conta todas as imagens primeiro (com DoEvents para responsividade)
    Dim totalImages As Long
    For i = 1 To doc.Paragraphs.count
        If i Mod 30 = 0 Then DoEvents ' Responsividade
        Set para = doc.Paragraphs(i)
        totalImages = totalImages + para.Range.InlineShapes.count
    Next i

    ' Adiciona shapes flutuantes
    totalImages = totalImages + doc.Shapes.count

    ' Redimensiona array se necessario
    If totalImages > 0 Then
        ReDim savedImages(totalImages - 1)

        ' Backup de imagens inline - apenas propriedades criticas
        For i = 1 To doc.Paragraphs.count
            If i Mod 30 = 0 Then DoEvents ' Responsividade
            Set para = doc.Paragraphs(i)

            For j = 1 To para.Range.InlineShapes.count
                Set shape = para.Range.InlineShapes(j)

                ' Salva apenas propriedades essenciais para protecao
                With tempImageInfo
                    .paraIndex = i
                    .ImageIndex = j
                    .ImageType = "Inline"
                    .Position = shape.Range.Start
                    .Width = shape.Width
                    .Height = shape.Height
                    Set .AnchorRange = shape.Range.Duplicate
                    .ImageData = "InlineShape_Protected"
                End With

                savedImages(imageCount) = tempImageInfo
                imageCount = imageCount + 1

                ' Evita overflow
                If imageCount >= UBound(savedImages) + 1 Then Exit For
            Next j

            ' Evita overflow
            If imageCount >= UBound(savedImages) + 1 Then Exit For
        Next i

        ' Backup de shapes flutuantes - apenas propriedades críticas
        Dim floatingShape As shape
        For i = 1 To doc.Shapes.count
            Set floatingShape = doc.Shapes(i)

            If floatingShape.Type = msoPicture Then
                ' Redimensiona array se necessário
                If imageCount >= UBound(savedImages) + 1 Then
                    ReDim Preserve savedImages(imageCount)
                End If

                With tempImageInfo
                    .paraIndex = -1 ' Indica que é flutuante
                    .ImageIndex = i
                    .ImageType = "Floating"
                    .WrapType = floatingShape.WrapFormat.Type
                    .Width = floatingShape.Width
                    .Height = floatingShape.Height
                    .LeftPosition = floatingShape.Left
                    .TopPosition = floatingShape.Top
                    .ImageData = "FloatingShape_Protected"
                End With

                savedImages(imageCount) = tempImageInfo
                imageCount = imageCount + 1
            End If
        Next i
    End If

    LogMessage "Backup de propriedades de imagens concluído: " & imageCount & " imagens catalogadas"
    BackupAllImages = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao fazer backup de propriedades de imagens: " & Err.Description, LOG_LEVEL_WARNING
    BackupAllImages = False
End Function

'================================================================================
' RESTAURAÇÃO DE IMAGENS
'================================================================================
Private Function RestoreAllImages(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If imageCount = 0 Then
        RestoreAllImages = True
        Exit Function
    End If

    Application.StatusBar = "Verificando integridade das imagens..."

    Dim i As Long
    Dim verifiedCount As Long
    Dim correctedCount As Long

    For i = 0 To imageCount - 1
        On Error Resume Next

        With savedImages(i)
            If .ImageType = "Inline" Then
                ' Verifica se a imagem inline ainda existe na posição esperada
                If .paraIndex <= doc.Paragraphs.count Then
                    Dim para As Paragraph
                    Set para = doc.Paragraphs(.paraIndex)

                    ' Se ainda há imagens inline no parágrafo, considera verificada
                    If para.Range.InlineShapes.count > 0 Then
                        verifiedCount = verifiedCount + 1
                    End If
                End If

            ElseIf .ImageType = "Floating" Then
                ' Verifica e corrige propriedades de shapes flutuantes se ainda existem
                If .ImageIndex <= doc.Shapes.count Then
                    Dim targetShape As shape
                    Set targetShape = doc.Shapes(.ImageIndex)

                    ' Verifica se as propriedades foram alteradas e corrige se necessário
                    Dim needsCorrection As Boolean
                    needsCorrection = False

                    If Abs(targetShape.Width - .Width) > 1 Then needsCorrection = True
                    If Abs(targetShape.Height - .Height) > 1 Then needsCorrection = True
                    If Abs(targetShape.Left - .LeftPosition) > 1 Then needsCorrection = True
                    If Abs(targetShape.Top - .TopPosition) > 1 Then needsCorrection = True

                    If needsCorrection Then
                        ' Restaura propriedades originais
                        With targetShape
                            .Width = savedImages(i).Width
                            .Height = savedImages(i).Height
                            .Left = savedImages(i).LeftPosition
                            .Top = savedImages(i).TopPosition
                            .WrapFormat.Type = savedImages(i).WrapType
                        End With
                        correctedCount = correctedCount + 1
                    End If

                    verifiedCount = verifiedCount + 1
                End If
            End If
        End With

        On Error GoTo ErrorHandler
    Next i

    If correctedCount > 0 Then
        LogMessage "Verificação de imagens concluída: " & verifiedCount & " verificadas, " & correctedCount & " corrigidas"
    Else
        LogMessage "Verificação de imagens concluída: " & verifiedCount & " imagens íntegras"
    End If

    RestoreAllImages = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao verificar imagens: " & Err.Description, LOG_LEVEL_WARNING
    RestoreAllImages = False
End Function

'================================================================================
' FORMAT IMAGE PARAGRAPHS INDENTS - Formata recuos de parágrafos com imagens
'================================================================================
Private Function FormatImageParagraphsIndents(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim formattedCount As Long
    formattedCount = 0

    ' Percorre todos os paragrafos
    Dim imgCounter As Long
    imgCounter = 0
    For Each para In doc.Paragraphs
        imgCounter = imgCounter + 1
        If imgCounter Mod 30 = 0 Then DoEvents ' Responsividade

        ' Verifica se o paragrafo contem imagens inline
        If para.Range.InlineShapes.count > 0 Then
            ' Zera o recuo a esquerda e centraliza
            With para.Format
                .leftIndent = 0
                .firstLineIndent = 0
                .alignment = wdAlignParagraphCenter
            End With
            formattedCount = formattedCount + 1
        End If
    Next para

    If formattedCount > 0 Then
        LogMessage "Recuos de parágrafos com imagens formatados: " & formattedCount & " parágrafos", LOG_LEVEL_INFO
    End If

    FormatImageParagraphsIndents = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao formatar recuos de imagens: " & Err.Description, LOG_LEVEL_WARNING
    FormatImageParagraphsIndents = False
End Function

'================================================================================
' BACKUP LIST FORMATS - Salva formatacoes de lista antes do processamento
'================================================================================
Private Function BackupListFormats(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim i As Long
    Dim tempListInfo As ListFormatInfo

    listFormatCount = 0
    ReDim savedListFormats(0)

    ' Conta quantos paragrafos tem formatacao de lista (com DoEvents)
    Dim totalLists As Long
    Dim countIter As Long
    totalLists = 0
    countIter = 0
    For Each para In doc.Paragraphs
        countIter = countIter + 1
        If countIter Mod 30 = 0 Then DoEvents ' Responsividade
        If para.Range.ListFormat.ListType <> 0 Then
            totalLists = totalLists + 1
        End If
    Next para

    If totalLists = 0 Then
        LogMessage "Nenhuma lista encontrada no documento", LOG_LEVEL_INFO
        BackupListFormats = True
        Exit Function
    End If

    ' Aloca array com tamanho adequado
    ReDim savedListFormats(totalLists - 1)

    ' Salva informacoes de cada paragrafo com lista (com DoEvents)
    Dim saveIter As Long
    saveIter = 0
    i = 1
    For Each para In doc.Paragraphs
        saveIter = saveIter + 1
        If saveIter Mod 30 = 0 Then DoEvents ' Responsividade

        If para.Range.ListFormat.ListType <> 0 Then
            With tempListInfo
                .paraIndex = i
                .HasList = True
                .ListType = para.Range.ListFormat.ListType

                ' Salva o nivel da lista se aplicavel
                On Error Resume Next
                .ListLevelNumber = para.Range.ListFormat.ListLevelNumber
                If Err.Number <> 0 Then
                    .ListLevelNumber = 1
                    Err.Clear
                End If
                On Error GoTo ErrorHandler

                ' Salva a string da lista (marcador ou numero)
                On Error Resume Next
                .ListString = para.Range.ListFormat.ListString
                If Err.Number <> 0 Then
                    .ListString = ""
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            End With

            savedListFormats(listFormatCount) = tempListInfo
            listFormatCount = listFormatCount + 1

            If listFormatCount >= UBound(savedListFormats) + 1 Then Exit For
        End If
        i = i + 1
    Next para

    LogMessage "Formatacoes de lista salvas: " & listFormatCount & " paragrafos com lista", LOG_LEVEL_INFO
    BackupListFormats = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao salvar formatacoes de lista: " & Err.Description, LOG_LEVEL_WARNING
    BackupListFormats = False
End Function

'================================================================================
' RESTORE LIST FORMATS - Restaura formatacoes de lista apos o processamento
'================================================================================
Private Function RestoreListFormats(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If listFormatCount = 0 Then
        RestoreListFormats = True
        Exit Function
    End If

    Dim i As Long
    Dim restoredCount As Long
    Dim para As Paragraph

    restoredCount = 0

    For i = 0 To listFormatCount - 1
        On Error Resume Next

        With savedListFormats(i)
            If .HasList And .paraIndex <= doc.Paragraphs.count Then
                Set para = doc.Paragraphs(.paraIndex)

                ' Remove qualquer formatacao de lista existente primeiro
                para.Range.ListFormat.RemoveNumbers

                ' Aplica a formatacao de lista original
                Select Case .ListType
                    Case 2 ' wdListBullet
                        ' Lista com marcadores
                        para.Range.ListFormat.ApplyBulletDefault

                    Case 3, 4 ' wdListSimpleNumbering, wdListListNumOnly
                        ' Lista numerada simples
                        para.Range.ListFormat.ApplyNumberDefault

                    Case 5 ' wdListMixedNumbering
                        ' Lista com numeracao mista
                        para.Range.ListFormat.ApplyNumberDefault

                    Case 6 ' wdListOutlineNumbering
                        ' Lista com numeracao de topicos
                        para.Range.ListFormat.ApplyOutlineNumberDefault

                    Case Else
                        ' Tenta aplicar formatacao padrao
                        If InStr(.ListString, ".") > 0 Or IsNumeric(Left(.ListString, 1)) Then
                            para.Range.ListFormat.ApplyNumberDefault
                        Else
                            para.Range.ListFormat.ApplyBulletDefault
                        End If
                End Select

                ' Tenta restaurar o nivel da lista
                If .ListLevelNumber > 0 And .ListLevelNumber <= 9 Then
                    para.Range.ListFormat.ListLevelNumber = .ListLevelNumber
                End If

                If Err.Number = 0 Then
                    restoredCount = restoredCount + 1
                Else
                    LogMessage "Aviso: Nao foi possivel restaurar lista no paragrafo " & .paraIndex & ": " & Err.Description, LOG_LEVEL_WARNING
                    Err.Clear
                End If
            End If
        End With

        On Error GoTo ErrorHandler
    Next i

    If restoredCount > 0 Then
        LogMessage "Formatacoes de lista restauradas: " & restoredCount & " paragrafos", LOG_LEVEL_INFO
    End If

    ' Limpa o array
    ReDim savedListFormats(0)
    listFormatCount = 0

    RestoreListFormats = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao restaurar formatacoes de lista: " & Err.Description, LOG_LEVEL_WARNING
    RestoreListFormats = False
End Function

'================================================================================
' FORMAT NUMBERED PARAGRAPHS INDENT - Aplica recuo em paragrafos iniciados com numero
'================================================================================
Private Function FormatNumberedParagraphsIndent(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim firstChar As String
    Dim formattedCount As Long
    Dim defaultIndent As Single

    formattedCount = 0

    ' Recuo padrao de lista numerada (36 pontos = 1.27 cm)
    defaultIndent = 36

    ' Percorre todos os paragrafos
    For Each para In doc.Paragraphs
        paraText = Trim(para.Range.text)

        ' Verifica se o paragrafo nao esta vazio
        If Len(paraText) > 0 Then
            ' Pega o primeiro caractere
            firstChar = Left(paraText, 1)

            ' Verifica se o primeiro caractere e um algarismo (0-9)
            If IsNumeric(firstChar) Then
                ' Verifica se o paragrafo nao tem formatacao de lista ja aplicada
                If para.Range.ListFormat.ListType = 0 Then
                    ' Aplica o recuo a esquerda igual ao de uma lista numerada
                    With para.Format
                        .leftIndent = defaultIndent
                        .firstLineIndent = 0
                    End With
                    formattedCount = formattedCount + 1
                End If
            End If
        End If
    Next para

    If formattedCount > 0 Then
        LogMessage "Paragrafos iniciados com numero formatados com recuo de lista: " & formattedCount, LOG_LEVEL_INFO
    End If

    FormatNumberedParagraphsIndent = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao formatar recuos de paragrafos numerados: " & Err.Description, LOG_LEVEL_WARNING
    FormatNumberedParagraphsIndent = False
End Function

'================================================================================
' FORMAT BULLETED PARAGRAPHS INDENT - Aplica recuo em paragrafos com marcadores
'================================================================================
Private Function FormatBulletedParagraphsIndent(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim firstChar As String
    Dim formattedCount As Long
    Dim defaultIndent As Single
    Dim i As Long

    formattedCount = 0

    ' Recuo padrao de lista com marcadores (36 pontos = 1.27 cm)
    defaultIndent = 36

    ' Array com os marcadores mais comuns
    Dim bulletMarkers() As String
    bulletMarkers = Split("*,-,>,+,~", ",")

    ' Percorre todos os paragrafos
    Dim bulletCounter As Long
    bulletCounter = 0
    For Each para In doc.Paragraphs
        bulletCounter = bulletCounter + 1
        If bulletCounter Mod 30 = 0 Then DoEvents ' Responsividade

        paraText = Trim(para.Range.text)

        ' Verifica se o paragrafo nao esta vazio
        If Len(paraText) > 0 Then
            ' Pega o primeiro caractere
            firstChar = Left(paraText, 1)

            ' Verifica se o primeiro caractere e um marcador comum
            Dim isBullet As Boolean
            isBullet = False

            For i = LBound(bulletMarkers) To UBound(bulletMarkers)
                If firstChar = bulletMarkers(i) Then
                    isBullet = True
                    Exit For
                End If
            Next i

            If isBullet Then
                ' Verifica se o paragrafo nao tem formatacao de lista ja aplicada
                If para.Range.ListFormat.ListType = 0 Then
                    ' Aplica o recuo a esquerda igual ao de uma lista com marcadores
                    With para.Format
                        .leftIndent = defaultIndent
                        .firstLineIndent = 0
                    End With
                    formattedCount = formattedCount + 1
                End If
            End If
        End If
    Next para

    If formattedCount > 0 Then
        LogMessage "Paragrafos iniciados com marcador formatados com recuo de lista: " & formattedCount, LOG_LEVEL_INFO
    End If

    FormatBulletedParagraphsIndent = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao formatar recuos de paragrafos com marcadores: " & Err.Description, LOG_LEVEL_WARNING
    FormatBulletedParagraphsIndent = False
End Function

'================================================================================
' REMOVER LINHAS EM BRANCO EXTRAS - Remove linhas duplicadas e aplica ajustes
'================================================================================
Private Sub RemoverLinhasEmBrancoExtras(doc As Document)
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim removedCount As Long
    Dim replacedCount As Long

    removedCount = 0
    replacedCount = 0

    LogMessage "Removendo linhas em branco extras e aplicando ajustes...", LOG_LEVEL_INFO

    ' --- Espacamento simples em todos os paragrafos ---
    Dim p As Paragraph
    For Each p In doc.Paragraphs
        On Error Resume Next
        With p.Format
            .LineSpacingRule = wdLineSpaceSingle
            .LineSpacing = 12
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
        On Error GoTo ErrorHandler
    Next p

    ' --- Remove linhas em branco extras ---
    For i = doc.Paragraphs.count To 2 Step -1
        Dim txtAtual As String, txtAnterior As String
        txtAtual = Trim(Replace(doc.Paragraphs(i).Range.text, vbCr, ""))
        txtAnterior = Trim(Replace(doc.Paragraphs(i - 1).Range.text, vbCr, ""))

        If txtAtual = "" And txtAnterior = "" Then
            On Error Resume Next
            doc.Paragraphs(i).Range.Delete
            If Err.Number = 0 Then removedCount = removedCount + 1
            Err.Clear
            On Error GoTo ErrorHandler
        End If
    Next i

    ' --- Substituicoes no texto ---
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = 1 ' wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False

        On Error Resume Next
        .text = "por intermedio do Setor,"
        .Replacement.text = "por intermedio do Setor competente,"
        If .Execute(Replace:=2) Then replacedCount = replacedCount + 1

        .text = "Indica ao Poder Executivo Municipal efetue"
        .Replacement.text = "Indica ao Poder Executivo Municipal que efetue"
        If .Execute(Replace:=2) Then replacedCount = replacedCount + 1

        .text = "Fomos procurados por municipes, solicitando essa providencia, pois segundo eles,"
        .Replacement.text = "Fomos procurados por municipes solicitando essa providencia, pois, segundo eles,"
        If .Execute(Replace:=2) Then replacedCount = replacedCount + 1
        On Error GoTo ErrorHandler
    End With

    ' --- Ajustes por paragrafo ---
    Dim para As Paragraph
    Dim adjustCounter As Long
    adjustCounter = 0
    For Each para In doc.Paragraphs
        adjustCounter = adjustCounter + 1
        If adjustCounter Mod 30 = 0 Then DoEvents ' Responsividade

        Dim cleanTxt As String
        cleanTxt = LCase(Trim(Replace(para.Range.text, vbCr, "")))
        cleanTxt = Replace(cleanTxt, "-", "")

        On Error Resume Next

        ' Espacamento extra antes e depois da data
        If InStr(cleanTxt, "plenario") > 0 And InStr(cleanTxt, "tancredo neves") > 0 Then
            para.Format.SpaceBefore = 24
            para.Format.SpaceAfter = 24
        End If

        ' Centraliza nome, cargo e partido
        If Left(cleanTxt, 8) = "vereador" _
           Or Left(cleanTxt, 9) = "vereadora" _
           Or InStr(cleanTxt, "vicepresidente") > 0 Then

            ' Cargo
            With para.Format
                .leftIndent = 0
                .RightIndent = 0
                .firstLineIndent = 0
                .alignment = wdAlignParagraphCenter
            End With

            ' Nome (paragrafo anterior)
            If Not para.Previous Is Nothing Then
                With para.Previous.Format
                    .leftIndent = 0
                    .RightIndent = 0
                    .firstLineIndent = 0
                    .alignment = wdAlignParagraphCenter
                End With
                para.Previous.Range.Font.Bold = True
            End If

            ' Partido (paragrafo seguinte)
            If Not para.Next Is Nothing Then
                With para.Next.Format
                    .leftIndent = 0
                    .RightIndent = 0
                    .firstLineIndent = 0
                    .alignment = wdAlignParagraphCenter
                End With
            End If
        End If

        On Error GoTo ErrorHandler
    Next para

    LogMessage "Linhas em branco removidas: " & removedCount & ", substituicoes: " & replacedCount, LOG_LEVEL_INFO
    Exit Sub

ErrorHandler:
    LogMessage "Erro em RemoverLinhasEmBrancoExtras: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' CENTER IMAGE AFTER PLENARIO - Centraliza imagem entre 5a e 7a linha apos Plenario
'================================================================================
Private Function CenterImageAfterPlenario(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim i As Long
    Dim plenarioIndex As Long
    Dim paraText As String
    Dim paraTextLower As String
    Dim lineCount As Long
    Dim centeredCount As Long

    plenarioIndex = 0
    centeredCount = 0

    ' Localiza o parágrafo "Plenário Dr. Tancredo Neves"
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(para.Range.text)
        paraTextLower = LCase(paraText)

        ' Procura por "Plenário" e "Tancredo Neves" com $DATAATUALEXTENSO$
        If InStr(paraTextLower, "plenário") > 0 And _
           InStr(paraTextLower, "tancredo neves") > 0 And _
           InStr(paraText, "$DATAATUALEXTENSO$") > 0 Then
            plenarioIndex = i
            Exit For
        End If
    Next i

    ' Se não encontrou o parágrafo do Plenário, retorna
    If plenarioIndex = 0 Then
        LogMessage "Parágrafo do Plenário não encontrado para centralizar imagem", LOG_LEVEL_INFO
        CenterImageAfterPlenario = True
        Exit Function
    End If

    ' Verifica as linhas 5, 6 e 7 após o Plenário (contando em branco e textuais)
    lineCount = 0
    For i = plenarioIndex + 1 To doc.Paragraphs.count
        lineCount = lineCount + 1

        ' Verifica apenas entre a 5ª e 7ª linha
        If lineCount >= 5 And lineCount <= 7 Then
            Set para = doc.Paragraphs(i)

            ' Se o parágrafo contém imagem, centraliza
            If para.Range.InlineShapes.count > 0 Then
                para.alignment = wdAlignParagraphCenter
                centeredCount = centeredCount + 1
                LogMessage "Imagem centralizada na linha " & lineCount & " após Plenário", LOG_LEVEL_INFO
            End If
        End If

        ' Para após a 7ª linha
        If lineCount > 7 Then
            Exit For
        End If
    Next i

    If centeredCount > 0 Then
        LogMessage "Imagens centralizadas após Plenário: " & centeredCount, LOG_LEVEL_INFO
    End If

    CenterImageAfterPlenario = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao centralizar imagem após Plenário: " & Err.Description, LOG_LEVEL_WARNING
    CenterImageAfterPlenario = False
End Function









'================================================================================
' GET CLIPBOARD DATA - Obtém dados da área de transferência
'================================================================================
Private Function GetClipboardData() As Variant
    On Error GoTo ErrorHandler

    ' Placeholder para dados da área de transferência
    ' Em uma implementação completa, seria necessário usar APIs do Windows
    ' ou métodos mais avançados para capturar dados binários
    GetClipboardData = "ImageDataPlaceholder"
    Exit Function

ErrorHandler:
    GetClipboardData = Empty
End Function

'================================================================================
' ENHANCED IMAGE PROTECTION - Proteção aprimorada durante formatação
'================================================================================
Private Function ProtectImagesInRange(targetRange As Range) As Boolean
    On Error GoTo ErrorHandler

    ' Verifica se há imagens no range antes de aplicar formatação
    If targetRange.InlineShapes.count > 0 Then
        ' OTIMIZADO: Aplica formatação caractere por caractere, protegendo imagens
        Dim i As Long
        Dim charRange As Range
        Dim charCount As Long
        charCount = SafeGetCharacterCount(targetRange) ' Cache da contagem segura

        If charCount > 0 Then ' Verificação de segurança
            For i = 1 To charCount
                Set charRange = targetRange.Characters(i)
                ' Só formata caracteres que não são parte de imagens
                If charRange.InlineShapes.count = 0 Then
                    With charRange.Font
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                        .Color = wdColorAutomatic
                    End With
                End If
            Next i
        End If
    Else
        ' Range sem imagens - formatação normal completa
        With targetRange.Font
            .Name = STANDARD_FONT
            .size = STANDARD_FONT_SIZE
            .Color = wdColorAutomatic
        End With
    End If

    ProtectImagesInRange = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na proteção de imagens: " & Err.Description, LOG_LEVEL_WARNING
    ProtectImagesInRange = False
End Function

'================================================================================
' LIMPEZA DE PROTEÇÃO DE IMAGENS
'================================================================================
Private Sub CleanupImageProtection()
    On Error Resume Next

    ' Limpa arrays de imagens
    If imageCount > 0 Then
        Dim i As Long
        For i = 0 To imageCount - 1
            Set savedImages(i).AnchorRange = Nothing
        Next i
    End If

    imageCount = 0
    ReDim savedImages(0)

    LogMessage "Variáveis de proteção de imagens limpas"
End Sub

'================================================================================
' VIEW SETTINGS PROTECTION SYSTEM - SISTEMA DE PROTEÇÃO DAS CONFIGURAÇÕES DE VISUALIZAÇÃO
'================================================================================

'================================================================================
' BACKUP VIEW SETTINGS - Faz backup das configurações de visualização originais
'================================================================================
Private Function BackupViewSettings(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Salvando visualização..."

    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow

    ' Backup das configurações de visualização
    With originalViewSettings
        .ViewType = docWindow.View.Type
        ' Réguas são controladas pelo Window, não pelo View
        On Error Resume Next
        .ShowHorizontalRuler = docWindow.DisplayRulers
        .ShowVerticalRuler = docWindow.DisplayVerticalRuler
        On Error GoTo ErrorHandler
        .ShowFieldCodes = docWindow.View.ShowFieldCodes
        .ShowBookmarks = docWindow.View.ShowBookmarks
        .ShowParagraphMarks = docWindow.View.ShowParagraphs
        .ShowSpaces = docWindow.View.ShowSpaces
        .ShowTabs = docWindow.View.ShowTabs
        .ShowHiddenText = docWindow.View.ShowHiddenText
        .ShowAll = docWindow.View.ShowAll
        .ShowDrawings = docWindow.View.ShowDrawings
        .ShowObjectAnchors = docWindow.View.ShowObjectAnchors
        .ShowTextBoundaries = docWindow.View.ShowTextBoundaries
        .ShowHighlight = docWindow.View.ShowHighlight
        ' .ShowAnimation removida - pode não existir em todas as versões
        .DraftFont = docWindow.View.Draft
        .WrapToWindow = docWindow.View.WrapToWindow
        .ShowPicturePlaceHolders = docWindow.View.ShowPicturePlaceHolders
        .ShowFieldShading = docWindow.View.FieldShading
        .TableGridlines = docWindow.View.TableGridlines
        ' .EnlargeFontsLessThan removida - pode não existir em todas as versões
    End With

    LogMessage "Backup das configurações de visualização concluído"
    BackupViewSettings = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao fazer backup das configurações de visualização: " & Err.Description, LOG_LEVEL_WARNING
    BackupViewSettings = False
End Function

'================================================================================
' RESTORE VIEW SETTINGS - Restaura as configurações de visualização originais
'================================================================================
Private Function RestoreViewSettings(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Restaurando visualização..."

    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow

    ' Restaura todas as configurações originais, EXCETO o zoom
    With docWindow.View
        .Type = originalViewSettings.ViewType
        .ShowFieldCodes = originalViewSettings.ShowFieldCodes
        .ShowBookmarks = originalViewSettings.ShowBookmarks
        .ShowParagraphs = originalViewSettings.ShowParagraphMarks
        .ShowSpaces = originalViewSettings.ShowSpaces
        .ShowTabs = originalViewSettings.ShowTabs
        .ShowHiddenText = originalViewSettings.ShowHiddenText
        .ShowAll = originalViewSettings.ShowAll
        .ShowDrawings = originalViewSettings.ShowDrawings
        .ShowObjectAnchors = originalViewSettings.ShowObjectAnchors
        .ShowTextBoundaries = originalViewSettings.ShowTextBoundaries
        .ShowHighlight = originalViewSettings.ShowHighlight
        ' .ShowAnimation removida para compatibilidade
        .Draft = originalViewSettings.DraftFont
        .WrapToWindow = originalViewSettings.WrapToWindow
        .ShowPicturePlaceHolders = originalViewSettings.ShowPicturePlaceHolders
        .FieldShading = originalViewSettings.ShowFieldShading
        .TableGridlines = originalViewSettings.TableGridlines
        ' .EnlargeFontsLessThan removida para compatibilidade

        ' ZOOM é mantido em 110% - única configuração que permanece alterada
        .Zoom.Percentage = 110
    End With

    ' Configurações específicas do Window (para réguas)
    docWindow.DisplayRulers = originalViewSettings.ShowHorizontalRuler
    docWindow.DisplayVerticalRuler = originalViewSettings.ShowVerticalRuler

    LogMessage "Configurações de visualização originais restauradas (zoom mantido em 110%)"
    RestoreViewSettings = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao restaurar configurações de visualização: " & Err.Description, LOG_LEVEL_WARNING
    RestoreViewSettings = False
End Function

'================================================================================
' CLEANUP VIEW SETTINGS - Limpeza das variáveis de configurações de visualização
'================================================================================
Private Sub CleanupViewSettings()
    On Error Resume Next

    ' Reinicializa a estrutura de configurações
    With originalViewSettings
        .ViewType = 0
        .ShowVerticalRuler = False
        .ShowHorizontalRuler = False
        .ShowFieldCodes = False
        .ShowBookmarks = False
        .ShowParagraphMarks = False
        .ShowSpaces = False
        .ShowTabs = False
        .ShowHiddenText = False
        .ShowOptionalHyphens = False
        .ShowAll = False
        .ShowDrawings = False
        .ShowObjectAnchors = False
        .ShowTextBoundaries = False
        .ShowHighlight = False
        ' .ShowAnimation removida para compatibilidade
        .DraftFont = False
        .WrapToWindow = False
        .ShowPicturePlaceHolders = False
        .ShowFieldShading = 0
        .TableGridlines = False
        ' .EnlargeFontsLessThan removida para compatibilidade
    End With

    LogMessage "Variáveis de configurações de visualização limpas"
End Sub

'================================================================================
' SUBSTITUIÇÃO DO PARÁGRAFO DE LOCAL E DATA
'================================================================================
Private Sub ReplacePlenarioDateParagraph(doc As Document)
    On Error GoTo ErrorHandler

    If doc Is Nothing Then Exit Sub

    Dim para As Paragraph
    Dim paraText As String
    Dim matchCount As Integer
    Dim terms() As String

    ' Define os termos de busca
    terms = Split("Palácio 15 de Junho,Plenário,Dr. Tancredo Neves," & _
                 " de janeiro de , de fevereiro de, de março de, de abril de," & _
                 " de maio de, de junho de, de julho de, de agosto de," & _
                 " de setembro de, de outubro de, de novembro de, de dezembro de", ",")

    ' Processa cada paragrafo
    Dim plenCounter As Long
    plenCounter = 0
    For Each para In doc.Paragraphs
        plenCounter = plenCounter + 1
        If plenCounter Mod 30 = 0 Then DoEvents ' Responsividade

        matchCount = 0

        ' Pula paragrafos muito longos
        If Len(para.Range.text) <= 80 Then
            paraText = para.Range.text

            ' Conta matches
            Dim term As Variant
            For Each term In terms
                If InStr(1, paraText, CStr(term), vbTextCompare) > 0 Then
                    matchCount = matchCount + 1
                End If
                If matchCount >= 2 Then
                    ' Encontrou 2+ matches, faz a substituição
                    ' Usa Delete + InsertAfter para preservar o marcador de parágrafo
                    para.Range.Select
                    Selection.MoveEnd unit:=wdCharacter, count:=-1 ' Exclui o marcador de parágrafo
                    Selection.Delete
                    Selection.InsertAfter "Plenário ""Dr. Tancredo Neves"", $DATAATUALEXTENSO$."
                    ' Aplica formatação: centralizado e sem recuos
                    With para.Range.ParagraphFormat
                        .leftIndent = 0
                        .firstLineIndent = 0
                        .alignment = wdAlignParagraphCenter
                    End With
                    LogMessage "Parágrafo de plenário substituído e formatado", LOG_LEVEL_INFO
                    Exit For
                End If
            Next term
        End If
    Next para

    Exit Sub

ErrorHandler:
    LogMessage "Erro ao processar parágrafos: " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' GERENCIAMENTO DE DIRETÓRIO DE BACKUP
'================================================================================
Private Function EnsureBackupDirectory(doc As Document) As String
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim backupPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Garante que a estrutura de pastas do projeto existe
    EnsureChainsawFolders

    ' SEMPRE USA %TEMP%\.chainsaw\props\backups para todos os documentos
    backupPath = GetChainsawBackupsPath()

    ' Cria o diretório se não existir
    If Not fso.FolderExists(backupPath) Then
        fso.CreateFolder backupPath
        LogMessage "Pasta de backup criada: " & backupPath, LOG_LEVEL_INFO
    End If

    EnsureBackupDirectory = backupPath
    Exit Function

ErrorHandler:
    LogMessage "Erro ao criar pasta de backup: " & Err.Description, LOG_LEVEL_ERROR
    ' Retorna pasta do documento ou TEMP como fallback
    If doc.Path <> "" Then
        EnsureBackupDirectory = doc.Path
    Else
        EnsureBackupDirectory = Environ("TEMP")
    End If
End Function

'================================================================================
' VERIFICAÇÃO DE VERSÃO E ATUALIZAÇÃO
'================================================================================

' Função: CheckForUpdates
' Descrição: Verifica se há uma nova versão disponível no GitHub
' Retorna: True se houver atualização disponível, False caso contrário
'================================================================================
Public Function CheckForUpdates() As Boolean
    On Error GoTo ErrorHandler

    Dim localVersion As String
    Dim remoteVersion As String
    Dim updateAvailable As Boolean

    CheckForUpdates = False

    ' Obtém versão local
    localVersion = GetLocalVersion()
    If localVersion = "" Then
        LogMessage "Não foi possível obter versão local", LOG_LEVEL_WARNING
        Exit Function
    End If

    ' Obtém versão remota do GitHub
    remoteVersion = GetRemoteVersion()
    If remoteVersion = "" Then
        LogMessage "Não foi possível obter versão remota", LOG_LEVEL_WARNING
        Exit Function
    End If

    ' Compara versões
    updateAvailable = CompareVersions(remoteVersion, localVersion) > 0

    If updateAvailable Then
        LogMessage "Atualização disponível: " & localVersion & " -> " & remoteVersion, LOG_LEVEL_INFO
        CheckForUpdates = True
    Else
        LogMessage "Sistema está atualizado (v" & localVersion & ")", LOG_LEVEL_INFO
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro ao verificar atualizações: " & Err.Description, LOG_LEVEL_ERROR
    CheckForUpdates = False
End Function

' Função: GetLocalVersion
' Descricao: Le a versao instalada do arquivo VERSION local
' Retorna: String com a versão local ou "" em caso de erro
'================================================================================
Private Function GetLocalVersion() As String
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim versionFile As String
    Dim fileContent As String
    Dim version As String

    GetLocalVersion = ""

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Caminho do arquivo de versao local
    versionFile = GetProjectRootPath() & "\VERSION"

    If Not fso.FileExists(versionFile) Then
        LogMessage "Arquivo de versão local não encontrado: " & versionFile, LOG_LEVEL_WARNING
        Exit Function
    End If

    ' Lê o arquivo
    fileContent = ReadTextFile(versionFile)

    ' Extrai versao (X.Y.Z)
    version = ExtractVersionFromText(fileContent)

    GetLocalVersion = version

    Exit Function

ErrorHandler:
    LogMessage "Erro ao obter versão local: " & Err.Description, LOG_LEVEL_ERROR
    GetLocalVersion = ""
End Function

' Função: GetRemoteVersion
' Descrição: Baixa e lê a versão disponível no GitHub
' Retorna: String com a versão remota ou "" em caso de erro
'================================================================================
Private Function GetRemoteVersion() As String
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim url As String
    Dim response As String
    Dim version As String

    GetRemoteVersion = ""

    ' URL do arquivo VERSION no GitHub
    url = "https://raw.githubusercontent.com/chrmsantos/chainsaw/main/VERSION"

    ' Cria objeto HTTP
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Faz requisição GET
    http.Open "GET", url, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.send

    ' Verifica resposta
    If http.Status = 200 Then
        response = http.responseText
        version = ExtractVersionFromText(response)
        GetRemoteVersion = version
    Else
        LogMessage "Erro HTTP ao buscar versão remota: " & http.Status, LOG_LEVEL_WARNING
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro ao obter versão remota: " & Err.Description, LOG_LEVEL_ERROR
    GetRemoteVersion = ""
End Function

' Função: ExtractJsonValue
' Descrição: Extrai um valor de um JSON simples usando regex
' Parâmetros:
'   - jsonText: String contendo o JSON
'   - key: Chave a ser extraída
' Retorna: Valor da chave ou "" se não encontrado
'================================================================================
Private Function ExtractVersionFromText(ByVal textValue As String) As String
    On Error GoTo ErrorHandler

    Dim regex As Object
    Dim matches As Object
    Dim pattern As String

    ExtractVersionFromText = ""

    Set regex = CreateObject("VBScript.RegExp")

    ' Pattern para extrair versao no formato X.Y.Z
    pattern = "([0-9]+)\.([0-9]+)\.([0-9]+)"

    regex.pattern = pattern
    regex.IgnoreCase = True
    regex.Global = False

    Set matches = regex.Execute(textValue)

    If matches.count > 0 Then
        ExtractVersionFromText = matches(0).Value
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro ao extrair versao: " & Err.Description, LOG_LEVEL_ERROR
    ExtractVersionFromText = ""
End Function

' Função: CompareVersions
' Descrição: Compara duas versões no formato X.Y.Z
' Parâmetros:
'   - version1: Primeira versão
'   - version2: Segunda versão
' Retorna: 1 se version1 > version2, -1 se version1 < version2, 0 se iguais
'================================================================================
Private Function CompareVersions(ByVal version1 As String, ByVal version2 As String) As Integer
    On Error GoTo ErrorHandler

    Dim v1Parts() As String
    Dim v2Parts() As String
    Dim i As Integer
    Dim v1Num As Long, v2Num As Long

    CompareVersions = 0

    ' Remove espaços
    version1 = Trim(version1)
    version2 = Trim(version2)

    ' Divide versões em partes
    v1Parts = Split(version1, ".")
    v2Parts = Split(version2, ".")

    ' Compara cada parte
    For i = 0 To 2
        v1Num = 0
        v2Num = 0

        If i <= UBound(v1Parts) Then v1Num = CLng(v1Parts(i))
        If i <= UBound(v2Parts) Then v2Num = CLng(v2Parts(i))

        If v1Num > v2Num Then
            CompareVersions = 1
            Exit Function
        ElseIf v1Num < v2Num Then
            CompareVersions = -1
            Exit Function
        End If
    Next i

    Exit Function

ErrorHandler:
    LogMessage "Erro ao comparar versões: " & Err.Description, LOG_LEVEL_ERROR
    CompareVersions = 0
End Function

' Função: ReadTextFile
' Descrição: Lê o conteúdo completo de um arquivo de texto
' Parâmetros:
'   - filePath: Caminho completo do arquivo
' Retorna: Conteúdo do arquivo como String
'================================================================================
Private Function ReadTextFile(ByVal filePath As String) As String
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim file As Object
    Dim content As String

    ReadTextFile = ""

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(filePath) Then
        Set file = fso.OpenTextFile(filePath, 1, False, -1) ' -1 = Unicode
        content = file.ReadAll
        file.Close
        ReadTextFile = content
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro ao ler arquivo: " & Err.Description, LOG_LEVEL_ERROR
    ReadTextFile = ""
End Function

' Sub: PromptForUpdate
' Descrição: Verifica se há atualização e pergunta ao usuário se deseja atualizar
'================================================================================
Public Sub PromptForUpdate()
    On Error GoTo ErrorHandler

    Dim updateAvailable As Boolean
    Dim response As VbMsgBoxResult
    Dim installerPath As String
    Dim shellCmd As String

    ' Verifica se há atualizações
    updateAvailable = CheckForUpdates()

    If Not updateAvailable Then
        MsgBox "Seu sistema CHAINSAW está atualizado!", vbInformation, "CHAINSAW - Verificação de Versão"
        Exit Sub
    End If

    ' Pergunta ao usuário se deseja atualizar
    response = MsgBox("Uma nova versão do CHAINSAW está disponível!" & vbCrLf & vbCrLf & _
                      "Deseja atualizar agora?" & vbCrLf & vbCrLf & _
                      "O instalador será executado e o Word será fechado.", _
                      vbYesNo + vbQuestion, "CHAINSAW - Atualização Disponível")

    If response = vbYes Then
        ' Caminho do instalador
        installerPath = Environ("USERPROFILE") & "\chainsaw\chainsaw_installer.cmd"

        ' Verifica se o instalador existe
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")

        If fso.FileExists(installerPath) Then
            ' Executa o instalador
            shellCmd = "cmd.exe /c """ & installerPath & """"

            ' Salva todos os documentos abertos
            Dim doc As Object
            For Each doc In Application.Documents
                If doc.Saved = False Then
                    doc.Save
                End If
            Next doc

            ' Executa instalador e fecha o Word
            CreateObject("WScript.Shell").Run shellCmd, 1, False

            MsgBox "O instalador será executado. O Word será fechado agora.", vbInformation, "CHAINSAW - Atualização"
            Application.Quit SaveChanges:=wdSaveChanges
        Else
            MsgBox "Instalador não encontrado em:" & vbCrLf & installerPath & vbCrLf & vbCrLf & _
                   "Baixe manualmente de: https://github.com/chrmsantos/chainsaw", _
                   vbExclamation, "CHAINSAW - Erro"
        End If
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Erro ao processar atualização: " & Err.Description, vbCritical, "CHAINSAW - Erro"
End Sub

'================================================================================
' Sub: ExecutarInstalador
' Descrição: Executa o chainsaw_installer.cmd a partir da interface do Word
' Uso: Pode ser chamado de um botão na ribbon ou atalho de teclado
'================================================================================
Public Sub ExecutarInstalador()
    On Error GoTo ErrorHandler

    Dim installerPath As String
    Dim shellCmd As String
    Dim fso As Object
    Dim response As VbMsgBoxResult

    ' Pergunta confirmação ao usuário
    response = MsgBox("Deseja executar o instalador do CHAINSAW?" & vbCrLf & vbCrLf & _
                      "Isso irá:" & vbCrLf & _
                      "• Baixar a versão mais recente do GitHub" & vbCrLf & _
                      "• Instalar/atualizar o sistema" & vbCrLf & _
                      "• Fechar o Word ao final da instalação" & vbCrLf & vbCrLf & _
                      "Continuar?", _
                      vbYesNo + vbQuestion, "CHAINSAW - Executar Instalador")

    If response <> vbYes Then
        Exit Sub
    End If

    ' Caminho do instalador
    installerPath = Environ("USERPROFILE") & "\chainsaw\chainsaw_installer.cmd"

    ' Verifica se o instalador existe
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(installerPath) Then
        MsgBox "Instalador não encontrado em:" & vbCrLf & installerPath & vbCrLf & vbCrLf & _
               "Baixe manualmente de: https://github.com/chrmsantos/chainsaw/raw/main/chainsaw_installer.cmd", _
               vbExclamation, "CHAINSAW - Instalador Não Encontrado"
        Exit Sub
    End If

    ' Salva todos os documentos abertos antes de executar o instalador
    Dim doc As Object
    For Each doc In Application.Documents
        If doc.Saved = False Then
            On Error Resume Next
            doc.Save
            On Error GoTo ErrorHandler
        End If
    Next doc

    ' Executa o instalador em uma nova janela de comando
    shellCmd = "cmd.exe /c """ & installerPath & """"
    CreateObject("WScript.Shell").Run shellCmd, 1, False

    ' Mensagem informativa
    MsgBox "O instalador foi iniciado em uma nova janela." & vbCrLf & vbCrLf & _
           "O Word será fechado ao final da instalação.", _
           vbInformation, "CHAINSAW - Instalador Iniciado"

    ' Fecha o Word após 2 segundos (tempo para o instalador iniciar)
    Application.OnTime Now + TimeValue("00:00:02"), "FecharWord"

    Exit Sub

ErrorHandler:
    MsgBox "Erro ao executar instalador: " & Err.Description, vbCritical, "CHAINSAW - Erro"
    LogMessage "Erro ao executar instalador: " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' Sub: FecharWord
' Descrição: Fecha o Word (usado após executar o instalador)
'================================================================================
Private Sub FecharWord()
    On Error Resume Next
    Application.Quit SaveChanges:=wdSaveChanges
End Sub

'================================================================================
' APLICAÇÃO DE FORMATAÇÃO FINAL UNIVERSAL
'================================================================================
Private Sub ApplyUniversalFinalFormatting(doc As Document)
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraCount As Long
    Dim formattedCount As Long

    paraCount = doc.Paragraphs.count
    formattedCount = 0

    LogMessage "Aplicando formatacao final universal: Arial 12, espacamento 1.0, 1 linha entre paragrafos...", LOG_LEVEL_INFO

    ' Processa todos os paragrafos
    Dim universalCounter As Long
    universalCounter = 0
    For Each para In doc.Paragraphs
        universalCounter = universalCounter + 1
        If universalCounter Mod 20 = 0 Then DoEvents ' Responsividade

        On Error Resume Next

        ' Aplica fonte Arial 12
        With para.Range.Font
            .Name = "Arial"
            .size = 12
        End With

        ' Aplica espacamento de linha 1.0 (simples)
        With para.Format
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 0   ' Sem espaco antes do paragrafo
            .SpaceAfter = 0    ' Sem espaco depois do paragrafo
        End With

        If Err.Number = 0 Then
            formattedCount = formattedCount + 1
        Else
            Err.Clear
        End If

        On Error GoTo ErrorHandler
    Next para

    LogMessage "Formatacao final aplicada: " & formattedCount & " de " & paraCount & " paragrafos", LOG_LEVEL_INFO
    Exit Sub

ErrorHandler:
    LogMessage "Erro ao aplicar formatacao final universal: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' ADIÇÃO DE ESPAÇAMENTO ESPECIAL (EMENTA, JUSTIFICATIVA, DATA)
'================================================================================
Private Sub AddSpecialElementsSpacing(doc As Document)
    On Error GoTo ErrorHandler

    Dim elementsProcessed As Long
    elementsProcessed = 0

    LogMessage "Adicionando espacamento especial para ementa, justificativa e data...", LOG_LEVEL_INFO

    ' Garante sem espaco antes e depois da Ementa
    If ementaParaIndex > 0 And ementaParaIndex <= doc.Paragraphs.count Then
        On Error Resume Next
        With doc.Paragraphs(ementaParaIndex).Format
            .SpaceBefore = 0   ' Sem espaco antes
            .SpaceAfter = 0    ' Sem espaco depois
        End With
        If Err.Number = 0 Then elementsProcessed = elementsProcessed + 1
        Err.Clear
        On Error GoTo ErrorHandler
    End If

    ' Garante sem espaco antes e depois do Título Justificativa
    If tituloJustificativaIndex > 0 And tituloJustificativaIndex <= doc.Paragraphs.count Then
        On Error Resume Next
        With doc.Paragraphs(tituloJustificativaIndex).Format
            .SpaceBefore = 0   ' Sem espaco antes
            .SpaceAfter = 0    ' Sem espaco depois
        End With
        If Err.Number = 0 Then elementsProcessed = elementsProcessed + 1
        Err.Clear
        On Error GoTo ErrorHandler
    End If

    ' Garante sem espaco antes e depois da Data
    If dataParaIndex > 0 And dataParaIndex <= doc.Paragraphs.count Then
        On Error Resume Next
        With doc.Paragraphs(dataParaIndex).Format
            .SpaceBefore = 0   ' Sem espaco antes
            .SpaceAfter = 0    ' Sem espaco depois
        End With
        If Err.Number = 0 Then elementsProcessed = elementsProcessed + 1
        Err.Clear
        On Error GoTo ErrorHandler
    End If

    LogMessage "Espacamento especial aplicado a " & elementsProcessed & " elementos", LOG_LEVEL_INFO
    Exit Sub

ErrorHandler:
    LogMessage "Erro ao adicionar espacamento especial: " & Err.Description, LOG_LEVEL_WARNING
End Sub
