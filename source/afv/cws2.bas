' =============================================================================
' CHAINSAW - Sistema de Padronização de Proposituras Legislativas
' =============================================================================
' Versão: 1.1-RC1-202511071045
' Licença: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
' Compatibilidade: Microsoft Word 2010+
' Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
' =============================================================================
' CHANGELOG v1.1:
' - Adicionado sistema de identificação de elementos estruturais da propositura
' - Novos identificadores: Título, Ementa, Proposição, Justificativa, Data,
'   Assinatura, Anexo e Propositura completa
' - Funções públicas de acesso aos elementos: GetTituloRange, GetEmentaRange,
'   GetProposicaoRange, GetJustificativaRange, GetDataRange, GetAssinaturaRange,
'   GetTituloAnexoRange, GetAnexoRange, GetProposituraRange
' - Integração com sistema de cache de parágrafos existente
' - Identificação automática durante BuildParagraphCache
' - Função GetElementInfo para relatório completo dos elementos
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

Private Const TOP_MARGIN_CM As Double = 4
Private Const BOTTOM_MARGIN_CM As Double = 1
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
Private Const MIN_SUPPORTED_VERSION As Long = 14
Private Const REQUIRED_STRING As String = "$NUMERO$/$ANO$"
Private Const BACKUP_FOLDER_NAME As String = "backups"
Private Const CHAINSAW_ROOT_FOLDER As String = ".chainsaw"
Private Const CHAINSAW_BACKUPS_FOLDER As String = "backups"
Private Const CHAINSAW_LOGS_FOLDER As String = "logs"
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
' CONSTANTES DE IDENTIFICAÇÃO DE ELEMENTOS ESTRUTURAIS
'================================================================================
' Critérios para identificação dos elementos da propositura
Private Const TITULO_MIN_LENGTH As Long = 15              ' Comprimento mínimo do título
Private Const EMENTA_MIN_LEFT_INDENT As Single = 6        ' Recuo mínimo à esquerda da ementa (em pontos)
Private Const PLENARIO_TEXT As String = "plenário ""dr. tancredo neves"""  ' Texto identificador da data
Private Const ANEXO_TEXT_SINGULAR As String = "anexo"     ' Texto identificador de anexo (singular)
Private Const ANEXO_TEXT_PLURAL As String = "anexos"      ' Texto identificador de anexo (plural)
Private Const ASSINATURA_PARAGRAPH_COUNT As Long = 3      ' Número de parágrafos da assinatura
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
' VARIÁVEIS DE IDENTIFICAÇÃO DE ELEMENTOS ESTRUTURAIS
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
        LogMessage "Nenhum documento acessível para processamento", LOG_LEVEL_ERROR
        Exit Sub
    End If
    On Error GoTo CriticalErrorHandler
    
    ' Valida integridade do documento
    If Not IsDocumentHealthy(doc) Then
        Application.StatusBar = "Erro: Documento inacessível"
        MsgBox "Documento corrompido ou inacessível." & vbCrLf & _
               "Salve uma cópia e reabra.", vbCritical, "Erro de Documento"
        Exit Sub
    End If
    
    If Not InitializeLogging(doc) Then
        LogMessage "Falha ao inicializar sistema de logs", LOG_LEVEL_WARNING
    End If
    
    LogMessage "Iniciando padronização do documento: " & doc.Name, LOG_LEVEL_INFO
    
    ' Inicializa barra de progresso (15 etapas principais)
    InitializeProgress 15
    
    ' ---------------------------------------------------------------------------
    ' INÍCIO DO GRUPO DE DESFAZER - TODAS as operações são agrupadas aqui
    ' ---------------------------------------------------------------------------
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "Padronização de Documento"
    If Err.Number = 0 Then
        undoGroupEnabled = True
        LogMessage "UndoRecord iniciado com sucesso", LOG_LEVEL_INFO
    Else
        LogMessage "Aviso: Não foi possível iniciar UndoRecord: " & Err.Description, LOG_LEVEL_WARNING
        undoGroupEnabled = False
    End If
    Err.Clear
    On Error GoTo CriticalErrorHandler
    ' ---------------------------------------------------------------------------
    
    If Not SetAppState(False, "Iniciando...") Then
        LogMessage "Falha ao configurar estado da aplicação", LOG_LEVEL_WARNING
    End If
    
    IncrementProgress "Verificando documento"
    If Not PreviousChecking(doc) Then
        GoTo CleanUp
    End If
    
    If doc.Path = "" Then
        If Not SaveDocumentFirst(doc) Then
            Application.StatusBar = "Cancelado: documento não salvo"
            LogMessage "Operação cancelada - documento não foi salvo", LOG_LEVEL_INFO
            GoTo CleanUp ' Garante fechamento do UndoRecord
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
    
    ' Backup de formatações de lista antes das formatações
    IncrementProgress "Protegendo listas"
    If Not BackupListFormats(doc) Then
        LogMessage "Aviso: Falha no backup de listas - formatações de lista podem ser perdidas", LOG_LEVEL_WARNING
    End If
    
    ' Constrói cache de parágrafos para otimização
    IncrementProgress "Indexando parágrafos"
    BuildParagraphCache doc

    IncrementProgress "Formatando documento"
    If Not PreviousFormatting(doc) Then
        GoTo CleanUp
    End If
    
RemoverLinhasEmBrancoExtras doc

    ' Restaura imagens após formatações
    IncrementProgress "Restaurando imagens"
    ' Centraliza "Vereador/Vereadora"
Dim i As Long
For i = 1 To doc.Paragraphs.count
    Dim p As Paragraph
    Set p = doc.Paragraphs(i)
    Dim cleanTxt As String
    cleanTxt = Trim(LCase(Replace(p.Range.text, vbCr, "")))
    If cleanTxt = "vereador" Or cleanTxt = "vereadora" Then
        With p.Format
            .leftIndent = 0
            .firstLineIndent = 0
            .alignment = wdAlignParagraphCenter
        End With
    End If
Next i

    If Not RestoreAllImages(doc) Then
        LogMessage "Aviso: Algumas imagens podem ter sido afetadas durante o processamento", LOG_LEVEL_WARNING
    End If
    
    ' Restaura formatações de lista após formatações
    IncrementProgress "Restaurando listas"
    If Not RestoreListFormats(doc) Then
        LogMessage "Aviso: Algumas formatações de lista podem não ter sido restauradas", LOG_LEVEL_WARNING
    End If
    
    ' Formata parágrafos iniciados com número (aplica recuo de lista numerada)
    IncrementProgress "Ajustando numeração"
    If Not FormatNumberedParagraphsIndent(doc) Then
        LogMessage "Aviso: Falha ao formatar recuos de parágrafos numerados", LOG_LEVEL_WARNING
    End If
    
    ' Formata parágrafos iniciados com marcador (aplica recuo de lista com marcadores)
    IncrementProgress "Ajustando marcadores"
    If Not FormatBulletedParagraphsIndent(doc) Then
        LogMessage "Aviso: Falha ao formatar recuos de parágrafos com marcadores", LOG_LEVEL_WARNING
    End If
    
    ' Formata recuos de parágrafos com imagens (zera recuo à esquerda)
    IncrementProgress "Ajustando layout"
    If Not FormatImageParagraphsIndents(doc) Then
        LogMessage "Aviso: Falha ao formatar recuos de imagens", LOG_LEVEL_WARNING
    End If
    
    ' Centraliza imagem entre 5ª e 7ª linha após Plenário
    IncrementProgress "Centralizando elementos"
    If Not CenterImageAfterPlenario(doc) Then
        LogMessage "Aviso: Falha ao centralizar imagem após Plenário", LOG_LEVEL_WARNING
    End If

    ' Restaura configurações de visualização originais (exceto zoom)
    IncrementProgress "Restaurando visualização"
    If Not RestoreViewSettings(doc) Then
        LogMessage "Aviso: Algumas configurações de visualização podem não ter sido restauradas", LOG_LEVEL_WARNING
    End If

    If formattingCancelled Then
        GoTo CleanUp
    End If

    IncrementProgress "Finalizando"
    LogMessage "Documento padronizado com sucesso", LOG_LEVEL_INFO
    
    ' Mostra 100% por 1 segundo antes de limpar
    UpdateProgress "Concluído!", 100
    
    ' Pausa de 1 segundo (Word VBA não tem Application.Wait)
    Dim pauseTime As Double
    pauseTime = Timer
    Do While Timer < pauseTime + 1
        DoEvents
    Loop

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
    
    ClearParagraphCache ' Limpa cache de parágrafos
    SafeCleanup
    CleanupImageProtection ' Nova função para limpar variáveis de proteção de imagens
    CleanupViewSettings    ' Nova função para limpar variáveis de configurações de visualização
    
    If Not SetAppState(True, "Concluído!") Then
        LogMessage "Falha ao restaurar estado da aplicação", LOG_LEVEL_WARNING
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
               "   " & IIf(backupFilePath <> "", backupFilePath, doc.Path & "\" & BACKUP_FOLDER_NAME & "\") & vbCrLf & vbCrLf & _
               "[LOG] Log salvo em:" & vbCrLf & _
               "   " & logFilePath & statusMsg, _
               vbInformation, "CHAINSAW - Padronização Concluída"
    End If
    
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
        If Not EOF(fileNumber) Then
            Close fileNumber
        End If
    Next fileNumber
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
    Do While Len(txt) > 0 And InStr(".,;:", Right(txt, 1)) > 0 And safetyCounter < 100
        txt = Left(txt, Len(txt) - 1)
        safetyCounter = safetyCounter + 1
    Loop
    
    GetCleanParagraphText = Trim(LCase(txt))
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
    result = text
    
    ' Remove caracteres de controle em uma única passagem
    result = Replace(result, vbCr, "")
    result = Replace(result, vbLf, "")
    result = Replace(result, vbTab, " ")
    
    ' Remove espaços múltiplos
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
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
' CONSTRUÇÃO DO CACHE DE PARÁGRAFOS - Otimização principal
'================================================================================
Private Sub BuildParagraphCache(doc As Document)
    On Error GoTo ErrorHandler
    
    Dim startTime As Double
    startTime = Timer
    
    LogMessage "Iniciando construção do cache de parágrafos...", LOG_LEVEL_INFO
    
    cacheSize = doc.Paragraphs.count
    ReDim paragraphCache(1 To cacheSize)
    
    Dim i As Long
    Dim para As Paragraph
    Dim rawText As String
    
    For i = 1 To cacheSize
        Set para = doc.Paragraphs(i)
        
        ' Captura o texto bruto uma única vez
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
        
        ' Atualiza progresso a cada 100 parágrafos
        If i Mod 100 = 0 Then
            UpdateProgress "Indexando: " & i & "/" & cacheSize, 5 + (i * 5 \ cacheSize)
        End If
    Next i
    
    cacheEnabled = True
    
    Dim elapsed As Single
    elapsed = Timer - startTime
    
    LogMessage "Cache construído: " & cacheSize & " parágrafos em " & Format(elapsed, "0.00") & "s", LOG_LEVEL_INFO
    
    ' Identifica a estrutura do documento após construir o cache
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
'--------------------------------------------------------------------------------
Public Function GetElementInfo(doc As Document) As String
    On Error Resume Next
    
    Dim info As String
    info = "=== INFORMAÇÕES DOS ELEMENTOS ESTRUTURAIS ===" & vbCrLf
    
    If tituloParaIndex > 0 Then
        info = info & "Título: Parágrafo " & tituloParaIndex & vbCrLf
    Else
        info = info & "Título: Não identificado" & vbCrLf
    End If
    
    If ementaParaIndex > 0 Then
        info = info & "Ementa: Parágrafo " & ementaParaIndex & vbCrLf
    Else
        info = info & "Ementa: Não identificado" & vbCrLf
    End If
    
    If proposicaoStartIndex > 0 And proposicaoEndIndex > 0 Then
        info = info & "Proposição: Parágrafos " & proposicaoStartIndex & " a " & proposicaoEndIndex & _
                      " (" & (proposicaoEndIndex - proposicaoStartIndex + 1) & " parágrafos)" & vbCrLf
    Else
        info = info & "Proposição: Não identificado" & vbCrLf
    End If
    
    If tituloJustificativaIndex > 0 Then
        info = info & "Título Justificativa: Parágrafo " & tituloJustificativaIndex & vbCrLf
    Else
        info = info & "Título Justificativa: Não identificado" & vbCrLf
    End If
    
    If justificativaStartIndex > 0 And justificativaEndIndex > 0 Then
        info = info & "Justificativa: Parágrafos " & justificativaStartIndex & " a " & justificativaEndIndex & _
                      " (" & (justificativaEndIndex - justificativaStartIndex + 1) & " parágrafos)" & vbCrLf
    Else
        info = info & "Justificativa: Não identificado" & vbCrLf
    End If
    
    If dataParaIndex > 0 Then
        info = info & "Data (Plenário): Parágrafo " & dataParaIndex & vbCrLf
    Else
        info = info & "Data (Plenário): Não identificado" & vbCrLf
    End If
    
    If assinaturaStartIndex > 0 And assinaturaEndIndex > 0 Then
        info = info & "Assinatura: Parágrafos " & assinaturaStartIndex & " a " & assinaturaEndIndex & _
                      " (" & (assinaturaEndIndex - assinaturaStartIndex + 1) & " parágrafos)" & vbCrLf
    Else
        info = info & "Assinatura: Não identificado" & vbCrLf
    End If
    
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
' ATUALIZAÇÃO DA BARRA DE PROGRESSO
'================================================================================
Private Sub UpdateProgress(message As String, percentComplete As Long)
    Dim progressBar As String
    Dim barLength As Long
    Dim filledLength As Long
    
    ' Limita entre 0 e 100
    If percentComplete < 0 Then percentComplete = 0
    If percentComplete > 100 Then percentComplete = 100
    
    ' Barra de 20 caracteres
    barLength = 20
    filledLength = CLng(barLength * percentComplete / 100)
    
    ' Constrói a barra visual
    progressBar = "["
    Dim i As Long
    For i = 1 To barLength
        If i <= filledLength Then
            progressBar = progressBar & "¦"
        Else
            progressBar = progressBar & "¦"
        End If
    Next i
    progressBar = progressBar & "] " & Format(percentComplete, "0") & "%"
    
    ' Atualiza StatusBar com mensagem e barra
    Application.StatusBar = message & " " & progressBar
    
    ' Força atualização da tela
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
    percent = CLng((currentStep * 100) / totalSteps)
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
' FUNÇÕES DE CAMINHO - Estrutura .chainsaw
'================================================================================

'--------------------------------------------------------------------------------
' GetChainsawTempPath - Retorna caminho base .chainsaw no Temp
'--------------------------------------------------------------------------------
Private Function GetChainsawTempPath() As String
    GetChainsawTempPath = Environ("USERPROFILE") & "\AppData\Local\Temp\" & CHAINSAW_ROOT_FOLDER
End Function

'--------------------------------------------------------------------------------
' GetChainsawBackupsPath - Retorna caminho para backups
'--------------------------------------------------------------------------------
Private Function GetChainsawBackupsPath() As String
    GetChainsawBackupsPath = GetChainsawTempPath() & "\" & CHAINSAW_BACKUPS_FOLDER
End Function

'--------------------------------------------------------------------------------
' GetChainsawLogsPath - Retorna caminho para logs
'--------------------------------------------------------------------------------
Private Function GetChainsawLogsPath() As String
    GetChainsawLogsPath = GetChainsawTempPath() & "\" & CHAINSAW_LOGS_FOLDER
End Function

'--------------------------------------------------------------------------------
' EnsureChainsawFolders - Cria estrutura de pastas .chainsaw se não existir
'--------------------------------------------------------------------------------
Private Sub EnsureChainsawFolders()
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Cria pasta raiz .chainsaw
    If Not fso.FolderExists(GetChainsawTempPath()) Then
        fso.CreateFolder GetChainsawTempPath()
    End If
    
    ' Cria pasta backups
    If Not fso.FolderExists(GetChainsawBackupsPath()) Then
        fso.CreateFolder GetChainsawBackupsPath()
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

Private Function InitializeLogging(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim logFolder As String
    Dim docNameClean As String
    Dim fileNum As Integer
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Garante que a estrutura .chainsaw existe
    EnsureChainsawFolders
    
    ' SEMPRE USA .chainsaw\logs para todos os documentos
    logFolder = GetChainsawLogsPath() & "\"
    
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
    headerText = headerText & "  Backup: " & IIf(doc.Path = "", "(Desabilitado)", doc.Path & "\backups\") & vbCrLf & vbCrLf
    headerText = headerText & String(80, "=") & vbCrLf & vbCrLf
    
    ' Escreve cabeçalho em UTF-8
    WriteTextUTF8 logFilePath, headerText, False
    
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
Private Function SetAppState(Optional ByVal enabled As Boolean = True, Optional ByVal statusMsg As String = "") As Boolean
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
        Application.StatusBar = "Erro: Tipo não suportado"
        LogMessage "Tipo de documento não suportado: " & doc.Type, LOG_LEVEL_ERROR
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
    EnsureSecondParagraphBlankLines doc
    EnsurePlenarioBlankLines doc
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
    InsertJustificativaBlankLines doc
    LogStepComplete "Ajustes finais de negrito e formatação"
    
    LogStepStart "Formatações especiais (diante do exposto, requeiro)"
    FormatDianteDoExposto doc
    FormatRequeiroParagraphs doc
    FormatPorTodasRazoesParagraphs doc
    LogStepComplete "Formatações especiais (diante do exposto, requeiro)"
    
    LogStepStart "Garantia de espaçamento entre parágrafos longos"
    EnsureBlankLinesBetweenLongParagraphs doc
    LogStepComplete "Garantia de espaçamento entre parágrafos longos"
    
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
    
    LogMessage "Aplicando fonte padrão (modo otimizado com cache)...", LOG_LEVEL_INFO
    
    ' SINGLE PASS - Processa todos os parágrafos em uma passagem usando cache
    For i = 1 To cacheSize
        cache = paragraphCache(i)
        
        ' Pula parágrafos vazios ou com imagens
        If Not cache.needsFormatting Then
            GoTo NextParagraph
        End If
        
        Set para = doc.Paragraphs(cache.index)
        
        ' Aplica fonte padrão
        On Error Resume Next
        With para.Range.Font
            .Name = STANDARD_FONT
            .size = STANDARD_FONT_SIZE
            .Color = wdColorAutomatic
            
            ' Remove sublinhado exceto para título (primeiro parágrafo com texto)
            If i > 3 Then
                .Underline = wdUnderlineNone
            End If
            
            ' Remove negrito exceto para parágrafos especiais
            If Not cache.isSpecial Or cache.specialType = "vereador" Then
                .Bold = False
            End If
        End With
        
        If Err.Number = 0 Then
            formattedCount = formattedCount + 1
        Else
            LogMessage "Erro ao formatar parágrafo " & i & ": " & Err.Description, LOG_LEVEL_WARNING
            Err.Clear
        End If
        On Error GoTo ErrorHandler
        
NextParagraph:
        ' Atualiza progresso a cada 500 parágrafos
        If i Mod 500 = 0 Then
            DoEvents ' Permite cancelamento
        End If
    Next i
    
    Dim elapsed As Single
    elapsed = Timer - startTime
    
    LogMessage "Fonte padrão aplicada: " & formattedCount & " parágrafos em " & Format(elapsed, "0.00") & "s", LOG_LEVEL_INFO
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
    
    ' Log otimizado
    If skippedCount > 0 Then
        LogMessage "Fontes formatadas: " & formattedCount & " parágrafos (incluindo " & skippedCount & " com proteção de imagens)"
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
    
    ' Log atualizado para refletir que todos os parágrafos são formatados
    If skippedCount > 0 Then
        LogMessage "Parágrafos formatados: " & formattedCount & " (incluindo " & skippedCount & " com proteção de imagens)"
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
            .leftIndent = CentimetersToPoints(8.89)      ' Recuo à esquerda de 8.89 cm
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
' INSERÇÃO DE NÚMEROS DE PÁGINA NO RODAPÉ + SIGLA AFV
'================================================================================
Private Function InsertFooterStamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim footer As HeaderFooter
    Dim rngAFV As Range
    Dim rngPage As Range
    Dim rngDash As Range
    Dim rngNum As Range
    Dim fPage As Field
    Dim fTotal As Field

    For Each sec In doc.Sections
        Set footer = sec.Footers(wdHeaderFooterPrimary)
        If footer.Exists Then
            footer.LinkToPrevious = False

            ' -------------------------------
            ' Limpa todo o rodapé
            ' -------------------------------
            footer.Range.Delete

            ' -------------------------------
            ' Insere "afv" à esquerda
            ' -------------------------------
            Set rngAFV = footer.Range
            rngAFV.Collapse Direction:=wdCollapseStart
            rngAFV.text = "afv"
            With rngAFV.Font
                .Name = "Arial"
                .size = 6
                .Color = RGB(128, 128, 128)
            End With
            rngAFV.ParagraphFormat.alignment = wdAlignParagraphLeft
            rngAFV.InsertParagraphAfter

            ' -------------------------------
            ' Insere números de página X-Y centralizados
            ' -------------------------------
            ' Cria um parágrafo limpo
            Set rngPage = footer.Range.Paragraphs.Last.Range
            rngPage.text = ""
            rngPage.Collapse Direction:=wdCollapseStart
           
            
            rngPage.text = "Página "
With rngPage.Font
    .Name = "Arial"
    .size = 9
End With
rngPage.Collapse Direction:=wdCollapseEnd


            ' Campo PAGE
            Set fPage = rngPage.Fields.Add(Range:=rngPage, Type:=wdFieldPage)

            ' De "de"
            Set rngDash = footer.Range.Paragraphs.Last.Range
            rngDash.Collapse Direction:=wdCollapseEnd
            rngDash.text = " de "

            ' Campo NUMPAGES
            Set rngNum = footer.Range.Paragraphs.Last.Range
            rngNum.Collapse Direction:=wdCollapseEnd
            Set fTotal = rngNum.Fields.Add(Range:=rngNum, Type:=wdFieldNumPages)

            ' Centraliza os números de página
            rngPage.ParagraphFormat.alignment = wdAlignParagraphCenter
            
            ' Formata os campos de número de página
            With fPage.result
             .Font.Name = "Arial"
            .Font.size = 9
            End With

            With fTotal.result
             .Font.Name = "Arial"
             .Font.size = 9
            End With

        End If
    Next sec

    InsertFooterStamp = True
    Exit Function

ErrorHandler:
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
' VALIDAÇÃO DE ESTRUTURA DO DOCUMENTO
'================================================================================
Private Function ValidateDocumentStructure(doc As Document) As Boolean
    On Error Resume Next
    
    ' Verificação básica e rápida
    If doc.Range.End > 0 And doc.Sections.count > 0 Then
        ValidateDocumentStructure = True
    Else
        LogMessage "Documento com estrutura inconsistente", LOG_LEVEL_WARNING
        ValidateDocumentStructure = False
    End If
End Function

'================================================================================
' VALIDAÇÃO DE CONSISTÊNCIA DE ENDEREÇOS
'================================================================================
Private Function ValidateAddressConsistency(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim textualParaCount As Long
    Dim secondTextualPara As Paragraph
    Dim firstTextualParaBelowEmenta As Paragraph
    Dim para2Text As String
    Dim para3Text As String
    Dim ruaPosition As Long
    Dim twoWords As String
    Dim word1 As String, word2 As String
    Dim i As Long
    
    textualParaCount = 0
    Set secondTextualPara = Nothing
    Set firstTextualParaBelowEmenta = Nothing
    
    ' Identifica o 2º parágrafo textual (ementa) e o 1º abaixo dele
    For Each para In doc.Paragraphs
        If Len(Trim(para.Range.text)) > 1 Then ' > 1 para ignorar apenas marca de parágrafo
            textualParaCount = textualParaCount + 1
            
            If textualParaCount = 2 Then
                Set secondTextualPara = para
            ElseIf textualParaCount = 3 Then
                ' Pula o 3º (geralmente data/local)
                ' Nada a fazer aqui
            ElseIf textualParaCount = 4 Then
                ' Este é o 1º parágrafo textual abaixo da ementa
                Set firstTextualParaBelowEmenta = para
                Exit For
            End If
        End If
    Next para
    
    ' Se não encontrou os parágrafos necessários, retorna True (sem verificação)
    If secondTextualPara Is Nothing Or firstTextualParaBelowEmenta Is Nothing Then
        ValidateAddressConsistency = True
        Exit Function
    End If
    
    para2Text = secondTextualPara.Range.text
    para3Text = firstTextualParaBelowEmenta.Range.text
    
    ' Procura pela palavra "Rua" (case insensitive) no segundo parágrafo (ementa)
    ruaPosition = InStr(1, para2Text, "rua", vbTextCompare)
    
    If ruaPosition = 0 Then
        ' Não encontrou "Rua", não há o que verificar
        ValidateAddressConsistency = True
        Exit Function
    End If
    
    ' Extrai o texto após "Rua"
    Dim textAfterRua As String
    textAfterRua = Mid(para2Text, ruaPosition + 3) ' +3 para pular "Rua"
    textAfterRua = Trim(textAfterRua)
    
    ' Remove caracteres de pontuação e quebras de linha
    textAfterRua = Replace(textAfterRua, vbCr, " ")
    textAfterRua = Replace(textAfterRua, vbLf, " ")
    textAfterRua = Replace(textAfterRua, vbTab, " ")
    textAfterRua = Replace(textAfterRua, ",", " ")
    textAfterRua = Replace(textAfterRua, ".", " ")
    textAfterRua = Replace(textAfterRua, ";", " ")
    textAfterRua = Replace(textAfterRua, ":", " ")
    
    ' Remove múltiplos espaços com proteção
    Dim spaceCounter As Long
    spaceCounter = 0
    Do While InStr(textAfterRua, "  ") > 0 And spaceCounter < MAX_LOOP_ITERATIONS
        textAfterRua = Replace(textAfterRua, "  ", " ")
        spaceCounter = spaceCounter + 1
    Loop
    textAfterRua = Trim(textAfterRua)
    
    ' Extrai as DUAS primeiras palavras/números após "Rua"
    Dim words() As String
    words = Split(textAfterRua, " ")
    
    If UBound(words) < 1 Then
        ' Não há duas palavras subsequentes, não há o que verificar
        ValidateAddressConsistency = True
        Exit Function
    End If
    
    word1 = Trim(words(0))
    word2 = Trim(words(1))
    
    ' Remove caracteres especiais das palavras
    word1 = Replace(word1, Chr(13), "")
    word1 = Replace(word1, Chr(10), "")
    word2 = Replace(word2, Chr(13), "")
    word2 = Replace(word2, Chr(10), "")
    
    ' Ignora palavras muito curtas (preposições, artigos)
    If Len(word1) <= 2 Then
        ' Se a primeira palavra é muito curta (ex: "de", "do"), usa a próxima
        If UBound(words) >= 2 Then
            word1 = word2
            word2 = Trim(words(2))
            word2 = Replace(word2, Chr(13), "")
            word2 = Replace(word2, Chr(10), "")
        End If
    End If
    
    ' Normaliza o texto do parágrafo textual para comparação mais flexível
    Dim normalizedPara3Text As String
    normalizedPara3Text = para3Text
    normalizedPara3Text = Replace(normalizedPara3Text, "n.º", " ")
    normalizedPara3Text = Replace(normalizedPara3Text, "nº", " ")
    normalizedPara3Text = Replace(normalizedPara3Text, "n°", " ")
    normalizedPara3Text = Replace(normalizedPara3Text, "número", " ")
    normalizedPara3Text = Replace(normalizedPara3Text, ",", " ")
    normalizedPara3Text = Replace(normalizedPara3Text, ".", " ")
    
    ' Verifica se as DUAS palavras existem no primeiro parágrafo textual abaixo da ementa (case insensitive)
    Dim foundWord1 As Boolean
    Dim foundWord2 As Boolean
    
    ' Busca com contexto "Rua" próximo para reduzir falsos positivos
    Dim ruaPosInPara3 As Long
    ruaPosInPara3 = InStr(1, normalizedPara3Text, "rua", vbTextCompare)
    
    If ruaPosInPara3 > 0 Then
        ' Extrai contexto de 100 caracteres após "Rua" no parágrafo textual
        Dim contextAfterRua As String
        contextAfterRua = Mid(normalizedPara3Text, ruaPosInPara3, 100)
        
        ' Busca as palavras no contexto próximo a "Rua"
        foundWord1 = InStr(1, contextAfterRua, word1, vbTextCompare) > 0
        foundWord2 = InStr(1, contextAfterRua, word2, vbTextCompare) > 0
    Else
        ' Se não encontrou "Rua" no texto, busca as palavras em todo o parágrafo
        foundWord1 = InStr(1, normalizedPara3Text, word1, vbTextCompare) > 0
        foundWord2 = InStr(1, normalizedPara3Text, word2, vbTextCompare) > 0
    End If
    
    ' Se as duas palavras não foram encontradas, exibe recomendação
    If Not (foundWord1 And foundWord2) Then
        Dim msg As String
        msg = "VERIFICAR ENDEREÇO" & vbCrLf & vbCrLf
        msg = msg & "Possível inconsistência entre ementa e texto." & vbCrLf & vbCrLf
        msg = msg & "Ementa (2º parágrafo): " & word1 & " " & word2 & vbCrLf & vbCrLf
        msg = msg & "Texto (1º parágrafo):" & vbCrLf
        msg = msg & "  • " & word1 & ": " & IIf(foundWord1, "Sim", "NÃO") & vbCrLf
        msg = msg & "  • " & word2 & ": " & IIf(foundWord2, "Sim", "NÃO") & vbCrLf & vbCrLf
        msg = msg & "Verifique a consistência dos endereços."
        
        MsgBox msg, vbExclamation, "Verificação de Endereço"
        
        LogMessage "Inconsistência de endereço detectada: '" & word1 & " " & word2 & "' não encontrado completamente no 1º parágrafo textual", LOG_LEVEL_WARNING
        
        ValidateAddressConsistency = False
        Exit Function
    End If
    
    ' Tudo OK, endereços consistentes
    LogMessage "Endereços validados com sucesso: ementa x 1º parágrafo textual", LOG_LEVEL_INFO
    ValidateAddressConsistency = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao validar consistência de endereços: " & Err.Description, LOG_LEVEL_WARNING
    ValidateAddressConsistency = True ' Retorna True para não bloquear o processamento
End Function

'================================================================================
' VERIFICAÇÃO DE DADOS SENSÍVEIS
'================================================================================
Private Function CheckSensitiveData(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim docText As String
    Dim lowerText As String
    Dim foundItems As String
    Dim itemCount As Long
    
    ' Obtém todo o texto do documento
    docText = doc.Range.text
    lowerText = LCase(docText)
    
    foundItems = ""
    itemCount = 0
    
    ' Array com as strings sensíveis a serem verificadas (em minúsculas)
    Dim sensitiveStrings() As String
    Dim sensitiveLabels() As String
    Dim i As Long
    
    ' Define as strings a serem buscadas e seus rótulos para exibição
    ReDim sensitiveStrings(11)
    ReDim sensitiveLabels(11)
    
    sensitiveStrings(0) = "cpf:"
    sensitiveLabels(0) = "CPF:"
    
    sensitiveStrings(1) = "cpf n°"
    sensitiveLabels(1) = "CPF n°"
    
    sensitiveStrings(2) = "rg:"
    sensitiveLabels(2) = "RG:"
    
    sensitiveStrings(3) = "rg n°"
    sensitiveLabels(3) = "RG n°"
    
    sensitiveStrings(4) = "nome da mãe:"
    sensitiveLabels(4) = "Nome da mãe:"
    
    sensitiveStrings(5) = "nascimento:"
    sensitiveLabels(5) = "Nascimento:"
    
    sensitiveStrings(6) = "naturalidade:"
    sensitiveLabels(6) = "Naturalidade:"
    
    sensitiveStrings(7) = "estado civil:"
    sensitiveLabels(7) = "Estado civil:"
    
    sensitiveStrings(8) = "placa:"
    sensitiveLabels(8) = "Placa:"
    
    sensitiveStrings(9) = "placa n°"
    sensitiveLabels(9) = "Placa n°"
    
    sensitiveStrings(10) = "renavam:"
    sensitiveLabels(10) = "Renavam:"
    
    sensitiveStrings(11) = "renavam n°"
    sensitiveLabels(11) = "Renavam n°"
    
    ' Verifica cada string sensível
    For i = LBound(sensitiveStrings) To UBound(sensitiveStrings)
        If InStr(1, lowerText, sensitiveStrings(i), vbTextCompare) > 0 Then
            If foundItems <> "" Then
                foundItems = foundItems & ", "
            End If
            foundItems = foundItems & sensitiveLabels(i)
            itemCount = itemCount + 1
        End If
    Next i
    
    ' Se encontrou dados sensíveis, exibe mensagem de aviso
    If itemCount > 0 Then
        Dim msg As String
        msg = "DADOS SENSÍVEIS DETECTADOS" & vbCrLf & vbCrLf
        msg = msg & "Encontrados " & itemCount & " campo(s):" & vbCrLf
        msg = msg & foundItems & vbCrLf & vbCrLf
        msg = msg & "AÇÃO:" & vbCrLf
        msg = msg & "Verifique se há CPF, RG, filiação, etc." & vbCrLf
        msg = msg & "Remova ou anonimize antes da publicação." & vbCrLf & vbCrLf
        msg = msg & "LGPD: Dados sensíveis exigem cuidado especial."
        
        MsgBox msg, vbExclamation, "Verificação de Dados Sensíveis"
        
        LogMessage "Possíveis dados sensíveis detectados: " & foundItems, LOG_LEVEL_WARNING
        
        CheckSensitiveData = False ' Retorna False para indicar que dados foram encontrados
        Exit Function
    End If
    
    ' Nenhum dado sensível encontrado
    LogMessage "Verificação de dados sensíveis concluída - nenhum campo sensível detectado", LOG_LEVEL_INFO
    CheckSensitiveData = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao verificar dados sensíveis: " & Err.Description, LOG_LEVEL_WARNING
    CheckSensitiveData = True ' Retorna True para não bloquear o processamento
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
        
        For Each para In doc.Paragraphs
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
                ' FORMATAÇÃO CONSOLIDADA: Aplica todas as configurações em uma única operação
                With para.Range
                    ' Reset completo de fonte em uma única operação
                    With .Font
                        .Reset
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                        .Color = wdColorAutomatic
                        .Bold = False
                        .Italic = False
                        .Underline = wdUnderlineNone
                    End With
                    
                    ' Reset completo de parágrafo em uma única operação
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
                ' OTIMIZADO: Para parágrafos com imagens, formatação protegida mais rápida
                Call FormatCharacterByCharacter(para, STANDARD_FONT, STANDARD_FONT_SIZE, wdColorAutomatic, True, True)
                paraCount = paraCount + 1
            End If
            
            ' Proteção otimizada contra loops infinitos
            If paraCount Mod 100 = 0 Then DoEvents ' Permite responsividade a cada 100 parágrafos
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
    
    ' OTIMIZADO: Reset de estilos em uma única passada
    For Each para In doc.Paragraphs
        On Error Resume Next
        para.Style = "Normal"
        styleResetCount = styleResetCount + 1
        ' Otimização: Permite responsividade e proteção contra loops
        If styleResetCount Mod 50 = 0 Then DoEvents
        If styleResetCount > 1000 Then Exit For
        On Error GoTo ErrorHandler
    Next para
    
    LogMessage "Formatação limpa: " & paraCount & " parágrafos resetados", LOG_LEVEL_INFO
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
    If isProposition And UBound(words) >= 0 Then
        ' Reconstrói o texto substituindo a última palavra com validação
        newText = ""
        If UBound(words) > 0 Then ' Verifica se há palavras suficientes
            For i = 0 To UBound(words) - 1
                If i <= UBound(words) Then ' Validação adicional
                    If i > 0 Then newText = newText & " "
                    newText = newText & words(i)
                End If
            Next i
        End If
        
        ' Adiciona $NUMERO$/$ANO$ no lugar da última palavra
        If newText <> "" Then newText = newText & " "
        newText = newText & "$NUMERO$/$ANO$"
    Else
        ' Se não for proposição, mantém o texto original
        newText = paraText
    End If
    
    ' SEMPRE aplica formatação de título: caixa alta, negrito, sublinhado
    firstPara.Range.text = UCase(newText) & vbCrLf
    
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
' FORMATAÇÃO DE PARÁGRAFOS "CONSIDERANDO"
'================================================================================
Private Function FormatConsiderandoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim rng As Range
    Dim totalFormatted As Long
    Dim i As Long
    
    ' Percorre todos os parágrafos procurando por "considerando" no início
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Verifica se o parágrafo começa com "considerando" (ignorando maiúsculas/minúsculas)
        If Len(paraText) >= 12 And LCase(Left(paraText, 12)) = "considerando" Then
            ' Verifica se após "considerando" vem espaço, vírgula, ponto-e-vírgula ou fim da linha
            Dim nextChar As String
            If Len(paraText) > 12 Then
                nextChar = Mid(paraText, 13, 1)
                If nextChar = " " Or nextChar = "," Or nextChar = ";" Or nextChar = ":" Then
                    ' É realmente "considerando" no início do parágrafo
                    Set rng = para.Range
                    
                    ' CORREÇÃO: Usa Find/Replace para preservar espaçamento
                    With rng.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .text = "considerando"
                        .Replacement.text = "CONSIDERANDO"
                        .Replacement.Font.Bold = True
                        .MatchCase = False
                        .MatchWholeWord = False  ' CORREÇÃO: False para não exigir palavra completa
                        .Forward = True
                        .Wrap = wdFindStop
                        
                        ' Limita a busca ao início do parágrafo
                        rng.End = rng.Start + 15  ' Seleciona apenas o início para evitar múltiplas substituições
                        
                        If .Execute(Replace:=True) Then
                            totalFormatted = totalFormatted + 1
                        End If
                    End With
                End If
            Else
                ' Parágrafo contém apenas "considerando"
                Set rng = para.Range
                rng.End = rng.Start + 12
                
                With rng
                    .text = "CONSIDERANDO"
                    .Font.Bold = True
                End With
                
                totalFormatted = totalFormatted + 1
            End If
        End If
    Next i
    
    LogMessage "Formatação 'considerando' aplicada: " & totalFormatted & " ocorrências em negrito e caixa alta", LOG_LEVEL_INFO
    FormatConsiderandoParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação 'considerando': " & Err.Description, LOG_LEVEL_ERROR
    FormatConsiderandoParagraphs = False
End Function

'================================================================================
' APLICAÇÃO DE SUBSTITUIÇÕES DE TEXTO
'================================================================================
Private Function ApplyTextReplacements(doc As Document) As Boolean
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
    Dim totalReplacements As Long
    totalReplacements = 0
    
    ' Funcionalidade 10: Substitui variantes de "d'Oeste"
    Dim dOesteVariants() As String
    Dim i As Long
    
    ' Define as variantes possíveis dos 3 primeiros caracteres de "d'Oeste"
    ReDim dOesteVariants(0 To 15)
    dOesteVariants(0) = "d'O"   ' Original
    dOesteVariants(1) = "d´O"   ' Acento agudo
    dOesteVariants(2) = "d`O"   ' Acento grave
    dOesteVariants(3) = "d" & Chr(8220) & "O"   ' Aspas curvas esquerda
    dOesteVariants(4) = "d'o"   ' Minúscula
    dOesteVariants(5) = "d´o"
    dOesteVariants(6) = "d`o"
    dOesteVariants(7) = "d" & Chr(8220) & "o"
    dOesteVariants(8) = "D'O"   ' Maiúscula no D
    dOesteVariants(9) = "D´O"
    dOesteVariants(10) = "D`O"
    dOesteVariants(11) = "D" & Chr(8220) & "O"
    dOesteVariants(12) = "D'o"
    dOesteVariants(13) = "D´o"
    dOesteVariants(14) = "D`o"
    dOesteVariants(15) = "D" & Chr(8220) & "o"
    
    ' Processa cada variante de forma segura
    For i = 0 To UBound(dOesteVariants)
        On Error Resume Next
        
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
            
            ' Executa a substituição e conta
            replacementCount = .Execute(Replace:=wdReplaceAll)
            
            ' Verifica se houve erro
            If Err.Number = 0 Then
                ' Conta quantas substituições foram feitas
                If replacementCount Then
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
    On Error Resume Next
    Set rng = Nothing
    Set rng = doc.Range
    If Not rng Is Nothing Then
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = " ao Setor, "
            .Replacement.text = " ao setor competente"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            
            replacementCount = .Execute(Replace:=wdReplaceAll)
            If Err.Number = 0 And replacementCount Then
                LogMessage "Substituição aplicada: ' ao Setor, ' ? ' ao setor competente'", LOG_LEVEL_INFO
            End If
        End With
    End If
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' Funcionalidade 12: Substitui " Setor Competente " por " setor competente " (case insensitive)
    On Error Resume Next
    Set rng = Nothing
    Set rng = doc.Range
    If Not rng Is Nothing Then
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = " Setor Competente "
            .Replacement.text = " setor competente "
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            
            replacementCount = .Execute(Replace:=wdReplaceAll)
            If Err.Number = 0 And replacementCount Then
                LogMessage "Substituição aplicada: ' Setor Competente ' ? ' setor competente '", LOG_LEVEL_INFO
            End If
        End With
    End If
    Err.Clear
    On Error GoTo ErrorHandler
    
    ApplyTextReplacements = True
    Exit Function

ErrorHandler:
    LogMessage "Erro crítico nas substituições de texto: " & Err.Description & " (Variante: " & i & ")", LOG_LEVEL_ERROR
    ApplyTextReplacements = False
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
    For Each para In doc.Paragraphs
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
        Exit Sub ' Não encontrou "Justificativa"
    End If
    
    ' FASE 2: Remove TODAS as linhas vazias ANTES de "Justificativa"
    i = justificativaIndex - 1
    Do While i >= 1
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        If paraText = "" And Not HasVisualContent(para) Then
            ' Remove linha vazia
            para.Range.Delete
            justificativaIndex = justificativaIndex - 1 ' Ajusta índice
            i = i - 1
        Else
            ' Encontrou conteúdo, para de remover
            Exit Do
        End If
    Loop
    
    ' FASE 3: Remove TODAS as linhas vazias DEPOIS de "Justificativa"
    i = justificativaIndex + 1
    Do While i <= doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        If paraText = "" And Not HasVisualContent(para) Then
            ' Remove linha vazia
            para.Range.Delete
            ' Não incrementa i pois removemos o parágrafo
        Else
            ' Encontrou conteúdo, para de remover
            Exit Do
        End If
    Loop
    
    ' FASE 4: Insere EXATAMENTE 2 linhas em branco ANTES
    Set para = doc.Paragraphs(justificativaIndex)
    para.Range.InsertParagraphBefore
    para.Range.InsertParagraphBefore
    
    ' FASE 5: Insere EXATAMENTE 2 linhas em branco DEPOIS
    ' Atualiza referência após inserções anteriores
    Set para = doc.Paragraphs(justificativaIndex + 2) ' +2 porque inserimos 2 antes
    para.Range.InsertParagraphAfter
    para.Range.InsertParagraphAfter
    
    LogMessage "Linhas em branco ajustadas: exatamente 2 antes e 2 depois de 'Justificativa'", LOG_LEVEL_INFO
    
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
        ' Remove TODAS as linhas vazias ANTES de "Plenário..."
        i = plenarioIndex - 1
        Do While i >= 1
            Set para = doc.Paragraphs(i)
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            
            If paraText = "" And Not HasVisualContent(para) Then
                ' Remove linha vazia
                para.Range.Delete
                plenarioIndex = plenarioIndex - 1 ' Ajusta índice
                i = i - 1
            Else
                ' Encontrou conteúdo, para de remover
                Exit Do
            End If
        Loop
        
        ' Remove TODAS as linhas vazias DEPOIS de "Plenário..."
        i = plenarioIndex + 1
        Do While i <= doc.Paragraphs.count
            Set para = doc.Paragraphs(i)
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            
            If paraText = "" And Not HasVisualContent(para) Then
                ' Remove linha vazia
                para.Range.Delete
                ' Não incrementa i pois removemos o parágrafo
            Else
                ' Encontrou conteúdo, para de remover
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
        
        LogMessage "2 linhas em branco inseridas antes e depois de 'Plenário Dr. Tancredo Neves'", LOG_LEVEL_INFO
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
        ' FASE 8: Remove TODAS as linhas vazias DEPOIS de "Excelentíssimo..."
        i = prefeitoIndex + 1
        Do While i <= doc.Paragraphs.count
            Set para = doc.Paragraphs(i)
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            
            If paraText = "" And Not HasVisualContent(para) Then
                ' Remove linha vazia
                para.Range.Delete
                ' Não incrementa i pois removemos o parágrafo
            Else
                ' Encontrou conteúdo, para de remover
                Exit Do
            End If
        Loop
        
        ' FASE 9: Insere EXATAMENTE 2 linhas em branco DEPOIS
        Set para = doc.Paragraphs(prefeitoIndex)
        para.Range.InsertParagraphAfter
        para.Range.InsertParagraphAfter
        
        LogMessage "2 linhas em branco inseridas após 'Excelentíssimo Senhor Prefeito Municipal,'", LOG_LEVEL_INFO
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
    For Each para In doc.Paragraphs
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
    For Each para In doc.Paragraphs
        If Not HasVisualContent(para) Then
            ' Obtém o texto do parágrafo
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            cleanText = LCase(paraText)
            
            ' Verifica se começa com "requeiro" (8 caracteres)
            If Len(paraText) >= 8 Then
                If Left(cleanText, 8) = "requeiro" Then
                    ' Aplica formatação APENAS à palavra "requeiro": negrito e caixa alta
                    Dim wordRange As Range
                    Set wordRange = para.Range.Duplicate
                    wordRange.Collapse wdCollapseStart
                    wordRange.MoveEnd wdCharacter, 8 ' Seleciona apenas "requeiro"
                    
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
' SUBROTINA PÚBLICA: ABRIR PASTA DE LOGS E BACKUPS
'================================================================================
Public Sub AbrirPastaLogsEBackups()
    On Error GoTo ErrorHandler
    
    Dim chainsawFolder As String
    
    ' Garante que a estrutura .chainsaw existe
    EnsureChainsawFolders
    
    ' Caminho da pasta .chainsaw
    chainsawFolder = GetChainsawTempPath()
    
    ' SEMPRE abre pasta .chainsaw
    Application.StatusBar = "Abrindo pasta .chainsaw"
    shell "explorer.exe """ & chainsawFolder & """", vbNormalFocus
    
    ' Log da operação se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Pasta .chainsaw aberta pelo usuário: " & chainsawFolder, LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir pasta"
    LogMessage "Erro ao abrir pasta .chainsaw: " & Err.Description, LOG_LEVEL_ERROR
    
    ' Fallback: tenta abrir novamente
    On Error Resume Next
    shell "explorer.exe """ & GetChainsawTempPath() & """", vbNormalFocus
    Application.StatusBar = "Pasta .chainsaw aberta"
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
    
    ' USA A FUNÇÃO QUE JÁ TEM A LÓGICA CORRETA (.chainsaw para não salvos)
    backupFolder = EnsureBackupDirectory(doc)
    
    ' Extrai nome e extensão do documento
    docName = fso.GetBaseName(doc.Name)
    docExtension = fso.GetExtensionName(doc.Name)
    
    ' Cria timestamp para o backup
    timeStamp = Format(Now, "yyyy-mm-dd_HHmmss")
    
    ' Nome do arquivo de backup
    backupFileName = docName & "_backup_" & timeStamp & "." & docExtension
    backupFilePath = backupFolder & "\" & backupFileName
    
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
' LIMPEZA DE BACKUPS ANTIGOS
'================================================================================
Private Sub CleanOldBackups(backupFolder As String, docBaseName As String)
    On Error Resume Next
    
    ' Limpeza simplificada - só remove se houver muitos arquivos
    Dim fso As Object
    Dim folder As Object
    Dim filesCount As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(backupFolder)
    
    filesCount = folder.Files.count
    
    ' Se há mais de 15 arquivos na pasta de backup, registra aviso
    If filesCount > 15 Then
        LogMessage "Muitos backups na pasta (" & filesCount & " arquivos) - considere limpeza manual", LOG_LEVEL_WARNING
    End If
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
    
    LogMessage "Limpeza de espaços concluída: " & spacesRemoved & " correções aplicadas (com proteção CONSIDERANDO)", LOG_LEVEL_INFO
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
    
    ' Conta todas as imagens primeiro
    Dim totalImages As Long
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        totalImages = totalImages + para.Range.InlineShapes.count
    Next i
    
    ' Adiciona shapes flutuantes
    totalImages = totalImages + doc.Shapes.count
    
    ' Redimensiona array se necessário
    If totalImages > 0 Then
        ReDim savedImages(totalImages - 1)
        
        ' Backup de imagens inline - apenas propriedades críticas
        For i = 1 To doc.Paragraphs.count
            Set para = doc.Paragraphs(i)
            
            For j = 1 To para.Range.InlineShapes.count
                Set shape = para.Range.InlineShapes(j)
                
                ' Salva apenas propriedades essenciais para proteção
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
    
    ' Percorre todos os parágrafos
    For Each para In doc.Paragraphs
        ' Verifica se o parágrafo contém imagens inline
        If para.Range.InlineShapes.count > 0 Then
            ' Zera o recuo à esquerda
            With para.Format
                .leftIndent = 0
                .firstLineIndent = 0
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
' CENTER IMAGE AFTER PLENARIO - Centraliza imagem entre 5ª e 7ª linha após Plenário
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
' BACKUP LIST FORMATS - Salva formatações de lista antes do processamento
'================================================================================
Private Function BackupListFormats(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim i As Long
    Dim tempListInfo As ListFormatInfo
    
    listFormatCount = 0
    ReDim savedListFormats(0)
    
    ' Conta quantos parágrafos têm formatação de lista
    Dim totalLists As Long
    totalLists = 0
    For Each para In doc.Paragraphs
        If para.Range.ListFormat.ListType <> wdListNoNumbering Then
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
    
    ' Salva informações de cada parágrafo com lista
    i = 1
    For Each para In doc.Paragraphs
        If para.Range.ListFormat.ListType <> wdListNoNumbering Then
            With tempListInfo
                .paraIndex = i
                .HasList = True
                .ListType = para.Range.ListFormat.ListType
                
                ' Salva o nível da lista se aplicável
                On Error Resume Next
                .ListLevelNumber = para.Range.ListFormat.ListLevelNumber
                If Err.Number <> 0 Then
                    .ListLevelNumber = 1
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
                
                ' Salva a string da lista (marcador ou número)
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
    
    LogMessage "Formatações de lista salvas: " & listFormatCount & " parágrafos com lista", LOG_LEVEL_INFO
    BackupListFormats = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao salvar formatações de lista: " & Err.Description, LOG_LEVEL_WARNING
    BackupListFormats = False
End Function

'================================================================================
' RESTORE LIST FORMATS - Restaura formatações de lista após o processamento
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
                
                ' Remove qualquer formatação de lista existente primeiro
                para.Range.ListFormat.RemoveNumbers
                
                ' Aplica a formatação de lista original
                Select Case .ListType
                    Case wdListBullet
                        ' Lista com marcadores
                        para.Range.ListFormat.ApplyBulletDefault
                        
                    Case wdListSimpleNumbering, wdListListNumOnly
                        ' Lista numerada simples
                        para.Range.ListFormat.ApplyNumberDefault
                        
                    Case wdListMixedNumbering
                        ' Lista com numeração mista
                        para.Range.ListFormat.ApplyNumberDefault
                        
                    Case wdListOutlineNumbering
                        ' Lista com numeração de tópicos
                        para.Range.ListFormat.ApplyOutlineNumberDefault
                        
                    Case Else
                        ' Tenta aplicar formatação padrão
                        If InStr(.ListString, ".") > 0 Or IsNumeric(Left(.ListString, 1)) Then
                            para.Range.ListFormat.ApplyNumberDefault
                        Else
                            para.Range.ListFormat.ApplyBulletDefault
                        End If
                End Select
                
                ' Tenta restaurar o nível da lista
                If .ListLevelNumber > 0 And .ListLevelNumber <= 9 Then
                    para.Range.ListFormat.ListLevelNumber = .ListLevelNumber
                End If
                
                If Err.Number = 0 Then
                    restoredCount = restoredCount + 1
                Else
                    LogMessage "Aviso: Não foi possível restaurar lista no parágrafo " & .paraIndex & ": " & Err.Description, LOG_LEVEL_WARNING
                    Err.Clear
                End If
            End If
        End With
        
        On Error GoTo ErrorHandler
    Next i
    
    If restoredCount > 0 Then
        LogMessage "Formatações de lista restauradas: " & restoredCount & " parágrafos", LOG_LEVEL_INFO
    End If
    
    ' Limpa o array
    ReDim savedListFormats(0)
    listFormatCount = 0
    
    RestoreListFormats = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao restaurar formatações de lista: " & Err.Description, LOG_LEVEL_WARNING
    RestoreListFormats = False
End Function

'================================================================================
' FORMAT NUMBERED PARAGRAPHS INDENT - Aplica recuo de lista em parágrafos iniciados com número
'================================================================================
Private Function FormatNumberedParagraphsIndent(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim firstChar As String
    Dim formattedCount As Long
    Dim defaultIndent As Single
    
    formattedCount = 0
    
    ' Obtém o recuo padrão de uma lista numerada (aproximadamente 36 pontos ou 1.27 cm)
    ' Esse é o recuo padrão do Word para listas numeradas
    defaultIndent = 36 ' pontos
    
    ' Percorre todos os parágrafos
    For Each para In doc.Paragraphs
        paraText = Trim(para.Range.text)
        
        ' Verifica se o parágrafo não está vazio
        If Len(paraText) > 0 Then
            ' Pega o primeiro caractere
            firstChar = Left(paraText, 1)
            
            ' Verifica se o primeiro caractere é um algarismo (0-9)
            If IsNumeric(firstChar) Then
                ' Verifica se o parágrafo não tem formatação de lista já aplicada
                ' (para não sobrescrever listas reais restauradas)
                If para.Range.ListFormat.ListType = wdListNoNumbering Then
                    ' Aplica o recuo à esquerda igual ao de uma lista numerada
                    With para.Format
                        .leftIndent = defaultIndent
                        ' Também pode ajustar o firstLineIndent se necessário
                        ' Para listas numeradas, geralmente é negativo para criar o "hanging indent"
                        .firstLineIndent = 0
                    End With
                    formattedCount = formattedCount + 1
                End If
            End If
        End If
    Next para
    
    If formattedCount > 0 Then
        LogMessage "Parágrafos iniciados com número formatados com recuo de lista: " & formattedCount, LOG_LEVEL_INFO
    End If
    
    FormatNumberedParagraphsIndent = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao formatar recuos de parágrafos numerados: " & Err.Description, LOG_LEVEL_WARNING
    FormatNumberedParagraphsIndent = False
End Function

'================================================================================
' FORMAT BULLETED PARAGRAPHS INDENT - Aplica recuo de lista em parágrafos iniciados com marcadores
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
    
    ' Obtém o recuo padrão de uma lista com marcadores (aproximadamente 36 pontos ou 1.27 cm)
    defaultIndent = 36 ' pontos
    
    ' Array com os marcadores mais comuns
    Dim bulletMarkers() As String
    bulletMarkers = Split("*,-,•,?,?,¦,?,?,?,–,—,?,>,+,~,·,?,?", ",")
    
    ' Percorre todos os parágrafos
    For Each para In doc.Paragraphs
        paraText = Trim(para.Range.text)
        
        ' Verifica se o parágrafo não está vazio
        If Len(paraText) > 0 Then
            ' Pega o primeiro caractere
            firstChar = Left(paraText, 1)
            
            ' Verifica se o primeiro caractere é um marcador comum
            Dim isBullet As Boolean
            isBullet = False
            
            For i = LBound(bulletMarkers) To UBound(bulletMarkers)
                If firstChar = bulletMarkers(i) Then
                    isBullet = True
                    Exit For
                End If
            Next i
            
            If isBullet Then
                ' Verifica se o parágrafo não tem formatação de lista já aplicada
                ' (para não sobrescrever listas reais restauradas)
                If para.Range.ListFormat.ListType = wdListNoNumbering Then
                    ' Aplica o recuo à esquerda igual ao de uma lista com marcadores
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
        LogMessage "Parágrafos iniciados com marcador formatados com recuo de lista: " & formattedCount, LOG_LEVEL_INFO
    End If
    
    FormatBulletedParagraphsIndent = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao formatar recuos de parágrafos com marcadores: " & Err.Description, LOG_LEVEL_WARNING
    FormatBulletedParagraphsIndent = False
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
    
    ' Processa cada parágrafo
    For Each para In doc.Paragraphs
        matchCount = 0
        
        ' Pula parágrafos muito longos
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
    
    ' Garante que a estrutura .chainsaw existe
    EnsureChainsawFolders
    
    ' SEMPRE USA .chainsaw\backups para todos os documentos
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
    ' Retorna pasta do documento ou .chainsaw como fallback
    If doc.Path <> "" Then
        EnsureBackupDirectory = doc.Path
    Else
        EnsureBackupDirectory = GetChainsawTempPath()
    End If
End Function











