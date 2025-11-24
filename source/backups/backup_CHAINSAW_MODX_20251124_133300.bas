Attribute VB_Name = "CHAINSAW_MODX"
' =============================================================================
' CHAINSAW - Sistema de Padronização de Proposituras Legislativas
' =============================================================================
' Versão: 1.0-RC1-202511050239
' Licença: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
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

Private Const TOP_MARGIN_CM As Double = 4.6
Private Const BOTTOM_MARGIN_CM As Double = 2
Private Const LEFT_MARGIN_CM As Double = 3
Private Const RIGHT_MARGIN_CM As Double = 3
Private Const HEADER_DISTANCE_CM As Double = 0.3
Private Const FOOTER_DISTANCE_CM As Double = 0.9

Private Const HEADER_IMAGE_RELATIVE_PATH As String = "\Documentos\CHAINSAW\assets\stamp.png"
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 21
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.7
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.19

'================================================================================
' CONSTANTES DE SISTEMA
'================================================================================
Private Const MIN_SUPPORTED_VERSION As Long = 14
Private Const REQUIRED_STRING As String = "$NUMERO$/$ANO$"
' BACKUP_FOLDER_NAME removida - backups são salvos na mesma pasta do documento
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
' PONTO DE ENTRADA PRINCIPAL
'================================================================================
Public Sub PadronizarDocumentoMain()
    On Error GoTo CriticalErrorHandler
    
    executionStartTime = Now
    formattingCancelled = False
    
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
    
    ' Valida o tipo de documento (INDICAÇÃO, REQUERIMENTO ou MOÇÃO)
    If Not ValidateDocumentType(doc) Then
        Application.StatusBar = "Cancelado: tipo de documento não reconhecido"
        LogMessage "Processamento cancelado pelo usuário após validação de tipo", LOG_LEVEL_INFO
        Exit Sub
    End If
    
    ' Inicializa barra de progresso (15 etapas principais)
    InitializeProgress 15
    
    StartUndoGroup "Padronização de Documento - " & doc.Name
    
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
            Exit Sub
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

    ' Restaura imagens após formatações
    IncrementProgress "Restaurando imagens"
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
    ClearParagraphCache ' Limpa cache de parágrafos
    SafeCleanup
    CleanupImageProtection ' Nova função para limpar variáveis de proteção de imagens
    CleanupViewSettings    ' Nova função para limpar variáveis de configurações de visualização
    
    If Not SetAppState(True, "Concluído!") Then
        LogMessage "Falha ao restaurar estado da aplicação", LOG_LEVEL_WARNING
    End If
    
    SafeFinalizeLogging
    
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "ERRO CRÍTICO #" & Err.Number & ": " & Err.Description & _
              " em " & Err.Source & " (Linha: " & Erl & ")"
    
    LogMessage errDesc, LOG_LEVEL_ERROR
    Application.StatusBar = "Erro - verificar logs"
    
    ShowUserFriendlyError Err.Number, Err.Description
    EmergencyRecovery
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
    
    MsgBox msg, vbCritical, "CHAINSAW Proposituras v1.0-beta1"
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
    
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
    End If
    
    ' Limpa variáveis de proteção de imagens em caso de erro
    CleanupImageProtection
    
    ' Limpa variáveis de configurações de visualização em caso de erro
    CleanupViewSettings
    
    LogMessage "Recuperação de emergência executada", LOG_LEVEL_ERROR
        undoGroupEnabled = False
    
    CloseAllOpenFiles
End Sub

'================================================================================
' LIMPEZA SEGURA DE RECURSOS
'================================================================================
Private Sub SafeCleanup()
    On Error Resume Next
    
    EndUndoGroup
    
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
End Sub

'================================================================================
' ATUALIZAÇÃO DA BARRA DE PROGRESSO
'================================================================================
Private Sub UpdateProgress(message As String, percentComplete As Long)
    ' Nota: parâmetro 'message' mantido por compatibilidade mas não é exibido
    ' A barra de status mostra apenas a barra visual sem texto descritivo
    
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
    
    ' Atualiza StatusBar apenas com a barra visual (sem texto descritivo)
    Application.StatusBar = progressBar
    
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
' GERENCIAMENTO DE DESFAZER
'================================================================================
Private Sub StartUndoGroup(groupName As String)
    On Error GoTo ErrorHandler
    
    If undoGroupEnabled Then
        EndUndoGroup
    End If
    
    Application.UndoRecord.StartCustomRecord groupName
    undoGroupEnabled = True
    
    Exit Sub
    
ErrorHandler:
    undoGroupEnabled = False
End Sub

Private Sub EndUndoGroup()
    On Error GoTo ErrorHandler
    
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
    End If
    
    Exit Sub
    
ErrorHandler:
    undoGroupEnabled = False
End Sub

'================================================================================
' SISTEMA DE REGISTRO DE LOGS
'================================================================================
Private Function InitializeLogging(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim logFolder As String
    Dim docNameClean As String
    Dim fileNum As Integer
    Dim fso As Object
    
    ' Define o caminho do log na mesma pasta do documento ativo
    If doc.Path <> "" Then
        logFolder = doc.Path & "\"
    Else
        logFolder = Environ("TEMP") & "\"
    End If
    
    ' Sanitiza nome do documento para uso em arquivo
    docNameClean = doc.Name
    docNameClean = Replace(docNameClean, ".doc", "")
    docNameClean = Replace(docNameClean, ".docx", "")
    docNameClean = Replace(docNameClean, ".docm", "")
    docNameClean = SanitizeFileName(docNameClean)
    
    ' Define nome do arquivo de log com timestamp
    logFilePath = logFolder & "CHAINSAW_" & Format(Now, "yyyymmdd_HHmmss") & "_" & docNameClean & ".log"
    
    ' Inicializa contadores e controles
    errorCount = 0
    warningCount = 0
    infoCount = 0
    logBufferEnabled = False
    logBuffer = ""
    lastFlushTime = Now
    logFileHandle = 0
    
    ' Cria arquivo de log com informações de contexto
    fileNum = FreeFile
    logFileHandle = fileNum
    
    Open logFilePath For Output As #fileNum
    
    ' Cabeçalho estruturado
    Print #fileNum, String(80, "=")
    Print #fileNum, "CHAINSAW - LOG DE PROCESSAMENTO DE DOCUMENTO"
    Print #fileNum, String(80, "=")
    Print #fileNum, ""
    Print #fileNum, "[SESSÃO]"
    Print #fileNum, "  Início: " & Format(Now, "dd/mm/yyyy HH:mm:ss")
    Print #fileNum, "  ID: " & Format(Now, "yyyymmddHHmmss")
    Print #fileNum, ""
    Print #fileNum, "[AMBIENTE]"
    Print #fileNum, "  Usuário: " & Environ("USERNAME")
    Print #fileNum, "  Computador: " & Environ("COMPUTERNAME")
    Print #fileNum, "  Domínio: " & Environ("USERDOMAIN")
    Print #fileNum, "  SO: Windows " & GetWindowsVersion()
    Print #fileNum, "  Word: " & Application.version & " (" & GetWordVersionName() & ")"
    Print #fileNum, ""
    Print #fileNum, "[DOCUMENTO]"
    Print #fileNum, "  Nome: " & doc.Name
    Print #fileNum, "  Caminho: " & IIf(doc.Path = "", "(Não salvo)", doc.Path)
    Print #fileNum, "  Tamanho: " & GetDocumentSize(doc)
    Print #fileNum, "  Parágrafos: " & doc.Paragraphs.count
    Print #fileNum, "  Páginas: " & doc.ComputeStatistics(wdStatisticPages)
    Print #fileNum, "  Proteção: " & GetProtectionType(doc)
    Print #fileNum, "  Idioma: " & doc.Range.LanguageID
    Print #fileNum, ""
    Print #fileNum, "[CONFIGURAÇÃO]"
    Print #fileNum, "  Debug: " & IIf(DEBUG_MODE, "Ativado", "Desativado")
    Print #fileNum, "  Log: " & logFilePath
    Print #fileNum, "  Backup: " & IIf(doc.Path = "", "(Desabilitado)", doc.Path & "\")
    Print #fileNum, ""
    Print #fileNum, String(80, "=")
    Print #fileNum, ""
    
    Close #fileNum
    
    loggingEnabled = True
    InitializeLogging = True
    
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If fileNum > 0 Then Close #fileNum
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
        
        fileNum = FreeFile
        Open logFilePath For Append As #fileNum
        If Len(logBuffer) > 0 Then
            Print #fileNum, logBuffer
            logBuffer = ""
        End If
        Print #fileNum, formattedMessage
        Close #fileNum
        
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
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logFilePath For Append As #fileNum
    Print #fileNum, logBuffer
    Close #fileNum
    
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
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open logFilePath For Append As #fileNum
    Print #fileNum, ""
    Print #fileNum, String(80, "-")
    Print #fileNum, "SEÇÃO: " & UCase(sectionName)
    Print #fileNum, String(80, "-")
    Close #fileNum
    
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
    
    ' Escreve rodapé estruturado
    fileNum = FreeFile
    Open logFilePath For Append As #fileNum
    
    Print #fileNum, ""
    Print #fileNum, String(80, "=")
    Print #fileNum, "RESUMO DA SESSÃO"
    Print #fileNum, String(80, "=")
    Print #fileNum, ""
    Print #fileNum, "[STATUS]"
    Print #fileNum, "  Final: " & statusText & " " & statusIcon
    Print #fileNum, "  Término: " & Format(Now, "dd/mm/yyyy HH:mm:ss")
    Print #fileNum, "  Duração: " & durationText
    Print #fileNum, ""
    Print #fileNum, "[ESTATÍSTICAS]"
    Print #fileNum, "  Total de eventos: " & totalEvents
    Print #fileNum, "  Informações: " & infoCount & " (" & Format(infoCount / IIf(totalEvents > 0, totalEvents, 1) * 100, "0.0") & "%)"
    Print #fileNum, "  Avisos: " & warningCount & " (" & Format(warningCount / IIf(totalEvents > 0, totalEvents, 1) * 100, "0.0") & "%)"
    Print #fileNum, "  Erros: " & errorCount & " (" & Format(errorCount / IIf(totalEvents > 0, totalEvents, 1) * 100, "0.0") & "%)"
    Print #fileNum, ""
    
    ' Adiciona informações de performance
    If totalEvents > 0 Then
        Print #fileNum, "[PERFORMANCE]"
        Print #fileNum, "  Eventos/segundo: " & Format(totalEvents / IIf(duration > 0, duration, 1), "0.0")
        Print #fileNum, "  Tempo médio/evento: " & Format((duration / totalEvents) * 1000, "0.0") & "ms"
        Print #fileNum, ""
    End If
    
    ' Recomendações se houver problemas
    If errorCount > 0 Or warningCount > 5 Then
        Print #fileNum, "[RECOMENDAÇÕES]"
        If errorCount > 0 Then
            Print #fileNum, "  • Verifique os erros acima e corrija problemas no documento"
        End If
        If warningCount > 5 Then
            Print #fileNum, "  • Múltiplos avisos detectados - revise o documento manualmente"
        End If
        If duration > 60 Then
            Print #fileNum, "  • Processamento demorado - considere otimizar o documento"
        End If
        Print #fileNum, ""
    End If
    
    Print #fileNum, String(80, "=")
    Print #fileNum, "FIM DO LOG"
    Print #fileNum, String(80, "=")
    
    Close #fileNum
    
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
    
    LogStepStart "Formatação de 'ante o exposto'"
    FormatAnteOExpostoParagraphs doc
    LogStepComplete "Formatação de 'ante o exposto'"
    
    LogStepStart "Formatação de 'por todas as razões aqui expostas'"
    FormatPorTodasRazoesParagraphs doc
    LogStepComplete "Formatação de 'por todas as razões aqui expostas'"
    
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
' GET FIRST WORD OF DOCUMENT - OBTEM PRIMEIRA PALAVRA DO DOCUMENTO
'================================================================================
' Função auxiliar que retorna a primeira palavra do documento (case insensitive)
' Usada para determinar o tipo de documento (INDICAÇÃO, REQUERIMENTO, etc)
Private Function GetFirstWordOfDocument(doc As Document) As String
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim firstWord As String
    Dim i As Long
    
    ' Valor padrão vazio
    GetFirstWordOfDocument = ""
    
    ' Verifica se o documento tem parágrafos
    If doc.Paragraphs.count = 0 Then Exit Function
    
    ' Procura o primeiro parágrafo com conteúdo (pula vazios)
    For i = 1 To doc.Paragraphs.count
        If i > 10 Then Exit For ' Proteção: analisa apenas os primeiros 10 parágrafos
        
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Se encontrou um parágrafo com texto
        If Len(paraText) > 0 Then
            ' Extrai a primeira palavra (tudo antes do primeiro espaço)
            Dim spacePos As Long
            spacePos = InStr(paraText, " ")
            
            If spacePos > 0 Then
                firstWord = Left(paraText, spacePos - 1)
            Else
                firstWord = paraText ' Parágrafo tem apenas uma palavra
            End If
            
            ' Remove pontuação comum no final da palavra
            firstWord = Replace(firstWord, ":", "")
            firstWord = Replace(firstWord, ",", "")
            firstWord = Replace(firstWord, ".", "")
            firstWord = Replace(firstWord, ";", "")
            
            ' Retorna em maiúsculas para comparação case insensitive
            GetFirstWordOfDocument = UCase(Trim(firstWord))
            Exit Function
        End If
    Next i
    
    Exit Function

ErrorHandler:
    LogMessage "Erro ao obter primeira palavra do documento: " & Err.Description, LOG_LEVEL_WARNING
    GetFirstWordOfDocument = ""
End Function

'================================================================================
' VALIDATE DOCUMENT TYPE - VALIDAÇÃO DO TIPO DE DOCUMENTO
'================================================================================
' Valida se o documento é do tipo esperado (INDICAÇÃO, REQUERIMENTO ou MOÇÃO)
' Retorna True para prosseguir, False para cancelar
Private Function ValidateDocumentType(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim firstWord As String
    Dim userResponse As VbMsgBoxResult
    Dim validTypes As String
    
    ' Valor padrão: assume cancelamento
    ValidateDocumentType = False
    
    ' Obtém a primeira palavra do documento
    firstWord = GetFirstWordOfDocument(doc)
    
    ' Se não conseguiu obter a primeira palavra, alerta o usuário
    If Len(firstWord) = 0 Then
        userResponse = MsgBox( _
            "Não foi possível identificar o tipo do documento." & vbCrLf & vbCrLf & _
            "O documento parece estar vazio ou sem texto válido." & vbCrLf & vbCrLf & _
            "Deseja cancelar ou prosseguir mesmo assim?", _
            vbExclamation + vbYesNo, _
            "Tipo de Documento Não Identificado")
        
        If userResponse = vbYes Then
            LogMessage "Usuário optou por prosseguir com documento de tipo não identificado", LOG_LEVEL_WARNING
            ValidateDocumentType = True
        Else
            LogMessage "Usuário cancelou - documento de tipo não identificado", LOG_LEVEL_INFO
            ValidateDocumentType = False
        End If
        Exit Function
    End If
    
    ' Verifica se é um dos tipos válidos (case insensitive)
    If firstWord = "INDICAÇÃO" Or firstWord = "REQUERIMENTO" Or firstWord = "MOÇÃO" Then
        ' Tipo válido - prossegue
        LogMessage "Documento identificado como: " & firstWord, LOG_LEVEL_INFO
        ValidateDocumentType = True
        Exit Function
    End If
    
    ' Tipo não reconhecido - pergunta ao usuário
    validTypes = "• INDICAÇÃO" & vbCrLf & "• REQUERIMENTO" & vbCrLf & "• MOÇÃO"
    
    userResponse = MsgBox( _
        "O documento parece não ser uma Indicação, Requerimento ou Moção." & vbCrLf & vbCrLf & _
        "Primeira palavra identificada: " & Chr(34) & firstWord & Chr(34) & vbCrLf & vbCrLf & _
        "Tipos válidos esperados:" & vbCrLf & validTypes & vbCrLf & vbCrLf & _
        "Possíveis causas:" & vbCrLf & _
        "• Erro de grafia no título da propositura" & vbCrLf & _
        "• Documento de tipo diferente" & vbCrLf & _
        "• Formatação incorreta do título" & vbCrLf & vbCrLf & _
        "Deseja cancelar ou prosseguir mesmo assim?", _
        vbExclamation + vbYesNo, _
        "Tipo de Documento Não Reconhecido")
    
    If userResponse = vbYes Then
        LogMessage "Usuário optou por prosseguir com documento tipo: " & firstWord, LOG_LEVEL_WARNING
        ValidateDocumentType = True
    Else
        LogMessage "Usuário cancelou processamento - tipo de documento não reconhecido: " & firstWord, LOG_LEVEL_INFO
        ValidateDocumentType = False
    End If
    
    Exit Function

ErrorHandler:
    LogMessage "Erro na validação do tipo de documento: " & Err.Description, LOG_LEVEL_ERROR
    ' Em caso de erro, pergunta ao usuário se quer continuar
    userResponse = MsgBox( _
        "Erro ao validar o tipo de documento:" & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & _
        "Deseja cancelar ou prosseguir?", _
        vbCritical + vbYesNo, _
        "Erro na Validação")
    
    ValidateDocumentType = (userResponse = vbYes)
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
        Dim docFirstWord As String
        wasReplaced = False
        
        ' Obtém a primeira palavra do documento para determinar o tipo
        docFirstWord = GetFirstWordOfDocument(doc)
        
        ' Verifica se inicia com "Solicita" (case insensitive)
        ' CONDICIONAL: Só substitui se a 1ª palavra do documento for "REQUERIMENTO"
        If Len(paraFullText) >= 8 Then
            lowerStart = LCase(Left(paraFullText, 8))
            If lowerStart = "solicita" Then
                If docFirstWord = "REQUERIMENTO" Then
                    para.Range.text = "Requer" & Mid(paraFullText, 9) & vbCr
                    LogMessage "Palavra inicial 'Solicita' substituída por 'Requer' no 2º parágrafo (documento tipo REQUERIMENTO)", LOG_LEVEL_INFO
                    wasReplaced = True
                Else
                    LogMessage "Palavra inicial 'Solicita' não substituída (documento não é REQUERIMENTO, é: " & docFirstWord & ")", LOG_LEVEL_INFO
                End If
            End If
        End If
        
        ' Verifica se inicia com "Pede" (case insensitive)
        ' CONDICIONAL: Só substitui se a 1ª palavra do documento for "REQUERIMENTO"
        If Not wasReplaced And Len(paraFullText) >= 4 Then
            lowerStart = LCase(Left(paraFullText, 4))
            If lowerStart = "pede" Then
                If docFirstWord = "REQUERIMENTO" Then
                    para.Range.text = "Requer" & Mid(paraFullText, 5) & vbCr
                    LogMessage "Palavra inicial 'Pede' substituída por 'Requer' no 2º parágrafo (documento tipo REQUERIMENTO)", LOG_LEVEL_INFO
                    wasReplaced = True
                Else
                    LogMessage "Palavra inicial 'Pede' não substituída (documento não é REQUERIMENTO, é: " & docFirstWord & ")", LOG_LEVEL_INFO
                End If
            End If
        End If
        
        ' Verifica se inicia com "Sugere" (case insensitive)
        ' CONDICIONAL: Só substitui se a 1ª palavra do documento for "INDICAÇÃO"
        If Not wasReplaced And Len(paraFullText) >= 6 Then
            lowerStart = LCase(Left(paraFullText, 6))
            If lowerStart = "sugere" Then
                If docFirstWord = "INDICAÇÃO" Then
                    para.Range.text = "Indica" & Mid(paraFullText, 7) & vbCr
                    LogMessage "Palavra inicial 'Sugere' substituída por 'Indica' no 2º parágrafo (documento tipo INDICAÇÃO)", LOG_LEVEL_INFO
                    wasReplaced = True
                Else
                    LogMessage "Palavra inicial 'Sugere' não substituída (documento não é INDICAÇÃO, é: " & docFirstWord & ")", LOG_LEVEL_INFO
                End If
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
Private Function CountBlankLinesBefore(doc As Document, paraIndex As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim count As Long
    Dim i As Long
    Dim para As Paragraph
    Dim paraText As String
    
    count = 0
    
    ' Verifica parágrafos anteriores (máximo 5 para performance)
    For i = paraIndex - 1 To 1 Step -1
        If i <= 0 Then Exit For
        
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
    
    CountBlankLinesBefore = count
    Exit Function
    
ErrorHandler:
    CountBlankLinesBefore = 0
End Function

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
        
        ' Formata as linhas em branco inseridas ANTES: centralizado e recuos 0
        Dim j As Long
        For j = plenarioIndex To plenarioIndex + 1
            If j <= doc.Paragraphs.count Then
                Set para = doc.Paragraphs(j)
                ' Remove formatação de lista
                On Error Resume Next
                para.Range.ListFormat.RemoveNumbers
                Err.Clear
                On Error GoTo ErrorHandler
                
                With para.Format
                    .leftIndent = 0
                    .firstLineIndent = 0
                    .RightIndent = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .alignment = wdAlignParagraphCenter
                End With
            End If
        Next j
        
        ' Insere EXATAMENTE 2 linhas em branco DEPOIS
        Set para = doc.Paragraphs(plenarioIndex + 2) ' +2 porque inserimos 2 antes
        para.Range.InsertParagraphAfter
        para.Range.InsertParagraphAfter
        
        ' Formata as linhas em branco inseridas DEPOIS: centralizado e recuos 0
        For j = plenarioIndex + 3 To plenarioIndex + 4
            If j <= doc.Paragraphs.count Then
                Set para = doc.Paragraphs(j)
                ' Remove formatação de lista
                On Error Resume Next
                para.Range.ListFormat.RemoveNumbers
                Err.Clear
                On Error GoTo ErrorHandler
                
                With para.Format
                    .leftIndent = 0
                    .firstLineIndent = 0
                    .RightIndent = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .alignment = wdAlignParagraphCenter
                End With
            End If
        Next j
        
        ' FORMATA AS 4 LINHAS TEXTUAIS após as 2 linhas em branco (posições +5, +6, +7, +8)
        ' Salva estado da formatação automática
        Dim autoFormatState As Boolean
        On Error Resume Next
        autoFormatState = Application.Options.AutoFormatAsYouTypeApplyBulletedLists
        Application.Options.AutoFormatAsYouTypeApplyBulletedLists = False
        Err.Clear
        On Error GoTo ErrorHandler
        
        For j = plenarioIndex + 5 To plenarioIndex + 8
            If j <= doc.Paragraphs.count Then
                Set para = doc.Paragraphs(j)
                ' Só formata se NÃO for linha vazia e NÃO tiver conteúdo visual
                paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
                If paraText <> "" And Not HasVisualContent(para) Then
                    ' CRÍTICO: Remove TODAS as formatações de lista (incluindo bullet automático)
                    On Error Resume Next
                    para.Range.ListFormat.RemoveNumbers
                    para.Range.ListFormat.RemoveNumbers ' Força remoção dupla
                    Err.Clear
                    On Error GoTo ErrorHandler
                    
                    ' Seleciona o parágrafo e limpa formatação
                    para.Range.Select
                    
                    ' Zera recuos de forma ABSOLUTA usando pontos (não cm)
                    With para.Format
                        .leftIndent = 0
                        .firstLineIndent = 0
                        .RightIndent = 0
                        .SpaceBefore = 0
                        .SpaceAfter = 0
                    End With
                    
                    ' Define alinhamento DEPOIS de zerar recuos
                    para.Format.alignment = wdAlignParagraphCenter
                    
                    ' FORÇA recuos a zero NOVAMENTE (tripla verificação para parágrafos com "-")
                    para.Format.leftIndent = 0
                    para.Format.firstLineIndent = 0
                    para.Format.RightIndent = 0
                    
                    ' PRIMEIRA linha textual após Plenário: aplica NEGRITO
                    If j = plenarioIndex + 5 Then
                        With para.Range.Font
                            .Bold = True
                            .Name = STANDARD_FONT
                            .size = STANDARD_FONT_SIZE
                        End With
                    End If
                End If
            End If
        Next j
        
        ' Restaura formatação automática
        On Error Resume Next
        Application.Options.AutoFormatAsYouTypeApplyBulletedLists = autoFormatState
        Err.Clear
        On Error GoTo ErrorHandler
        
        LogMessage "Linhas em branco do Plenário reforçadas: 2 antes e 2 depois + 4 linhas textuais (centralizadas, recuos 0, sem lista)", LOG_LEVEL_INFO
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
    Dim nextPara As Paragraph
    Dim paraText As String
    Dim nextParaText As String
    Dim insertionPoint As Range
    Dim addedCount As Long
    
    addedCount = 0
    
    ' Percorre todos os parágrafos de trás para frente para não afetar os índices
    For i = doc.Paragraphs.count - 1 To 1 Step -1
        Set para = doc.Paragraphs(i)
        Set nextPara = doc.Paragraphs(i + 1)
        
        ' Obtém texto limpo dos parágrafos
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        nextParaText = Trim(Replace(Replace(nextPara.Range.text, vbCr, ""), vbLf, ""))
        
        ' Se ambos os parágrafos têm conteúdo (texto ou imagem)
        If (paraText <> "" Or HasVisualContent(para)) And _
           (nextParaText <> "" Or HasVisualContent(nextPara)) Then
            
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
                Set insertionPoint = nextPara.Range
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
    Dim nextPara As Paragraph
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
            Set nextPara = doc.Paragraphs(i + 1)
        Else
            Exit For
        End If
        
        ' Obtém texto limpo dos parágrafos
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        nextParaText = Trim(Replace(Replace(nextPara.Range.text, vbCr, ""), vbLf, ""))
        
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
                Set insertionPoint = nextPara.Range
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
    Dim documentsPath As String
    Dim headerImagePath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")

    ' Obtém pasta Documents do usuário atual (compatível com Windows)
    documentsPath = shell.SpecialFolders("MyDocuments")
    If Right(documentsPath, 1) = "\" Then
        documentsPath = Left(documentsPath, Len(documentsPath) - 1)
    End If

    ' Constrói caminho absoluto para a imagem desejada
    headerImagePath = documentsPath & "\Documentos\CHAINSAW\assets\stamp.png"

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
    imgFile = Environ("USERPROFILE") & "\Documentos\CHAINSAW\assets\stamp.png"

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
    Dim nextPara As Paragraph
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
                Set nextPara = doc.Paragraphs(i + 1)
                Dim nextText As String
                nextText = Trim(Replace(Replace(nextPara.Range.text, vbCr, ""), vbLf, ""))
                
                ' Verifica se a próxima linha está em branco
                If nextText = "" And Not HasVisualContent(nextPara) Then
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
                    Set nextPara = doc.Paragraphs(i)
                    nextText = Trim(Replace(Replace(nextPara.Range.text, vbCr, ""), vbLf, ""))
                    
                    ' Confirma que ainda está vazia antes de deletar
                    If nextText = "" And Not HasVisualContent(nextPara) Then
                        nextPara.Range.Delete
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
' FORMAT "ANTE O EXPOSTO" PARAGRAPHS - Formata "ante o exposto" em caixa alta e negrito
'================================================================================
Private Function FormatAnteOExpostoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim rng As Range
    Dim totalFormatted As Long
    Dim i As Long
    
    ' Percorre todos os parágrafos procurando por "ante o exposto" no início
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Verifica se o parágrafo começa com "ante o exposto" (ignorando maiúsculas/minúsculas)
        If Len(paraText) >= 14 And LCase(Left(paraText, 14)) = "ante o exposto" Then
            ' Verifica se após "ante o exposto" vem espaço, vírgula, ponto-e-vírgula ou fim da linha
            Dim nextChar As String
            If Len(paraText) > 14 Then
                nextChar = Mid(paraText, 15, 1)
                If nextChar = " " Or nextChar = "," Or nextChar = ";" Or nextChar = ":" Or nextChar = "." Then
                    ' É realmente "ante o exposto" no início do parágrafo
                    Set rng = para.Range
                    
                    ' Usa Find/Replace para preservar espaçamento
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
                        
                        ' Limita a busca ao início do parágrafo
                        rng.End = rng.Start + 20  ' Seleciona apenas o início para evitar múltiplas substituições
                        
                        If .Execute(Replace:=True) Then
                            totalFormatted = totalFormatted + 1
                        End If
                    End With
                End If
            Else
                ' Parágrafo contém apenas "ante o exposto"
                Set rng = para.Range
                rng.End = rng.Start + 14
                
                With rng
                    .text = "ANTE O EXPOSTO"
                    .Font.Bold = True
                End With
                
                totalFormatted = totalFormatted + 1
            End If
        End If
    Next i
    
    LogMessage "Formatação 'ante o exposto' aplicada: " & totalFormatted & " ocorrências em negrito e caixa alta", LOG_LEVEL_INFO
    FormatAnteOExpostoParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação 'ante o exposto': " & Err.Description, LOG_LEVEL_ERROR
    FormatAnteOExpostoParagraphs = False
End Function

'================================================================================
' FORMATAÇÃO DE "POR TODAS AS RAZÕES AQUI EXPOSTAS"
'================================================================================
Private Function FormatPorTodasRazoesParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim totalFormatted As Long
    
    totalFormatted = 0
    
    ' Procura parágrafos que começam com "Por todas as razões aqui expostas"
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(para.Range.text)
        
        ' Remove marcador de parágrafo para análise
        If Right(paraText, 1) = vbCr Or Right(paraText, 1) = vbLf Then
            paraText = Left(paraText, Len(paraText) - 1)
            paraText = Trim(paraText)
        End If
        
        ' Verifica se começa com "Por todas as razões aqui expostas" (case insensitive)
        If Len(paraText) >= 35 Then
            Dim firstPart As String
            firstPart = Left(LCase(paraText), 35)
            
            If firstPart = "por todas as razões aqui expostas" Or _
               firstPart = "por todas as razoes aqui expostas" Then
                ' Aplica negrito ao parágrafo inteiro
                With para.Range.Font
                    .Bold = True
                End With
                totalFormatted = totalFormatted + 1
                LogMessage "Negrito aplicado em parágrafo 'Por todas as razões aqui expostas' (parágrafo " & i & ")", LOG_LEVEL_INFO
            End If
        End If
    Next i
    
    If totalFormatted > 0 Then
        LogMessage "Formatação 'Por todas as razões aqui expostas' aplicada: " & totalFormatted & " parágrafo(s) em negrito", LOG_LEVEL_INFO
    End If
    
    FormatPorTodasRazoesParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação 'Por todas as razões aqui expostas': " & Err.Description, LOG_LEVEL_ERROR
    FormatPorTodasRazoesParagraphs = False
End Function

'================================================================================
' APLICAÇÃO DE SUBSTITUIÇÕES DE TEXTO
'================================================================================
Private Function ApplyTextReplacements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ' ========== VALIDAÇÕES INICIAIS ==========
    ' Validação de documento
    If doc Is Nothing Then
        LogMessage "Erro: Documento inválido em ApplyTextReplacements", LOG_LEVEL_ERROR
        ApplyTextReplacements = False
        Exit Function
    End If
    
    ' Validação de acesso ao Range
    On Error Resume Next
    Dim testRange As Range
    Set testRange = doc.Range
    If Err.Number <> 0 Or testRange Is Nothing Then
        On Error GoTo ErrorHandler
        LogMessage "Erro: Não foi possível acessar o Range do documento", LOG_LEVEL_ERROR
        ApplyTextReplacements = False
        Exit Function
    End If
    Set testRange = Nothing
    On Error GoTo ErrorHandler
    
    ' ========== VARIÁVEIS DE CONTROLE ==========
    Dim rng As Range
    Dim totalActualReplacements As Long  ' Conta substituições REAIS, não variantes
    Dim variantProcessedCount As Long     ' Conta variantes processadas
    Dim i As Long
    Dim safetyCounter As Long
    Dim searchText As String
    Dim replacementText As String
    Dim executeResult As Boolean
    
    totalActualReplacements = 0
    variantProcessedCount = 0
    safetyCounter = 0
    
    ' ========== DEFINIÇÃO DE VARIANTES ==========
    ' Funcionalidade: Substitui variantes de "d'Oeste" por formato padronizado
    Dim dOesteVariants() As String
    ReDim dOesteVariants(0 To 13)  ' 14 variantes (0-13)
    
    ' Variantes com diferentes tipos de apóstrofos e capitalizações
    dOesteVariants(0) = "d'O"    ' Apóstrofo padrão (U+0027)
    dOesteVariants(1) = "d´O"    ' Acento agudo (U+00B4)
    dOesteVariants(2) = "d`O"    ' Acento grave (U+0060)
    dOesteVariants(3) = "d'O"    ' Apóstrofo tipográfico direito (U+2019)
    dOesteVariants(4) = "d'o"    ' Minúscula com apóstrofo padrão
    dOesteVariants(5) = "d´o"    ' Minúscula com acento agudo
    dOesteVariants(6) = "d`o"    ' Minúscula com acento grave
    dOesteVariants(7) = "d'o"    ' Minúscula com apóstrofo tipográfico
    dOesteVariants(8) = "D'O"    ' Maiúscula no D com apóstrofo padrão
    dOesteVariants(9) = "D´O"    ' Maiúscula no D com acento agudo
    dOesteVariants(10) = "D`O"   ' Maiúscula no D com acento grave
    dOesteVariants(11) = "D'O"   ' Maiúscula no D com apóstrofo tipográfico
    dOesteVariants(12) = "doO"   ' Sem apóstrofo (erro comum)
    dOesteVariants(13) = "DOO"   ' Tudo maiúsculo sem apóstrofo
    
    ' Texto de substituição padronizado (sempre o mesmo)
    replacementText = "d'Oeste"
    
    ' ========== PROCESSAMENTO DE VARIANTES ==========
    LogMessage "Iniciando substituições de texto: processando " & (UBound(dOesteVariants) + 1) & " variantes", LOG_LEVEL_INFO
    
    For i = LBound(dOesteVariants) To UBound(dOesteVariants)
        ' Proteção contra loops infinitos
        safetyCounter = safetyCounter + 1
        If safetyCounter > 100 Then
            LogMessage "AVISO: Limite de segurança atingido em ApplyTextReplacements", LOG_LEVEL_WARNING
            Exit For
        End If
        
        ' Construção segura do texto de busca
        On Error Resume Next
        searchText = dOesteVariants(i) & "este"
        If Err.Number <> 0 Then
            LogMessage "Erro ao construir texto de busca para variante #" & i, LOG_LEVEL_WARNING
            Err.Clear
            GoTo NextVariant
        End If
        On Error GoTo ErrorHandler
        
        ' Validação do texto de busca
        If Len(searchText) < 5 Or Len(searchText) > 20 Then
            LogMessage "Texto de busca inválido para variante #" & i & ": '" & searchText & "'", LOG_LEVEL_WARNING
            GoTo NextVariant
        End If
        
        ' ===== EXECUÇÃO DA SUBSTITUIÇÃO COM PROTEÇÃO MÁXIMA =====
        On Error Resume Next
        
        ' Cria novo range SEMPRE (nunca reutiliza)
        Set rng = Nothing
        Set rng = doc.Range
        
        ' Validação crítica do range
        If rng Is Nothing Then
            LogMessage "Erro: Range inválido para variante #" & i, LOG_LEVEL_WARNING
            Err.Clear
            GoTo NextVariant
        End If
        
        ' Limpa erro anterior
        Err.Clear
        
        ' Configuração COMPLETA e EXPLÍCITA de todos os parâmetros Find
        With rng.Find
            ' Limpa formatações anteriores
            .ClearFormatting
            .Replacement.ClearFormatting
            
            ' Parâmetros de busca
            .text = searchText
            .Replacement.text = replacementText
            
            ' Direção e escopo
            .Forward = True
            .Wrap = wdFindContinue  ' Continua do início se necessário
            
            ' Opções de formatação
            .Format = False
            .MatchCase = False      ' Case-insensitive (já definido nas variantes)
            .MatchWholeWord = False ' Busca em qualquer parte
            
            ' Opções avançadas (TODAS explícitas para segurança)
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .MatchPrefix = False
            .MatchSuffix = False
            .IgnoreSpace = False
            .IgnorePunct = False
            
            ' Executa a substituição
            executeResult = .Execute(Replace:=wdReplaceAll)
            
            ' Verifica resultado da execução
            If Err.Number <> 0 Then
                LogMessage "Erro ao executar substituição #" & i & " ('" & searchText & "'): " & Err.Description & " (Código: " & Err.Number & ")", LOG_LEVEL_WARNING
                Err.Clear
                executeResult = False
            End If
        End With
        
        ' Contabilização de sucesso
        If executeResult = True Or executeResult = -1 Then
            ' Execute retorna True/-1 se houve pelo menos 1 substituição
            totalActualReplacements = totalActualReplacements + 1
            variantProcessedCount = variantProcessedCount + 1
            
            ' Log detalhado apenas se houver substituições
            If DEBUG_MODE Then
                LogMessage "Variante #" & i & " substituída: '" & searchText & "' -> '" & replacementText & "'", LOG_LEVEL_INFO
            End If
        Else
            ' Não houve substituições (não é erro, apenas não encontrou)
            variantProcessedCount = variantProcessedCount + 1
        End If
        
        ' Limpa o objeto range
        Set rng = Nothing
        
NextVariant:
        On Error GoTo ErrorHandler
        
        ' Permite responsividade da interface
        If i Mod 5 = 0 Then DoEvents
    Next i
    
    ' ========== LOG FINAL ==========
    If totalActualReplacements > 0 Then
        LogMessage "Substituições concluídas: " & totalActualReplacements & " variante(s) com ocorrências substituídas de " & variantProcessedCount & " processadas", LOG_LEVEL_INFO
    Else
        LogMessage "Substituições concluídas: nenhuma ocorrência encontrada em " & variantProcessedCount & " variantes processadas", LOG_LEVEL_INFO
    End If
    
    ' ========== LIMPEZA FINAL ==========
    Set rng = Nothing
    Set testRange = Nothing
    
    ApplyTextReplacements = True
    Exit Function

ErrorHandler:
    ' Log detalhado do erro
    LogMessage "ERRO CRÍTICO em ApplyTextReplacements: " & Err.Description & " (Código: " & Err.Number & ") [Variante: " & i & "/" & UBound(dOesteVariants) & "]", LOG_LEVEL_ERROR
    
    ' Limpeza de recursos mesmo em erro
    On Error Resume Next
    Set rng = Nothing
    Set testRange = Nothing
    On Error GoTo 0
    
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
    Dim nextPara As Paragraph
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
                    Set nextPara = doc.Paragraphs(i + 1)
                    If Not HasVisualContent(nextPara) Then
                        With nextPara.Format
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
        
        ' Formata as linhas em branco inseridas ANTES: centralizado e recuos 0
        For i = plenarioIndex To plenarioIndex + 1
            If i <= doc.Paragraphs.count Then
                Set para = doc.Paragraphs(i)
                ' Remove formatação de lista
                On Error Resume Next
                para.Range.ListFormat.RemoveNumbers
                Err.Clear
                On Error GoTo ErrorHandler
                
                With para.Format
                    .leftIndent = 0
                    .firstLineIndent = 0
                    .RightIndent = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .alignment = wdAlignParagraphCenter
                End With
            End If
        Next i
        
        ' Insere EXATAMENTE 2 linhas em branco DEPOIS
        Set para = doc.Paragraphs(plenarioIndex + 2) ' +2 porque inserimos 2 antes
        para.Range.InsertParagraphAfter
        para.Range.InsertParagraphAfter
        
        ' Formata as linhas em branco inseridas DEPOIS: centralizado e recuos 0
        For i = plenarioIndex + 3 To plenarioIndex + 4
            If i <= doc.Paragraphs.count Then
                Set para = doc.Paragraphs(i)
                ' Remove formatação de lista
                On Error Resume Next
                para.Range.ListFormat.RemoveNumbers
                Err.Clear
                On Error GoTo ErrorHandler
                
                With para.Format
                    .leftIndent = 0
                    .firstLineIndent = 0
                    .RightIndent = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .alignment = wdAlignParagraphCenter
                End With
            End If
        Next i
        
        ' FORMATA AS 3 LINHAS TEXTUAIS após as 2 linhas em branco (posições +5, +6, +7)
        For i = plenarioIndex + 5 To plenarioIndex + 7
            If i <= doc.Paragraphs.count Then
                Set para = doc.Paragraphs(i)
                ' Só formata se NÃO for linha vazia e NÃO tiver conteúdo visual
                paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
                If paraText <> "" And Not HasVisualContent(para) Then
                    ' CRÍTICO: Remove qualquer formatação de lista antes de zerar recuos
                    On Error Resume Next
                    para.Range.ListFormat.RemoveNumbers
                    Err.Clear
                    On Error GoTo ErrorHandler
                    
                    With para.Format
                        .leftIndent = 0
                        .firstLineIndent = 0
                        .RightIndent = 0
                        .SpaceBefore = 0
                        .SpaceAfter = 0
                        .alignment = wdAlignParagraphCenter
                    End With
                    
                    ' PRIMEIRA linha textual após Plenário: aplica NEGRITO
                    If i = plenarioIndex + 5 Then
                        With para.Range.Font
                            .Bold = True
                            .Name = STANDARD_FONT
                            .size = STANDARD_FONT_SIZE
                        End With
                    End If
                End If
            End If
        Next i
        
        LogMessage "2 linhas em branco + 3 linhas textuais formatadas (centralizadas, recuos 0) após 'Plenário Dr. Tancredo Neves'", LOG_LEVEL_INFO
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
' FORMAT REQUEIRO PARAGRAPHS - Formata apenas a palavra "requeiro" no início
'================================================================================
Private Sub FormatRequeiroParagraphs(doc As Document)
    On Error GoTo ErrorHandler
    
    If Not ValidateDocument(doc) Then Exit Sub
    
    Dim para As Paragraph
    Dim paraText As String
    Dim cleanText As String
    Dim formattedCount As Long
    Dim wordRange As Range
    Dim endPos As Long
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
                    ' Determina até onde formatar (palavra "requeiro" + vírgula se houver)
                    endPos = 8 ' Tamanho de "requeiro"
                    If Len(paraText) > 8 And Mid(paraText, 9, 1) = "," Then
                        endPos = 9 ' Inclui a vírgula
                    End If
                    
                    ' Cria range apenas para a palavra "REQUEIRO" (ou "REQUEIRO,")
                    Set wordRange = para.Range
                    wordRange.End = wordRange.Start + endPos
                    
                    ' Aplica formatação APENAS à palavra: negrito e caixa alta
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
        LogMessage "Formatação 'Requeiro': " & formattedCount & " ocorrências formatadas (apenas a palavra em negrito e caixa alta)", LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    LogMessage "Erro ao formatar parágrafos 'Requeiro': " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' SUBROTINA PÚBLICA: ABRIR PASTA DE LOGS E BACKUPS
'================================================================================
Public Sub AbrirPastaLogsEBackups()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim docFolder As String
    Dim backupFolder As String
    Dim fso As Object
    Dim folderToOpen As String
    Dim hasBackups As Boolean
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Tenta obter documento ativo
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler
    
    ' Verifica se há documento ativo salvo
    If doc Is Nothing Or doc.Path = "" Then
        Application.StatusBar = "Abrindo pasta temporária"
        shell "explorer.exe """ & Environ("TEMP") & """", vbNormalFocus
        Exit Sub
    End If
    
    ' Obtém a pasta do documento ativo (logs e backups ficam juntos)
    docFolder = doc.Path
    folderToOpen = docFolder
    
    ' Abre a pasta no Windows Explorer
    Application.StatusBar = "Abrindo pasta do documento"
    shell "explorer.exe """ & folderToOpen & """", vbNormalFocus
    
    ' Log da operação se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Pasta de logs/backups aberta pelo usuário: " & folderToOpen, LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir pasta"
    LogMessage "Erro ao abrir pasta de logs/backups: " & Err.Description, LOG_LEVEL_ERROR
    
    ' Fallback: tenta abrir pasta do documento ou TEMP
    On Error Resume Next
    If Not doc Is Nothing And doc.Path <> "" Then
        shell "explorer.exe """ & doc.Path & """", vbNormalFocus
        Application.StatusBar = "Pasta alternativa aberta"
    Else
        shell "explorer.exe """ & Environ("TEMP") & """", vbNormalFocus
        Application.StatusBar = "Pasta temporária aberta"
    End If
End Sub

'================================================================================
' ABRIR README - COPIA README.MD PARA TEMP E ABRE NO NOTEPAD
'================================================================================
Public Sub AbrirReadme()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim sourceFile As String
    Dim tempFolder As String
    Dim destFile As String
    Dim notepadPath As String
    
    ' Cria objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Define caminhos
    sourceFile = Environ("USERPROFILE") & "\CHAINSAW\README.md"
    tempFolder = Environ("USERPROFILE") & "\AppData\Local\Temp"
    destFile = tempFolder & "\CHAINSAW_README.md"
    notepadPath = Environ("WINDIR") & "\notepad.exe"
    
    ' Verifica se o arquivo de origem existe
    If Not fso.FileExists(sourceFile) Then
        Application.StatusBar = "Erro: README.md não encontrado"
        MsgBox "Arquivo README.md não encontrado em:" & vbCrLf & vbCrLf & _
               sourceFile & vbCrLf & vbCrLf & _
               "Verifique se a instalação foi feita corretamente.", _
               vbExclamation, "Arquivo Não Encontrado"
        
        LogMessage "README.md não encontrado em: " & sourceFile, LOG_LEVEL_ERROR
        Exit Sub
    End If
    
    ' Verifica se a pasta Temp existe (deve sempre existir)
    If Not fso.FolderExists(tempFolder) Then
        Application.StatusBar = "Erro: Pasta Temp não encontrada"
        MsgBox "Pasta temporária não encontrada:" & vbCrLf & vbCrLf & _
               tempFolder & vbCrLf & vbCrLf & _
               "Erro crítico do sistema.", _
               vbCritical, "Erro do Sistema"
        
        LogMessage "Pasta Temp não encontrada: " & tempFolder, LOG_LEVEL_ERROR
        Exit Sub
    End If
    
    ' Remove arquivo de destino se já existir (para garantir cópia atualizada)
    If fso.FileExists(destFile) Then
        On Error Resume Next
        fso.DeleteFile destFile, True
        On Error GoTo ErrorHandler
    End If
    
    ' Copia o arquivo para Temp
    Application.StatusBar = "Copiando README.md..."
    fso.CopyFile sourceFile, destFile, True
    
    ' Verifica se a cópia foi bem-sucedida
    If Not fso.FileExists(destFile) Then
        Application.StatusBar = "Erro ao copiar README.md"
        MsgBox "Não foi possível copiar o arquivo para a pasta temporária." & vbCrLf & vbCrLf & _
               "Destino: " & destFile, _
               vbExclamation, "Erro na Cópia"
        
        LogMessage "Falha ao copiar README.md para: " & destFile, LOG_LEVEL_ERROR
        Exit Sub
    End If
    
    ' Abre o arquivo com Notepad
    Application.StatusBar = "Abrindo README.md no Notepad..."
    shell notepadPath & " """ & destFile & """", vbNormalFocus
    
    Application.StatusBar = "README.md aberto com sucesso"
    
    ' Log da operação
    If loggingEnabled Then
        LogMessage "README.md copiado para Temp e aberto no Notepad: " & destFile, LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir README.md"
    
    Dim errorMsg As String
    errorMsg = "Erro ao abrir o arquivo README.md:" & vbCrLf & vbCrLf & _
               "Erro: " & Err.Description & vbCrLf & _
               "Número: " & Err.Number
    
    MsgBox errorMsg, vbCritical, "Erro"
    
    LogMessage "Erro ao abrir README.md: " & Err.Description & " (Erro #" & Err.Number & ")", LOG_LEVEL_ERROR
    
    ' Limpeza
    On Error Resume Next
    Set fso = Nothing
End Sub

'================================================================================
' SISTEMA DE BACKUP
'================================================================================
Private Function CreateDocumentBackup(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ' Não faz backup se documento não foi salvo
    If doc.Path = "" Then
        LogMessage "Backup ignorado - documento não salvo", LOG_LEVEL_INFO
        CreateDocumentBackup = True
        Exit Function
    End If
    
    Dim backupFolder As String
    Dim fso As Object
    Dim docName As String
    Dim docExtension As String
    Dim timeStamp As String
    Dim backupFileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Define pasta de backup (mesma pasta do documento)
    backupFolder = doc.Path
    
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
    Dim failedCount As Long
    Dim para As Paragraph
    Dim prevPara As Paragraph
    
    restoredCount = 0
    failedCount = 0
    
    ' FASE 1: Restaura as listas em ordem sequencial para manter continuidade
    For i = 0 To listFormatCount - 1
        On Error Resume Next
        
        With savedListFormats(i)
            If .HasList And .paraIndex <= doc.Paragraphs.count Then
                Set para = doc.Paragraphs(.paraIndex)
                
                ' Verifica se o parágrafo ainda existe e tem conteúdo similar
                Dim paraText As String
                paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
                
                ' Só restaura se o parágrafo não estiver vazio
                If Len(paraText) > 0 Then
                    ' Remove qualquer formatação de lista existente primeiro
                    para.Range.ListFormat.RemoveNumbers
                    Err.Clear
                    
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
                            ' Tenta aplicar formatação padrão baseada na string original
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
                    
                    ' FASE 2: Restaura nível de numeração baseado no parágrafo anterior
                    ' (ContinuePreviousList não disponível em todas as versões)
                    If i > 0 And .paraIndex > 1 Then
                        Set prevPara = doc.Paragraphs(.paraIndex - 1)
                        ' Se o parágrafo anterior tem lista do mesmo tipo, mantém a sequência
                        If prevPara.Range.ListFormat.ListType = .ListType Then
                            ' A sequência já continua automaticamente pelo ApplyNumberDefault/ApplyBulletDefault
                            ' Apenas garante que o nível está correto
                            On Error Resume Next
                            If .ListLevelNumber > 0 And .ListLevelNumber <= 9 Then
                                para.Range.ListFormat.ListLevelNumber = .ListLevelNumber
                            End If
                            Err.Clear
                            On Error GoTo ErrorHandler
                        End If
                    End If
                    
                    If Err.Number = 0 Then
                        restoredCount = restoredCount + 1
                    Else
                        failedCount = failedCount + 1
                        LogMessage "Aviso: Falha ao restaurar lista no parágrafo " & .paraIndex & ": " & Err.Description, LOG_LEVEL_WARNING
                        Err.Clear
                    End If
                End If
            End If
        End With
        
        On Error GoTo ErrorHandler
    Next i
    
    If restoredCount > 0 Then
        LogMessage "Formatações de lista restauradas: " & restoredCount & " de " & listFormatCount & " parágrafos", LOG_LEVEL_INFO
    End If
    
    If failedCount > 0 Then
        LogMessage "Aviso: " & failedCount & " formatações de lista não puderam ser restauradas", LOG_LEVEL_WARNING
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
    Dim cleanText As String
    Dim firstChars As String
    Dim formattedCount As Long
    Dim defaultIndent As Single
    
    formattedCount = 0
    
    ' Obtém o recuo padrão de uma lista numerada (aproximadamente 36 pontos ou 1.27 cm)
    ' Esse é o recuo padrão do Word para listas numeradas
    defaultIndent = 36 ' pontos
    
    ' Percorre todos os parágrafos
    For Each para In doc.Paragraphs
        paraText = Trim(para.Range.text)
        cleanText = Trim(Replace(Replace(paraText, vbCr, ""), vbLf, ""))
        
        ' Verifica se o parágrafo não está vazio e tem pelo menos 3 caracteres
        If Len(cleanText) >= 3 Then
            ' Pega os primeiros 3 caracteres para análise mais precisa
            firstChars = Left(cleanText, 3)
            
            ' Verifica se segue o padrão de lista numerada: "N." ou "N)" ou "N-"
            ' onde N é um ou dois dígitos
            Dim isListPattern As Boolean
            isListPattern = False
            
            ' Padrão: 1. ou 10. ou 1) ou 10) ou 1- ou 10-
            If IsNumeric(Left(firstChars, 1)) Then
                Dim secondChar As String
                secondChar = Mid(firstChars, 2, 1)
                
                ' Verifica padrões válidos de lista
                If secondChar = "." Or secondChar = ")" Or secondChar = "-" Or secondChar = " " Then
                    ' Padrão válido de 1 dígito (ex: "1.", "2)", "3-")
                    isListPattern = True
                ElseIf IsNumeric(secondChar) And Len(cleanText) >= 3 Then
                    ' Pode ser 2 dígitos (ex: "10.", "25)")
                    Dim thirdChar As String
                    thirdChar = Mid(cleanText, 3, 1)
                    If thirdChar = "." Or thirdChar = ")" Or thirdChar = "-" Or thirdChar = " " Then
                        isListPattern = True
                    End If
                End If
            End If
            
            ' Só formata se:
            ' 1. Segue o padrão de lista numerada
            ' 2. Não tem formatação de lista já aplicada (para não sobrescrever listas restauradas)
            ' 3. Não tem conteúdo visual (imagens)
            If isListPattern And para.Range.ListFormat.ListType = wdListNoNumbering And Not HasVisualContent(para) Then
                ' Aplica o recuo à esquerda igual ao de uma lista numerada
                With para.Format
                    .leftIndent = defaultIndent
                    .firstLineIndent = 0
                End With
                formattedCount = formattedCount + 1
            End If
        End If
    Next para
    
    If formattedCount > 0 Then
        LogMessage "Parágrafos com padrão de lista numerada formatados com recuo: " & formattedCount, LOG_LEVEL_INFO
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
    
    Dim backupPath As String
    
    ' Define o caminho para backups (mesma pasta do documento)
    If doc.Path <> "" Then
        ' Documento salvo - backups na mesma pasta do documento
        backupPath = doc.Path
    Else
        ' Documento não salvo - usa TEMP como fallback
        backupPath = Environ("TEMP")
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



