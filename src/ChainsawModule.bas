' =============================================================================
' PROJETO: CHAINSAW FOR PROPOSALS (CHAINSW-FPROPS)
' =============================================================================
'
' Sistema automatizado de padronizaÃ§Ã£o de documentos legislativos no Microsoft Word
'
' LicenÃ§a: Apache 2.0 modificada (ver LICENSE)
' VersÃ£o: 1.0-alpha8-optimized | Data: 2025-09-18
' RepositÃ³rio: github.com/chrmsantos/chainsaw-fprops
' Autor: Christian Martin dos Santos <chrmsantos@gmail.com>
'
' =============================================================================
' FUNCIONALIDADES PRINCIPAIS:
' =============================================================================
'
' â€¢ VERIFICAÃ‡Ã•ES DE SEGURANÃ‡A E COMPATIBILIDADE:
'   - ValidaÃ§Ã£o de versÃ£o do Word (mÃ­nimo: 2010)
'   - VerificaÃ§Ã£o de tipo e proteÃ§Ã£o do documento
'   - Controle de espaÃ§o em disco e estrutura mÃ­nima
'   - ProteÃ§Ã£o contra falhas e recuperaÃ§Ã£o automÃ¡tica
'
' â€¢ SISTEMA DE BACKUP AUTOMÃTICO:
'   - Backup automÃ¡tico antes de qualquer modificaÃ§Ã£o
'   - Pasta de backups organizada por documento
'
' CONFIGURAÃ‡Ã•ES PADRÃƒO DE FORMATAÃ‡ÃƒO
Option Explicit

Private executionStartTime As Date
Private formattingCancelled As Boolean
Private undoGroupEnabled As Boolean
Private loggingEnabled As Boolean
Private logFilePath As String

Private Const TOP_MARGIN_CM As Double = 3#
Private Const BOTTOM_MARGIN_CM As Double = 2#
Private Const LEFT_MARGIN_CM As Double = 3#
Private Const RIGHT_MARGIN_CM As Double = 2#
Private Const HEADER_DISTANCE_CM As Double = 1.25
Private Const FOOTER_DISTANCE_CM As Double = 1.25
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 11#
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.22
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 1#
Private Const FOOTER_FONT_SIZE As Long = 10

Private Type ViewSettings
    ViewType As Long
    ShowHorizontalRuler As Boolean
    ShowVerticalRuler As Boolean
    ShowFieldCodes As Boolean
    ShowBookmarks As Boolean
    ShowParagraphMarks As Boolean
    ShowSpaces As Boolean
    ShowTabs As Boolean
    ShowHiddenText As Boolean
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
    ShowOptionalHyphens As Boolean
End Type

Private Const LINE_SPACING As Double = 12# * 1.15# ' 1,15 na interface do Word (12 pt base)
Private Const WORD_HANG_TIMEOUT_SECONDS As Double = 30# ' Tempo limite para detectar travamento iminente
Private Const USER_DATA_ROOT_FOLDER As String = "chainsaw-proposituras"
Private Const LOG_FOLDER_NAME As String = USER_DATA_ROOT_FOLDER & "\logs"
Private Const BACKUP_FOLDER_NAME As String = USER_DATA_ROOT_FOLDER & "\backups"
Private Const MIN_SUPPORTED_VERSION As Double = 14# ' Word 2010
Private Const STANDARD_FONT As String = "Times New Roman"
Private Const STANDARD_FONT_SIZE As Long = 12
Private Const LOG_LEVEL_INFO As Long = 0
Private Const LOG_LEVEL_WARNING As Long = 1
Private Const LOG_LEVEL_ERROR As Long = 2
Private Const PARAGRAPH_BREAK As String = vbCr
Private hangDetectionStart As Double
Private hangDetectionTriggered As Boolean
Private originalViewSettings As ViewSettings
 
Private Function EnsureBlankLinesAroundParagraphIndex(doc As Document, ByRef paraIndex As Long, _
    ByVal requiredBefore As Long, ByVal requiredAfter As Long, _
    Optional ByRef finalBefore As Long, Optional ByRef finalAfter As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim blankBefore As Long
    Dim blankAfter As Long
    Dim addedBefore As Long
    Dim addedAfter As Long
    Dim rng As Range

    If doc Is Nothing Then GoTo ErrorHandler
    If paraIndex < 1 Or paraIndex > doc.Paragraphs.count Then GoTo ErrorHandler

    Set para = doc.Paragraphs(paraIndex)

    If requiredBefore > 0 Then
        blankBefore = CountBlankLinesBefore(doc, paraIndex)
        If blankBefore < requiredBefore Then
            addedBefore = requiredBefore - blankBefore
            Set rng = para.Range
            rng.Collapse wdCollapseStart
            rng.InsertBefore String$(addedBefore, Chr(13))
            paraIndex = paraIndex + addedBefore
            If paraIndex > doc.Paragraphs.count Then paraIndex = doc.Paragraphs.count
            Set para = doc.Paragraphs(paraIndex)
            blankBefore = blankBefore + addedBefore
        End If
    Else
        blankBefore = CountBlankLinesBefore(doc, paraIndex)
    End If

    If requiredAfter > 0 Then
        blankAfter = CountBlankLinesAfter(doc, paraIndex)
        If blankAfter < requiredAfter Then
            addedAfter = requiredAfter - blankAfter
            Set rng = para.Range
            rng.Collapse wdCollapseEnd
            rng.InsertAfter String$(addedAfter, Chr(13))
            blankAfter = blankAfter + addedAfter
        End If
    Else
        blankAfter = CountBlankLinesAfter(doc, paraIndex)
    End If

    finalBefore = blankBefore
    finalAfter = blankAfter
    EnsureBlankLinesAroundParagraphIndex = True
    Exit Function

ErrorHandler:
    EnsureBlankLinesAroundParagraphIndex = False
End Function

Private Function GetNthParagraphIndex(doc As Document, ByVal targetOrder As Long) As Long
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long

    If doc Is Nothing Then GoTo ErrorHandler
    If targetOrder < 1 Then GoTo ErrorHandler

    actualParaIndex = 0

    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim$(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1
            If actualParaIndex = targetOrder Then
                GetNthParagraphIndex = i
                Exit Function
            End If
        End If

        If i > 50 Then Exit For
    Next i

    GetNthParagraphIndex = 0
    Exit Function

ErrorHandler:
    GetNthParagraphIndex = 0
End Function

Private Function GetSecondParagraphIndex(doc As Document) As Long
    GetSecondParagraphIndex = GetNthParagraphIndex(doc, 2)
End Function

Private Function GetThirdParagraphIndex(doc As Document) As Long
    GetThirdParagraphIndex = GetNthParagraphIndex(doc, 3)
End Function

'================================================================================
' FORMAT FIRST PARAGRAPH - FORMATAÃ‡ÃƒO DO 1Âº PARÃGRAFO
'================================================================================
Private Function FormatFirstParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim firstParaIndex As Long
    Dim n As Long
    Dim charCount As Long
    Dim charRange As Range

    actualParaIndex = 0
    firstParaIndex = 0

    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim$(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))

        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1
            If actualParaIndex = 1 Then
                firstParaIndex = i
                Exit For
            End If
        End If

        If i > 20 Then Exit For
    Next i

    If firstParaIndex > 0 And firstParaIndex <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(firstParaIndex)

        If HasVisualContent(para) Then
            charCount = SafeGetCharacterCount(para.Range)
            If charCount > 0 Then
                For n = 1 To charCount
                    Set charRange = para.Range.Characters(n)
                    If charRange.InlineShapes.count = 0 Then
                        With charRange.Font
                            .AllCaps = True
                            .Bold = True
                            .Underline = wdUnderlineSingle
                        End With
                    End If
                Next n
            End If
            LogMessage "1Âº parÃ¡grafo formatado com conteÃºdo visual preservado (posiÃ§Ã£o: " & firstParaIndex & ")"
        Else
            With para.Range.Font
                .AllCaps = True
                .Bold = True
                .Underline = wdUnderlineSingle
            End With
        End If

        With para.Format
            .alignment = wdAlignParagraphCenter
            .leftIndent = 0
            .firstLineIndent = 0
            .RightIndent = 0
        End With
    Else
        LogMessage "1Âº parÃ¡grafo nÃ£o encontrado para formataÃ§Ã£o", LOG_LEVEL_WARNING
    End If

    FormatFirstParagraph = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formataÃ§Ã£o do 1Âº parÃ¡grafo: " & Err.Description, LOG_LEVEL_ERROR
    FormatFirstParagraph = False
End Function
'================================================================================
' MAIN ENTRY POINT
'================================================================================
Public Sub PadronizarDocumentoMain()
    On Error GoTo CriticalErrorHandler
    
    executionStartTime = Now
    formattingCancelled = False
    ResetHangDetection
    
    If Not CheckWordVersion() Then
        Application.StatusBar = "Erro: VersÃ£o do Word nÃ£o suportada (mÃ­nimo: Word 2010)"
        LogMessage "VersÃ£o do Word " & Application.version & " nÃ£o suportada. MÃ­nimo: " & CStr(MIN_SUPPORTED_VERSION), LOG_LEVEL_ERROR
     MsgBox "Esta ferramenta requer Microsoft Word 2010 ou superior." & PARAGRAPH_BREAK & _
         "VersÃ£o atual: " & Application.version & PARAGRAPH_BREAK & _
         "VersÃ£o mÃ­nima: " & CStr(MIN_SUPPORTED_VERSION), vbCritical, "VersÃ£o IncompatÃ­vel"
        Exit Sub
    End If
    
    Dim doc As Document
    Set doc = Nothing
    
    On Error Resume Next
    Set doc = ActiveDocument
    If doc Is Nothing Then
        Application.StatusBar = "Erro: Nenhum documento estÃ¡ acessÃ­vel"
        LogMessage "Nenhum documento acessÃ­vel para processamento", LOG_LEVEL_ERROR
        Exit Sub
    End If
    On Error GoTo CriticalErrorHandler
    
    If Not InitializeLogging(doc) Then
        LogMessage "Falha ao inicializar sistema de logs", LOG_LEVEL_WARNING
    End If
    
    LogMessage "Iniciando padronizaÃ§Ã£o do documento: " & doc.Name, LOG_LEVEL_INFO
    
    StartUndoGroup "PadronizaÃ§Ã£o de Documento - " & doc.Name
    
    If Not SetAppState(False, "Formatando documento...") Then
        LogMessage "Falha ao configurar estado da aplicaÃ§Ã£o", LOG_LEVEL_WARNING
    End If
    
    If Not PreviousChecking(doc) Then
        GoTo CleanUp
    End If
    
    If doc.Path = "" Then
        If Not SaveDocumentFirst(doc) Then
            Application.StatusBar = "OperaÃ§Ã£o cancelada: documento precisa ser salvo"
            LogMessage "OperaÃ§Ã£o cancelada - documento nÃ£o foi salvo", LOG_LEVEL_INFO
            Exit Sub
        End If
    End If
    
    ' Cria backup do documento antes de qualquer modificaÃ§Ã£o
    If Not CreateDocumentBackup(doc) Then
        LogMessage "Falha ao criar backup - continuando sem backup", LOG_LEVEL_WARNING
        Application.StatusBar = "Aviso: Backup nÃ£o foi possÃ­vel - processando sem backup"
    Else
        Application.StatusBar = "Backup criado - formatando documento..."
    End If
    
    ' Backup das configuraÃ§Ãµes de visualizaÃ§Ã£o originais
    If Not BackupViewSettings(doc) Then
        LogMessage "Aviso: Falha no backup das configuraÃ§Ãµes de visualizaÃ§Ã£o", LOG_LEVEL_WARNING
    End If

    If Not PreviousFormatting(doc) Then
        GoTo CleanUp
    End If

    ' Restaura configuraÃ§Ãµes de visualizaÃ§Ã£o originais (exceto zoom)
    If Not RestoreViewSettings(doc) Then
        LogMessage "Aviso: Algumas configuraÃ§Ãµes de visualizaÃ§Ã£o podem nÃ£o ter sido restauradas", LOG_LEVEL_WARNING
    End If

    If formattingCancelled Then
        GoTo CleanUp
    End If

    Application.StatusBar = "Documento padronizado com sucesso!"
    LogMessage "Documento padronizado com sucesso", LOG_LEVEL_INFO

CleanUp:
    SafeCleanup
    CleanupViewSettings    ' Nova funÃ§Ã£o para limpar variÃ¡veis de configuraÃ§Ãµes de visualizaÃ§Ã£o
    
    If Not SetAppState(True, "Documento padronizado com sucesso!") Then
        LogMessage "Falha ao restaurar estado da aplicaÃ§Ã£o", LOG_LEVEL_WARNING
    End If
    
    SafeFinalizeLogging
    
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "ERRO CRÃTICO #" & Err.Number & ": " & Err.Description & _
              " em " & Err.Source & " (Linha: " & Erl & ")"
    
    LogMessage errDesc, LOG_LEVEL_ERROR
    Application.StatusBar = "Erro crÃ­tico durante processamento - verificar logs"
    
    EmergencyRecovery
End Sub

'================================================================================
' EMERGENCY RECOVERY
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
    
    ' Limpa variÃ¡veis de configuraÃ§Ãµes de visualizaÃ§Ã£o em caso de erro
    CleanupViewSettings
    
    LogMessage "RecuperaÃ§Ã£o de emergÃªncia executada", LOG_LEVEL_ERROR
        undoGroupEnabled = False
    
    CloseAllOpenFiles
End Sub

'================================================================================
' SAFE CLEANUP - LIMPEZA SEGURA
'================================================================================
Private Sub SafeCleanup()
    On Error Resume Next
    
    EndUndoGroup
    
    ReleaseObjects
End Sub

'================================================================================
' RELEASE OBJECTS
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
' CLOSE ALL OPEN FILES
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
' VERSION COMPATIBILITY AND SAFETY CHECKS
'================================================================================
Private Function CheckWordVersion() As Boolean
    On Error GoTo ErrorHandler
    
    Dim version As Double
    ' Uso de CDbl para garantir conversÃ£o correta em todas as versÃµes
    version = CDbl(Application.version)
    
    If version < MIN_SUPPORTED_VERSION Then
        CheckWordVersion = False
        LogMessage "VersÃ£o detectada: " & CStr(version) & " - MÃ­nima suportada: " & CStr(MIN_SUPPORTED_VERSION), LOG_LEVEL_ERROR
    Else
        CheckWordVersion = True
        LogMessage "VersÃ£o do Word compatÃ­vel: " & CStr(version), LOG_LEVEL_INFO
    End If
    
    Exit Function
    
ErrorHandler:
    ' Se nÃ£o conseguir detectar a versÃ£o, assume incompatibilidade por seguranÃ§a
    CheckWordVersion = False
    LogMessage "Erro ao detectar versÃ£o do Word: " & Err.Description, LOG_LEVEL_ERROR
End Function

'================================================================================
' SAFE PROPERTY ACCESS FUNCTIONS - Compatibilidade total com Word 2010+
'================================================================================
Private Function SafeGetCharacterCount(targetRange As Range) As Long
    On Error GoTo FallbackMethod
    
    ' MÃ©todo preferido - mais rÃ¡pido
    SafeGetCharacterCount = targetRange.Characters.count
    Exit Function
    
FallbackMethod:
    On Error GoTo ErrorHandler
    ' MÃ©todo alternativo para versÃµes com problemas de .Characters.Count
    SafeGetCharacterCount = Len(targetRange.text)
    Exit Function
    
ErrorHandler:
    ' Ãšltimo recurso - valor padrÃ£o seguro
    SafeGetCharacterCount = 0
    LogMessage "Erro ao obter contagem de caracteres: " & Err.Description, LOG_LEVEL_WARNING
End Function

Private Function SafeSetFont(targetRange As Range, fontName As String, fontSize As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' Aplica formataÃ§Ã£o de fonte de forma segura
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

Private Function SafeHasVisualContent(para As Paragraph) As Boolean
    On Error GoTo SafeMode
    
    ' VerificaÃ§Ã£o padrÃ£o mais robusta
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
    ' MÃ©todo alternativo mais simples
    SafeHasVisualContent = (para.Range.InlineShapes.count > 0)
    Exit Function
    
ErrorHandler:
    ' Em caso de erro, assume que nÃ£o hÃ¡ conteÃºdo visual
    SafeHasVisualContent = False
End Function

'================================================================================
' SAFE CHARACTER ACCESS FUNCTIONS - Compatibilidade total
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
    ' MÃ©todo alternativo usando Right()
    On Error GoTo FinalFallback
    SafeGetLastCharacter = Right(rng.text, 1)
    Exit Function
    
FinalFallback:
    SafeGetLastCharacter = ""
End Function

'================================================================================
' UNDO GROUP MANAGEMENT
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
' LOGGING MANAGEMENT - APRIMORADO COM DETALHES
'================================================================================
Private Function InitializeLogging(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim logsFolder As String
    Dim baseName As String
    Dim timeStamp As String
    Dim fileNo As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    logsFolder = EnsureUserDataDirectory(LOG_FOLDER_NAME)
    If Len(logsFolder) = 0 Then
        logsFolder = Environ("TEMP")
    End If
    
    If Not fso.FolderExists(logsFolder) Then
        fso.CreateFolder logsFolder
    End If
    
    If doc Is Nothing Then
        baseName = "documento"
    ElseIf doc.Name <> "" Then
        baseName = fso.GetBaseName(doc.Name)
    Else
        baseName = "documento"
    End If
    
    baseName = SanitizeFileName(baseName)
    timeStamp = Format(Now, "yyyy-mm-dd_HHmmss")
    logFilePath = fso.BuildPath(logsFolder, timeStamp & "_" & baseName & "_FormattingLog.txt")
    
    fileNo = 0
    fileNo = FreeFile

    Open logFilePath For Output As #fileNo
    Print #fileNo, "========================================================"
    Print #fileNo, "LOG DE FORMATAÃ‡ÃƒO DE DOCUMENTO - SISTEMA DE REGISTRO"
    Print #fileNo, "========================================================"
    Print #fileNo, "DuraÃ§Ã£o: " & Format(Now - executionStartTime, "HH:MM:ss")
    Print #fileNo, "Erros: " & Err.Number & " - " & Err.Description
    Print #fileNo, "Status: INICIANDO"
    Print #fileNo, "--------------------------------------------------------"
    Print #fileNo, "SessÃ£o: " & Format(Now, "yyyy-mm-dd HH:MM:ss")
    Print #fileNo, "UsuÃ¡rio: " & Environ("USERNAME")
    Print #fileNo, "EstaÃ§Ã£o: " & Environ("COMPUTERNAME")
    Print #fileNo, "VersÃ£o Word: " & Application.version
    Print #fileNo, "Documento: " & doc.Name
    Print #fileNo, "Local: " & IIf(doc.Path = "", "(NÃ£o salvo)", doc.Path)
    Print #fileNo, "ProteÃ§Ã£o: " & GetProtectionType(doc)
    Print #fileNo, "Tamanho: " & GetDocumentSize(doc)
    Print #fileNo, "========================================================"
    Close #fileNo
    
    loggingEnabled = True
    InitializeLogging = True
    
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If fileNo <> 0 Then Close #fileNo
    On Error GoTo 0
    loggingEnabled = False
    InitializeLogging = False
End Function

Private Sub LogMessage(message As String, Optional level As Long = LOG_LEVEL_INFO)
    On Error GoTo ErrorHandler
    
    Dim fileNo As Integer
    fileNo = 0

    If Not loggingEnabled Then Exit Sub
    
    Dim levelText As String
    Dim levelIcon As String
    
    Select Case level
        Case LOG_LEVEL_INFO
            levelText = "INFO"
            levelIcon = ""
        Case LOG_LEVEL_WARNING
            levelText = "AVISO"
            levelIcon = ""
        Case LOG_LEVEL_ERROR
            levelText = "ERRO"
            levelIcon = ""
        Case Else
            levelText = "OUTRO"
            levelIcon = ""
    End Select
    
    Dim formattedMessage As String
    formattedMessage = Format(Now, "yyyy-mm-dd HH:MM:ss") & " [" & levelText & "] " & levelIcon & " " & message
    
    fileNo = FreeFile

    Open logFilePath For Append As #fileNo
    Print #fileNo, formattedMessage
    Close #fileNo
    
    Debug.Print "LOG: " & formattedMessage
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If fileNo <> 0 Then Close #fileNo
    On Error GoTo 0
    Debug.Print "FALHA NO LOGGING: " & message
End Sub

Private Sub SafeFinalizeLogging()
    On Error GoTo ErrorHandler
    
    Dim fileNo As Integer
    fileNo = 0

    If loggingEnabled Then
        fileNo = FreeFile

        Open logFilePath For Append As #fileNo
        Print #fileNo, "================================================"
        Print #fileNo, "FIM DA SESSÃƒO - " & Format(Now, "yyyy-mm-dd HH:MM:ss")
        Print #fileNo, "DuraÃ§Ã£o: " & Format(Now - executionStartTime, "HH:MM:ss")
        Print #fileNo, "Erros: " & IIf(Err.Number = 0, "Nenhum", Err.Number & " - " & Err.Description)
        Print #fileNo, "Status: " & IIf(formattingCancelled, "CANCELADO", "CONCLUÃDO")
        Print #fileNo, "================================================"
        Close #fileNo
    End If
    
    loggingEnabled = False
    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If fileNo <> 0 Then Close #fileNo
    On Error GoTo 0
    Debug.Print "Erro ao finalizar logging: " & Err.Description
    loggingEnabled = False
End Sub

'================================================================================
' UTILITY: GET PROTECTION TYPE
'================================================================================
Private Function GetProtectionType(doc As Document) As String
    On Error Resume Next
    
    Select Case doc.protectionType
        Case wdNoProtection: GetProtectionType = "Sem proteÃ§Ã£o"
        Case 1: GetProtectionType = "Protegido contra revisÃµes"
        Case 2: GetProtectionType = "Protegido contra comentÃ¡rios"
        Case 3: GetProtectionType = "Protegido contra formulÃ¡rios"
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
    
    If size < 1024 Then
        GetDocumentSize = size & " bytes"
    ElseIf size < 1048576 Then
        GetDocumentSize = Format(size / 1024, "0.0") & " KB"
    Else
        GetDocumentSize = Format(size / 1048576, "0.0") & " MB"
    End If
End Function

'================================================================================
' APPLICATION STATE HANDLER
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
' GLOBAL CHECKING - VERIFICAÃ‡Ã•ES ROBUSTAS
'================================================================================
Private Function PreviousChecking(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        Application.StatusBar = "Erro: Documento nÃ£o acessÃ­vel para verificaÃ§Ã£o"
        LogMessage "Documento nÃ£o acessÃ­vel para verificaÃ§Ã£o", LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If doc.Type <> wdTypeDocument Then
        Application.StatusBar = "Erro: Tipo de documento nÃ£o suportado (Tipo: " & doc.Type & ")"
        LogMessage "Tipo de documento nÃ£o suportado: " & doc.Type, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If doc.protectionType <> wdNoProtection Then
        Dim protectionType As String
        protectionType = GetProtectionType(doc)
        Application.StatusBar = "Erro: Documento protegido (" & protectionType & ")"
        LogMessage "Documento protegido detectado: " & protectionType, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If
    
    If doc.ReadOnly Then
        Application.StatusBar = "Erro: Documento em modo somente leitura"
        LogMessage "Documento em modo somente leitura: " & doc.FullName, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If Not CheckDiskSpace(doc) Then
        Application.StatusBar = "Erro: EspaÃ§o em disco insuficiente"
        LogMessage "EspaÃ§o em disco insuficiente para operaÃ§Ã£o segura", LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If Not ValidateDocumentStructure(doc) Then
        LogMessage "Estrutura do documento validada com avisos", LOG_LEVEL_WARNING
    End If

    WarnSensitiveData doc

    LogMessage "VerificaÃ§Ãµes de seguranÃ§a concluÃ­das com sucesso", LOG_LEVEL_INFO
    PreviousChecking = True
    Exit Function

ErrorHandler:
    Application.StatusBar = "Erro durante verificaÃ§Ãµes de seguranÃ§a"
    LogMessage "Erro durante verificaÃ§Ãµes: " & Err.Description, LOG_LEVEL_ERROR
    PreviousChecking = False
End Function

'================================================================================
' DISK SPACE CHECK - VERIFICAÃ‡ÃƒO SIMPLIFICADA
'================================================================================
Private Function CheckDiskSpace(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ' VerificaÃ§Ã£o simplificada - assume espaÃ§o suficiente se nÃ£o conseguir verificar
    Dim fso As Object
    Dim drive As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If doc.Path <> "" Then
        Set drive = fso.GetDrive(Left(doc.Path, 3))
    Else
        Set drive = fso.GetDrive(Left(Environ("TEMP"), 3))
    End If
    
    ' VerificaÃ§Ã£o bÃ¡sica - 10MB mÃ­nimo
    If drive.AvailableSpace < 10485760 Then ' 10MB em bytes
        LogMessage "EspaÃ§o em disco muito baixo", LOG_LEVEL_WARNING
        CheckDiskSpace = False
    Else
        CheckDiskSpace = True
    End If
    
    Exit Function
    
ErrorHandler:
    ' Se nÃ£o conseguir verificar, assume que hÃ¡ espaÃ§o suficiente
    CheckDiskSpace = True
End Function

'================================================================================
' WORD HANG DETECTION - MONITORAMENTO DE TRAVAMENTO
'================================================================================
Private Sub ResetHangDetection()
    hangDetectionStart = 0
    hangDetectionTriggered = False
End Sub

Private Function IsWordReady() As Boolean
    On Error GoTo NotReady

    ' Word is considered ready when it is not performing background tasks
    If Application.BackgroundSavingStatus <> 0 Then GoTo NotReady
    If Application.BackgroundPrintingStatus <> 0 Then GoTo NotReady

    IsWordReady = True
    Exit Function

NotReady:
    IsWordReady = False
End Function

Private Function GetTimerElapsedSeconds(ByVal startValue As Double, ByVal currentValue As Double) As Double
    If startValue <= 0 Then
        GetTimerElapsedSeconds = 0
    ElseIf currentValue >= startValue Then
        GetTimerElapsedSeconds = currentValue - startValue
    Else
        GetTimerElapsedSeconds = (86400# - startValue) + currentValue
    End If
End Function

Private Function ShouldAbortForWordHang(ByVal context As String) As Boolean
    On Error GoTo ErrorHandler

    Dim nowTime As Double
    Dim elapsed As Double

    DoEvents
    nowTime = Timer

    If IsWordReady() Then
        hangDetectionStart = 0
        ShouldAbortForWordHang = False
        Exit Function
    End If

    If hangDetectionStart = 0 Then
        hangDetectionStart = nowTime
        ShouldAbortForWordHang = False
        Exit Function
    End If

    elapsed = GetTimerElapsedSeconds(hangDetectionStart, nowTime)

    If elapsed < WORD_HANG_TIMEOUT_SECONDS Then
        ShouldAbortForWordHang = False
        Exit Function
    End If

    If Not hangDetectionTriggered Then
        hangDetectionTriggered = True
        formattingCancelled = True
        Application.StatusBar = "Processo cancelado: possÃ­vel travamento do Word detectado."
        LogMessage "ExecuÃ§Ã£o abortada por seguranÃ§a (contexto: " & context & ") - Word nÃ£o responde hÃ¡ " & Format(elapsed, "0.0") & "s", LOG_LEVEL_ERROR
        MsgBox "A automatizaÃ§Ã£o foi interrompida por seguranÃ§a apÃ³s detectar possÃ­vel travamento do Word. Reabra o documento e tente novamente.", _
               vbCritical, "Processo interrompido"
    End If

    ShouldAbortForWordHang = True
    Exit Function

ErrorHandler:
    ShouldAbortForWordHang = False
End Function

'================================================================================
' SENSITIVE DATA DETECTION - AVISO PARA DADOS PESSOAIS SENSÃVEIS
'================================================================================
Private Sub WarnSensitiveData(doc As Document)
    On Error GoTo ErrorHandler

    Dim docText As String
    Dim lowerText As String
    Dim sensitiveTerms As Variant
    Dim term As Variant
    Dim found As Boolean
    Dim regEx As Object

    If doc Is Nothing Then Exit Sub

    docText = doc.Range.text
    If Len(docText) = 0 Then Exit Sub

    lowerText = LCase$(docText)
    sensitiveTerms = Array("cpf:", "rg:", "cnh:", "filiaÃ§Ã£o", "filiacao", "mÃ£e:", "mae:", "naturalidade:", "estado civil:")

    For Each term In sensitiveTerms
        If InStr(1, lowerText, term, vbBinaryCompare) > 0 Then
            found = True
            Exit For
        End If
    Next term

    If Not found Then
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = False
        regEx.IgnoreCase = True
        regEx.Pattern = "\bpai\s*[:\-]"
        If regEx.Test(docText) Then
            found = True
        End If
    End If

    If found Then
        Application.StatusBar = "Aviso: possÃ­vel presenÃ§a de dados sensÃ­veis. Revise o documento."
        LogMessage "PossÃ­vel presenÃ§a de dados sensÃ­veis detectada no documento", LOG_LEVEL_WARNING
        MsgBox "Aviso: foram encontrados indÃ­cios de dados sensÃ­veis (como CPF, RG, filiaÃ§Ã£o, etc.). Revise o documento antes de prosseguir.", _
               vbExclamation, "VerificaÃ§Ã£o de Dados SensÃ­veis"
    End If

    Exit Sub

ErrorHandler:
    LogMessage "Falha ao verificar dados sensÃ­veis: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' MAIN FORMATTING ROUTINE
'================================================================================
Private Function PreviousFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If Not NormalizeParagraphBreaks(doc) Then
        LogMessage "Falha ao normalizar quebras de parÃ¡grafo", LOG_LEVEL_WARNING
    End If

    ' FormataÃ§Ãµes bÃ¡sicas de pÃ¡gina e estrutura
    If Not ApplyPageSetup(doc) Then
        LogMessage "Falha na configuraÃ§Ã£o de pÃ¡gina", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If

    ' PreparaÃ§Ãµes estruturais antes das formataÃ§Ãµes especÃ­ficas
    CleanDocumentStructure doc
    If formattingCancelled Then GoTo HangAbort
    ValidatePropositionType doc
    If formattingCancelled Then GoTo HangAbort
    FormatDocumentTitle doc
    If formattingCancelled Then GoTo HangAbort
    
    ' FormataÃ§Ãµes principais
    If Not ApplyStdFont(doc) Then
        LogMessage "Falha na formataÃ§Ã£o de fontes", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    
    If Not ApplyStdParagraphs(doc) Then
        LogMessage "Falha na formataÃ§Ã£o de parÃ¡grafos", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If

    ' FormataÃ§Ã£o especÃ­fica do 1Âº parÃ¡grafo (caixa alta, negrito, sublinhado)
    FormatFirstParagraph doc

    ' FormataÃ§Ã£o especÃ­fica do 2Âº parÃ¡grafo
    FormatSecondParagraph doc

    ' FormataÃ§Ãµes especÃ­ficas (sem verificaÃ§Ã£o de retorno para performance)
    FormatConsiderandoParagraphs doc
    If formattingCancelled Then GoTo HangAbort
    ApplyTextReplacements doc
    If formattingCancelled Then GoTo HangAbort
    
    ' FormataÃ§Ã£o especÃ­fica para Justificativa/Anexo/Anexos
    FormatJustificativaAnexoParagraphs doc
    
    EnableHyphenation doc
    If formattingCancelled Then GoTo HangAbort
    RemoveWatermark doc
    If formattingCancelled Then GoTo HangAbort
    InsertHeaderstamp doc
    If formattingCancelled Then GoTo HangAbort
    
    ' Limpeza final de espaÃ§os mÃºltiplos em todo o documento
    CleanMultipleSpaces doc
    If formattingCancelled Then GoTo HangAbort
    
    ' Controle de linhas em branco sequenciais (mÃ¡ximo 2)
    LimitSequentialEmptyLines doc
    If formattingCancelled Then GoTo HangAbort
    
    ' REFORÃ‡O: Garante que o 2Âº parÃ¡grafo mantenha suas 2 linhas em branco
    EnsureSecondParagraphBlankLines doc
    If formattingCancelled Then GoTo HangAbort
    ' REFORÃ‡O: Aplica o mesmo padrÃ£o ao 3Âº parÃ¡grafo
    EnsureThirdParagraphBlankLines doc
    If formattingCancelled Then GoTo HangAbort
    ' REFORÃ‡O: Centraliza controle de espaÃ§amento em parÃ¡grafos "Justificativa"
    EnsureJustificativaBlankLines doc
    If formattingCancelled Then GoTo HangAbort

    ' SubstituiÃ§Ã£o de datas no parÃ¡grafo de plenÃ¡rio
    ReplacePlenarioDateParagraph doc
    If formattingCancelled Then GoTo HangAbort
    
    ' ConfiguraÃ§Ã£o final da visualizaÃ§Ã£o
    ConfigureDocumentView doc
    If formattingCancelled Then GoTo HangAbort
    
    If Not InsertFooterStamp(doc) Then
        LogMessage "Falha na inserÃ§Ã£o do rodapÃ©", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    
    LogMessage "FormataÃ§Ã£o completa aplicada", LOG_LEVEL_INFO
    PreviousFormatting = True
    Exit Function

HangAbort:
    PreviousFormatting = False
    Exit Function

ErrorHandler:
    LogMessage "Erro durante formataÃ§Ã£o: " & Err.Description, LOG_LEVEL_ERROR
    PreviousFormatting = False
End Function

'================================================================================
' PAGE SETUP
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
    
    ' ConfiguraÃ§Ã£o de pÃ¡gina aplicada (sem log detalhado para performance)
    ApplyPageSetup = True
    Exit Function
    
ErrorHandler:
    LogMessage "Erro na configuraÃ§Ã£o de pÃ¡gina: " & Err.Description, LOG_LEVEL_ERROR
    ApplyPageSetup = False
End Function

' ================================================================================
' FONT FORMMATTING
' ================================================================================
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

    For i = doc.Paragraphs.count To 1 Step -1
        Set para = doc.Paragraphs(i)
        If ShouldAbortForWordHang("formataÃ§Ã£o de fontes") Then
            ApplyStdFont = False
            Exit Function
        End If
        hasInlineImage = False
        isTitle = False
        hasConsiderando = False
        needsUnderlineRemoval = False
        needsBoldRemoval = False
        
        ' SUPER OTIMIZADO: VerificaÃ§Ã£o prÃ©via consolidada - uma Ãºnica leitura das propriedades
        Dim paraFont As Font
        Set paraFont = para.Range.Font
        Dim needsFontFormatting As Boolean
        needsFontFormatting = (paraFont.Name <> STANDARD_FONT) Or _
                             (paraFont.size <> STANDARD_FONT_SIZE) Or _
                             (paraFont.Color <> wdColorAutomatic)
        
        ' Cache das verificaÃ§Ãµes de formataÃ§Ã£o especial
        needsUnderlineRemoval = (paraFont.Underline <> wdUnderlineNone)
        needsBoldRemoval = (paraFont.Bold = True)
        
        ' Cache da contagem de InlineShapes para evitar mÃºltiplas chamadas
        Dim inlineShapesCount As Long
        inlineShapesCount = para.Range.InlineShapes.count
        
        ' OTIMIZAÃ‡ÃƒO MÃXIMA: Se nÃ£o precisa de nenhuma formataÃ§Ã£o, pula imediatamente
        If Not needsFontFormatting And Not needsUnderlineRemoval And Not needsBoldRemoval And inlineShapesCount = 0 Then
            formattedCount = formattedCount + 1
            GoTo NextParagraph
        End If

        If inlineShapesCount > 0 Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If
        
        ' OTIMIZADO: VerificaÃ§Ã£o de conteÃºdo visual sÃ³ quando necessÃ¡rio
        If Not hasInlineImage And (needsFontFormatting Or needsUnderlineRemoval Or needsBoldRemoval) Then
            If HasVisualContent(para) Then
                hasInlineImage = True
                skippedCount = skippedCount + 1
            End If
        End If
        
        
        ' OTIMIZADO: VerificaÃ§Ã£o consolidada de tipo de parÃ¡grafo - uma Ãºnica leitura do texto
        Dim paraFullText As String
        Dim isSpecialParagraph As Boolean
        isSpecialParagraph = False
        
        ' SÃ³ faz verificaÃ§Ã£o de texto se for necessÃ¡rio para formataÃ§Ã£o especial
        If needsUnderlineRemoval Or needsBoldRemoval Then
            paraFullText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            
            ' Verifica se Ã© o primeiro parÃ¡grafo com texto (tÃ­tulo) - otimizado
            If i <= 3 And para.Format.alignment = wdAlignParagraphCenter And paraFullText <> "" Then
                isTitle = True
            End If
            
            ' Verifica se o parÃ¡grafo comeÃ§a com "considerando" - otimizado
            If Len(paraFullText) >= 12 And LCase(Left(paraFullText, 12)) = "considerando" Then
                hasConsiderando = True
            End If
            
            ' Verifica se Ã© um parÃ¡grafo especial - otimizado
            Dim cleanParaText As String
            cleanParaText = paraFullText
            ' Remove pontuaÃ§Ã£o final para anÃ¡lise
            Do While Len(cleanParaText) > 0 And (Right(cleanParaText, 1) = "." Or Right(cleanParaText, 1) = "," Or Right(cleanParaText, 1) = ":" Or Right(cleanParaText, 1) = ";")
                cleanParaText = Left(cleanParaText, Len(cleanParaText) - 1)
            Loop
            cleanParaText = Trim(LCase(cleanParaText))
            
            If IsJustificativaHeading(cleanParaText) Or IsVereadorPattern(cleanParaText) Or IsAnexoPattern(cleanParaText) Then
                isSpecialParagraph = True
            End If
            
            ' Verifica se Ã© o parÃ¡grafo ANTERIOR a "- vereador -" (tambÃ©m deve preservar negrito)
            Dim isBeforeVereador As Boolean
            isBeforeVereador = False
            If i < doc.Paragraphs.count Then
                Dim nextPara As Paragraph
                Set nextPara = doc.Paragraphs(i + 1)
                If Not HasVisualContent(nextPara) Then
                    Dim nextParaText As String
                    nextParaText = Trim(Replace(Replace(nextPara.Range.text, vbCr, ""), vbLf, ""))
                    ' Remove pontuaÃ§Ã£o final para anÃ¡lise
                    Dim nextCleanText As String
                    nextCleanText = nextParaText
                    Do While Len(nextCleanText) > 0 And (Right(nextCleanText, 1) = "." Or Right(nextCleanText, 1) = "," Or Right(nextCleanText, 1) = ":" Or Right(nextCleanText, 1) = ";")
                        nextCleanText = Left(nextCleanText, Len(nextCleanText) - 1)
                    Loop
                    nextCleanText = Trim(LCase(nextCleanText))
                    
                    If IsVereadorPattern(nextCleanText) Then
                        isBeforeVereador = True
                    End If
                End If
            End If
        End If

        ' FORMATAÃ‡ÃƒO PRINCIPAL - SÃ³ executa se necessÃ¡rio
        If needsFontFormatting Then
            If Not hasInlineImage Then
                ' FormataÃ§Ã£o rÃ¡pida para parÃ¡grafos sem conteÃºdo visual usando mÃ©todo seguro
                If SafeSetFont(para.Range, STANDARD_FONT, STANDARD_FONT_SIZE) Then
                    formattedCount = formattedCount + 1
                Else
                    ' Fallback para mÃ©todo tradicional em caso de erro
                    With paraFont
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                        .Color = wdColorAutomatic
                    End With
                    formattedCount = formattedCount + 1
                End If
            Else
                ' ParÃ¡grafos com conteÃºdo visual recebem formataÃ§Ã£o caractere a caractere
                Call FormatCharacterByCharacter(para, STANDARD_FONT, STANDARD_FONT_SIZE, wdColorAutomatic, False, False)
                formattedCount = formattedCount + 1
            End If
        End If
        
        ' FORMATAÃ‡ÃƒO ESPECIAL CONSOLIDADA - Remove sublinhado e negrito em uma Ãºnica passada
        If needsUnderlineRemoval Or needsBoldRemoval Then
            ' Determina quais formataÃ§Ãµes remover
            Dim removeUnderline As Boolean
            Dim removeBold As Boolean
            removeUnderline = needsUnderlineRemoval And Not isTitle
            removeBold = needsBoldRemoval And Not isTitle And Not hasConsiderando And Not isSpecialParagraph And Not isBeforeVereador
            
            ' Se precisa remover alguma formataÃ§Ã£o
            If removeUnderline Or removeBold Then
                If Not hasInlineImage Then
                    ' FormataÃ§Ã£o rÃ¡pida para parÃ¡grafos sem conteÃºdo visual
                    If removeUnderline Then paraFont.Underline = wdUnderlineNone
                    If removeBold Then paraFont.Bold = False
                Else
                    ' FormataÃ§Ã£o caractere a caractere para preservar conteÃºdo visual
                    Call FormatCharacterByCharacter(para, "", 0, 0, removeUnderline, removeBold)
                End If
                
                If removeUnderline Then underlineRemovedCount = underlineRemovedCount + 1
            End If
        End If

NextParagraph:
    Next i
    
    ' Log otimizado
    If skippedCount > 0 Then
        LogMessage "Fontes formatadas: " & formattedCount & " parÃ¡grafos (incluindo " & skippedCount & " com conteÃºdo visual preservado)"
    End If
    
    ApplyStdFont = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formataÃ§Ã£o de fonte: " & Err.Description, LOG_LEVEL_ERROR
    ApplyStdFont = False
End Function

'================================================================================
' FORMATAÃ‡ÃƒO CARACTERE POR CARACTERE CONSOLIDADA - #OPTIMIZED
'================================================================================
Private Sub FormatCharacterByCharacter(para As Paragraph, fontName As String, fontSize As Long, fontColor As Long, removeUnderline As Boolean, removeBold As Boolean)
    On Error Resume Next
    
    Dim j As Long
    Dim charCount As Long
    Dim charRange As Range
    
    charCount = SafeGetCharacterCount(para.Range) ' Cache da contagem segura
    
    If charCount > 0 Then ' VerificaÃ§Ã£o de seguranÃ§a
        For j = 1 To charCount
            Set charRange = para.Range.Characters(j)
            If charRange.InlineShapes.count = 0 Then
                With charRange.Font
                    ' Aplica formataÃ§Ã£o de fonte se especificada
                    If fontName <> "" Then .Name = fontName
                    If fontSize > 0 Then .size = fontSize
                    If fontColor >= 0 Then .Color = fontColor
                    
                    ' Remove formataÃ§Ãµes especiais se solicitado
                    If removeUnderline Then .Underline = wdUnderlineNone
                    If removeBold Then .Bold = False
                End With
            End If
        Next j
    End If
End Sub

'================================================================================
' PARAGRAPH FORMATTING
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
    Dim cleanText As String

    rightMarginPoints = 0

    For i = doc.Paragraphs.count To 1 Step -1
        Set para = doc.Paragraphs(i)
        If ShouldAbortForWordHang("formataÃ§Ã£o de parÃ¡grafos") Then
            ApplyStdParagraphs = False
            Exit Function
        End If
        hasInlineImage = False

        If para.Range.InlineShapes.count > 0 Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If
        
        ' ProteÃ§Ã£o adicional: verifica outros tipos de conteÃºdo visual
        If Not hasInlineImage And HasVisualContent(para) Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If

    ' Aplica formataÃ§Ã£o de parÃ¡grafo para TODOS os parÃ¡grafos
    ' (independente de conterem conteÃºdo visual ou nÃ£o)
        
        ' Limpeza robusta de espaÃ§os mÃºltiplos - SEMPRE aplicada
        cleanText = para.Range.text
        
        ' OTIMIZADO: CombinaÃ§Ã£o de mÃºltiplas operaÃ§Ãµes de limpeza em um bloco
        If InStr(cleanText, "  ") > 0 Or InStr(cleanText, vbTab) > 0 Then
            ' Remove mÃºltiplos espaÃ§os consecutivos
            Do While InStr(cleanText, "  ") > 0
                cleanText = Replace(cleanText, "  ", " ")
            Loop
            
            ' Remove espaÃ§os antes/depois de quebras de linha
            cleanText = Replace(cleanText, " " & vbCr, vbCr)
            cleanText = Replace(cleanText, vbCr & " ", vbCr)
            cleanText = Replace(cleanText, " " & vbLf, vbLf)
            cleanText = Replace(cleanText, vbLf & " ", vbLf)
            
            ' Remove tabs extras e converte para espaÃ§os
            Do While InStr(cleanText, vbTab & vbTab) > 0
                cleanText = Replace(cleanText, vbTab & vbTab, vbTab)
            Loop
            cleanText = Replace(cleanText, vbTab, " ")
            
            ' Limpeza final de espaÃ§os mÃºltiplos
            Do While InStr(cleanText, "  ") > 0
                cleanText = Replace(cleanText, "  ", " ")
            Loop
        End If
        
    ' Aplica o texto limpo apenas se nÃ£o hÃ¡ conteÃºdo visual
        If cleanText <> para.Range.text And Not hasInlineImage Then
            para.Range.text = cleanText
        End If

        'paraText = Trim(LCase(Replace(Replace(para.Range.text, ".", ""), ",", ""), ";", ""))
        paraText = Trim(LCase(Replace(Replace(para.Range.text, vbCr, ""), vbLf, "")))
        paraText = Replace(paraText, vbTab, "")
        ' FormataÃ§Ã£o de parÃ¡grafo - SEMPRE aplicada
        With para.Format
            .LineSpacingRule = wdLineSpaceMultiple
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
        
        formattedCount = formattedCount + 1
    Next i
    
    ' Log atualizado para refletir que todos os parÃ¡grafos sÃ£o formatados
    If skippedCount > 0 Then
        LogMessage "ParÃ¡grafos formatados: " & formattedCount & " (incluindo " & skippedCount & " com conteÃºdo visual preservado)"
    End If
    
    ApplyStdParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formataÃ§Ã£o de parÃ¡grafos: " & Err.Description, LOG_LEVEL_ERROR
    ApplyStdParagraphs = False
End Function

'================================================================================
' FORMAT SECOND PARAGRAPH - FORMATAÃ‡ÃƒO APENAS DO 2Âº PARÃGRAFO
'================================================================================
Private Function FormatSecondParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim secondParaIndex As Long
    
    ' Identifica apenas o 2Âº parÃ¡grafo (considerando apenas parÃ¡grafos com texto)
    actualParaIndex = 0
    secondParaIndex = 0
    
    ' Encontra o 2Âº parÃ¡grafo com conteÃºdo (pula vazios)
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        If ShouldAbortForWordHang("detecÃ§Ã£o do 2Âº parÃ¡grafo") Then
            FormatSecondParagraph = False
            Exit Function
        End If
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Se o parÃ¡grafo tem texto ou conteÃºdo visual, conta como parÃ¡grafo vÃ¡lido
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1
            
            ' Registra o Ã­ndice do 2Âº parÃ¡grafo
            If actualParaIndex = 2 Then
                secondParaIndex = i
                Exit For ' JÃ¡ encontramos o 2Âº parÃ¡grafo
            End If
        End If
        
        ' ProteÃ§Ã£o expandida: processa atÃ© 20 parÃ¡grafos para encontrar o 2Âº
        If i > 20 Then Exit For
    Next i
    
    ' Aplica formataÃ§Ã£o especÃ­fica apenas ao 2Âº parÃ¡grafo
    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(secondParaIndex)
        
        ' PRIMEIRO: Adiciona 2 linhas em branco ANTES do 2Âº parÃ¡grafo
        Dim insertionPoint As Range
        Set insertionPoint = para.Range
        insertionPoint.Collapse wdCollapseStart
        
        ' Verifica se jÃ¡ existem linhas em branco antes
        Dim blankLinesBefore As Long
        blankLinesBefore = CountBlankLinesBefore(doc, secondParaIndex)
        
        ' Adiciona linhas em branco conforme necessÃ¡rio para chegar a 2
        If blankLinesBefore < 2 Then
            Dim linesToAdd As Long
            linesToAdd = 2 - blankLinesBefore
            
            Dim newLines As String
            newLines = String$(linesToAdd, PARAGRAPH_BREAK)
            insertionPoint.InsertBefore newLines
            
            ' Atualiza o Ã­ndice do segundo parÃ¡grafo (foi deslocado)
            secondParaIndex = secondParaIndex + linesToAdd
            Set para = doc.Paragraphs(secondParaIndex)
        End If
        
        ' FORMATAÃ‡ÃƒO PRINCIPAL: Aplica formataÃ§Ã£o SEMPRE, preservando conteÃºdo visual quando presente
        With para.Format
            .leftIndent = CentimetersToPoints(9)      ' Recuo Ã  esquerda de 9 cm
            .firstLineIndent = 0                      ' Sem recuo da primeira linha
            .RightIndent = 0                          ' Sem recuo Ã  direita
            .alignment = wdAlignParagraphJustify      ' Justificado
        End With
        
        ' SEGUNDO: Adiciona 2 linhas em branco DEPOIS do 2Âº parÃ¡grafo
        Dim insertionPointAfter As Range
        Set insertionPointAfter = para.Range
        insertionPointAfter.Collapse wdCollapseEnd
        
        ' Verifica se jÃ¡ existem linhas em branco depois
        Dim blankLinesAfter As Long
        blankLinesAfter = CountBlankLinesAfter(doc, secondParaIndex)
        
        ' Adiciona linhas em branco conforme necessÃ¡rio para chegar a 2
        If blankLinesAfter < 2 Then
            Dim linesToAddAfter As Long
            linesToAddAfter = 2 - blankLinesAfter
            
            Dim newLinesAfter As String
            newLinesAfter = String$(linesToAddAfter, PARAGRAPH_BREAK)
            insertionPointAfter.InsertAfter newLinesAfter
        End If
        
    ' Se hÃ¡ conteÃºdo visual, apenas registra (mas nÃ£o pula a formataÃ§Ã£o)
        If HasVisualContent(para) Then
            LogMessage "2Âº parÃ¡grafo formatado com conteÃºdo visual preservado e linhas em branco (posiÃ§Ã£o: " & secondParaIndex & ")", LOG_LEVEL_INFO
        Else
            LogMessage "2Âº parÃ¡grafo formatado com 2 linhas em branco antes e depois (posiÃ§Ã£o: " & secondParaIndex & ")", LOG_LEVEL_INFO
        End If
    Else
        LogMessage "2Âº parÃ¡grafo nÃ£o encontrado para formataÃ§Ã£o", LOG_LEVEL_WARNING
    End If
    
    FormatSecondParagraph = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formataÃ§Ã£o do 2Âº parÃ¡grafo: " & Err.Description, LOG_LEVEL_ERROR
    FormatSecondParagraph = False
End Function

'================================================================================
' HELPER FUNCTIONS FOR BLANK LINES - FunÃ§Ãµes auxiliares para linhas em branco
'================================================================================
Private Function CountBlankLinesBefore(doc As Document, paraIndex As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim count As Long
    Dim i As Long
    Dim para As Paragraph
    Dim paraText As String
    
    count = 0
    
    ' Verifica parÃ¡grafos anteriores (mÃ¡ximo 5 para performance)
    For i = paraIndex - 1 To 1 Step -1
        If i <= 0 Then Exit For
        
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Se o parÃ¡grafo estÃ¡ vazio, conta como linha em branco
        If paraText = "" And Not HasVisualContent(para) Then
            count = count + 1
        Else
            ' Se encontrou parÃ¡grafo com conteÃºdo, para de contar
            Exit For
        End If
        
        ' Limite de seguranÃ§a
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
    
    ' Verifica parÃ¡grafos posteriores (mÃ¡ximo 5 para performance)
    For i = paraIndex + 1 To doc.Paragraphs.count
        If i > doc.Paragraphs.count Then Exit For
        
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Se o parÃ¡grafo estÃ¡ vazio, conta como linha em branco
        If paraText = "" And Not HasVisualContent(para) Then
            count = count + 1
        Else
            ' Se encontrou parÃ¡grafo com conteÃºdo, para de contar
            Exit For
        End If
        
        ' Limite de seguranÃ§a
        If count >= 5 Then Exit For
    Next i
    
    CountBlankLinesAfter = count
    Exit Function
    
ErrorHandler:
    CountBlankLinesAfter = 0
End Function

'================================================================================
' SECOND PARAGRAPH BLANK LINES - ReforÃ§a linhas em branco do 2Âº parÃ¡grafo
'================================================================================
Private Function EnsureSecondParagraphBlankLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim secondParaIndex As Long
    Dim beforeResult As Long
    Dim afterResult As Long

    If doc Is Nothing Then
        EnsureSecondParagraphBlankLines = True
        Exit Function
    End If

    secondParaIndex = GetSecondParagraphIndex(doc)

    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.count Then
        If EnsureBlankLinesAroundParagraphIndex(doc, secondParaIndex, 2, 2, beforeResult, afterResult) Then
            LogMessage "Linhas em branco do 2Âº parÃ¡grafo reforÃ§adas (antes: " & beforeResult & ", depois: " & afterResult & ")", LOG_LEVEL_INFO
        End If
    End If

    EnsureSecondParagraphBlankLines = True
    Exit Function

ErrorHandler:
    EnsureSecondParagraphBlankLines = False
    LogMessage "Erro ao garantir linhas em branco do 2Âº parÃ¡grafo: " & Err.Description, LOG_LEVEL_WARNING
End Function

Private Function EnsureThirdParagraphBlankLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim thirdParaIndex As Long
    Dim beforeResult As Long
    Dim afterResult As Long

    If doc Is Nothing Then
        EnsureThirdParagraphBlankLines = True
        Exit Function
    End If

    thirdParaIndex = GetThirdParagraphIndex(doc)

    If thirdParaIndex > 0 And thirdParaIndex <= doc.Paragraphs.count Then
        If EnsureBlankLinesAroundParagraphIndex(doc, thirdParaIndex, 2, 2, beforeResult, afterResult) Then
            LogMessage "Linhas em branco do 3Âº parÃ¡grafo reforÃ§adas (antes: " & beforeResult & ", depois: " & afterResult & ")", LOG_LEVEL_INFO
        End If
    End If

    EnsureThirdParagraphBlankLines = True
    Exit Function

ErrorHandler:
    EnsureThirdParagraphBlankLines = False
    LogMessage "Erro ao garantir linhas em branco do 3Âº parÃ¡grafo: " & Err.Description, LOG_LEVEL_WARNING
End Function

Private Function EnsureJustificativaBlankLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim para As Paragraph
    Dim paraText As String
    Dim normalized As String
    Dim paraIndex As Long
    Dim adjustedCount As Long

    If doc Is Nothing Then
        EnsureJustificativaBlankLines = True
        Exit Function
    End If

    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        If ShouldAbortForWordHang("ajuste de justificativa") Then
            EnsureJustificativaBlankLines = False
            Exit Function
        End If
        paraText = Replace(Replace(para.Range.text, vbCr, ""), vbLf, "")
        paraText = Trim$(paraText)
        If Len(paraText) = 0 Then GoTo ContinueLoop

        normalized = NormalizeHeadingKey(paraText)
        
        If normalized = "justificativa" Then
            paraIndex = i
            If EnsureBlankLinesAroundParagraphIndex(doc, paraIndex, 2, 2) Then
                adjustedCount = adjustedCount + 1
                i = paraIndex
            End If
        End If

ContinueLoop:
    Next i

    If adjustedCount > 0 Then
        LogMessage "Linhas em branco reforÃ§adas em " & adjustedCount & " parÃ¡grafo(s) 'Justificativa'", LOG_LEVEL_INFO
    End If

    EnsureJustificativaBlankLines = True
    Exit Function

ErrorHandler:
    EnsureJustificativaBlankLines = False
    LogMessage "Erro ao reforÃ§ar linhas em branco de 'Justificativa': " & Err.Description, LOG_LEVEL_WARNING
End Function


'================================================================================
' ENABLE HYPHENATION
'================================================================================
Private Function EnableHyphenation(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    If Not doc.AutoHyphenation Then
        doc.AutoHyphenation = True
        doc.HyphenationZone = CentimetersToPoints(0.63)
        doc.HyphenateCaps = True
        ' Log removido para performance
        EnableHyphenation = True
    Else
        ' Log removido para performance
        EnableHyphenation = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "Erro ao ativar hifenizaÃ§Ã£o: " & Err.Description, LOG_LEVEL_ERROR
    EnableHyphenation = False
End Function

'================================================================================
' REMOVE WATERMARK
'================================================================================
Private Function RemoveWatermark(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As section
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
        LogMessage "Marcas d'Ã¡gua removidas: " & removedCount & " itens"
    End If
    ' Log de "nenhuma marca d'Ã¡gua" removido para performance
    
    RemoveWatermark = True
    Exit Function
    
ErrorHandler:
    LogMessage "Erro ao remover marcas d'Ã¡gua: " & Err.Description, LOG_LEVEL_ERROR
    RemoveWatermark = False
End Function

'================================================================================
' USER DATA PATH HELPERS
'================================================================================
Private Function GetUserDocumentsPath() As String
    On Error GoTo ErrorHandler
    
    Dim shell As Object
    Dim documentsPath As String
    
    Set shell = CreateObject("WScript.Shell")
    documentsPath = shell.SpecialFolders("MyDocuments")
    
    If Right(documentsPath, 1) = "\" Then
        documentsPath = Left(documentsPath, Len(documentsPath) - 1)
    End If
    
    GetUserDocumentsPath = documentsPath
    Exit Function
    
ErrorHandler:
    GetUserDocumentsPath = Environ("USERPROFILE") & "\Documents"
End Function

Private Function EnsureUserDataDirectory(relativePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim basePath As String
    Dim currentPath As String
    Dim pathParts As Variant
    Dim part As Variant
    
    basePath = GetUserDocumentsPath()
    If Len(basePath) = 0 Then GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    currentPath = basePath
    
    If Len(relativePath) > 0 Then
        pathParts = Split(relativePath, "\")
        For Each part In pathParts
            If Len(Trim$(CStr(part))) > 0 Then
                currentPath = fso.BuildPath(currentPath, CStr(part))
                If Not fso.FolderExists(currentPath) Then
                    fso.CreateFolder currentPath
                End If
            End If
        Next part
    End If
    
    EnsureUserDataDirectory = currentPath
    Exit Function
    
ErrorHandler:
    EnsureUserDataDirectory = ""
End Function

Private Function SanitizeFileName(ByVal rawName As String) As String
    Dim invalidChars As Variant
    Dim ch As Variant
    
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    
    For Each ch In invalidChars
        rawName = Replace(rawName, CStr(ch), "_")
    Next ch
    
    SanitizeFileName = rawName
End Function

'================================================================================
' INSERT HEADER IMAGE
'================================================================================
Private Function InsertHeaderstamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As section
    Dim header As HeaderFooter
    Dim imgFile As String
    Dim username As String
    Dim imgWidth As Single
    Dim imgHeight As Single
    Dim shp As shape
    Dim imgFound As Boolean
    Dim sectionsProcessed As Long

    ' Resolve image file path using configuration if provided
    'imgFile = Trim(config.headerImagePath)
    If Len(imgFile) = 0 Then
        ' No configured path; try common locations relative to document or repo folder
        Dim baseFolder As String
        On Error Resume Next
        baseFolder = IIf(doc.Path <> "", doc.Path, Environ("USERPROFILE") & "\Documents")
        On Error GoTo ErrorHandler
        If Right(baseFolder, 1) <> "\" Then baseFolder = baseFolder & "\"
        If Dir(baseFolder & "assets\stamp.png") <> "" Then
            imgFile = baseFolder & "assets\stamp.png"
        ElseIf Dir(Environ("USERPROFILE") & "\Documents\chainsaw\assets\stamp.png") <> "" Then
            imgFile = Environ("USERPROFILE") & "\Documents\chainsaw\assets\stamp.png"
        End If
    Else
        ' If relative path (no drive letter), try resolve from current document folder and repo root
        If InStr(1, imgFile, ":", vbTextCompare) = 0 And Left(imgFile, 2) <> "\\" Then
            On Error Resume Next
            baseFolder = IIf(doc.Path <> "", doc.Path, Environ("USERPROFILE") & "\Documents")
            On Error GoTo ErrorHandler
            If Right(baseFolder, 1) <> "\" Then baseFolder = baseFolder & "\"
            If Dir(baseFolder & imgFile) <> "" Then
                imgFile = baseFolder & imgFile
            ElseIf Dir(Environ("USERPROFILE") & "\Documents\chainsaw\" & imgFile) <> "" Then
                imgFile = Environ("USERPROFILE") & "\Documents\chainsaw\" & imgFile
            End If
        End If
    End If

    If Dir(imgFile) = "" Then
        Application.StatusBar = "Warning: Header image not found"
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
            
            Set shp = header.Shapes.AddPicture( _
                FileName:=imgFile, _
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
' INSERT FOOTER PAGE NUMBERS
'================================================================================
Private Function InsertFooterStamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As section
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
    LogMessage "Erro ao inserir rodapÃ©: " & Err.Description, LOG_LEVEL_ERROR
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
' VALIDATE DOCUMENT STRUCTURE - SIMPLIFICADO
'================================================================================
Private Function ValidateDocumentStructure(doc As Document) As Boolean
    On Error Resume Next
    
    ' VerificaÃ§Ã£o bÃ¡sica e rÃ¡pida
    If doc.Range.End > 0 And doc.Sections.count > 0 Then
        ValidateDocumentStructure = True
    Else
        LogMessage "Documento com estrutura inconsistente", LOG_LEVEL_WARNING
        ValidateDocumentStructure = False
    End If
End Function

'================================================================================
' CRITICAL FIX: SAVE DOCUMENT BEFORE PROCESSING
' TO PREVENT CRASHES ON NEW NON SAVED DOCUMENTS
'================================================================================
Private Function SaveDocumentFirst(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Aguardando salvamento do documento..."
    ' Log de inÃ­cio removido para performance
    
    Dim saveDialog As Object
    Set saveDialog = Application.Dialogs(wdDialogFileSaveAs)

    If saveDialog.Show <> -1 Then
        LogMessage "OperaÃ§Ã£o de salvamento cancelada pelo usuÃ¡rio", LOG_LEVEL_INFO
        Application.StatusBar = "Salvamento cancelado pelo usuÃ¡rio"
        SaveDocumentFirst = False
        Exit Function
    End If

    ' Aguarda confirmaÃ§Ã£o do salvamento com timeout de seguranÃ§a
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
        Application.StatusBar = "Aguardando salvamento... (" & waitCount & "/" & maxWait & ")"
    Next waitCount

    If doc.Path = "" Then
        LogMessage "Falha ao salvar documento apÃ³s " & maxWait & " tentativas", LOG_LEVEL_ERROR
        Application.StatusBar = "Falha no salvamento - operaÃ§Ã£o cancelada"
        SaveDocumentFirst = False
    Else
        ' Log de sucesso removido para performance
        Application.StatusBar = "Documento salvo com sucesso"
        SaveDocumentFirst = True
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro durante salvamento: " & Err.Description & " (Erro #" & Err.Number & ")", LOG_LEVEL_ERROR
    Application.StatusBar = "Erro durante salvamento"
    SaveDocumentFirst = False
End Function

'================================================================================
' NORMALIZE PARAGRAPH BREAKS
'================================================================================
Private Function NormalizeParagraphBreaks(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim rng As Range
    Dim replacedPairs As Long
    Dim replacedLineFeeds As Long
    Dim guardCounter As Long
    Dim docText As String
    Dim hasLineFeeds As Boolean
    Dim hasCarriagePairs As Boolean

    If doc Is Nothing Then
        NormalizeParagraphBreaks = True
        Exit Function
    End If

    docText = doc.Content.text
    hasLineFeeds = (InStr(docText, vbLf) > 0)
    hasCarriagePairs = (InStr(docText, vbCrLf) > 0)

    If Not hasLineFeeds Then
        NormalizeParagraphBreaks = True
        Exit Function
    End If

    Application.StatusBar = "Normalizando quebras de parÃ¡grafo..."

    If hasCarriagePairs Then
        Set rng = doc.Content
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = vbCrLf
            .Replacement.text = PARAGRAPH_BREAK
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False

            Do While .Execute(Replace:=True)
                replacedPairs = replacedPairs + 1
                guardCounter = guardCounter + 1
                If guardCounter Mod 100 = 0 Then
                    DoEvents
                    If ShouldAbortForWordHang("normalizaÃ§Ã£o de quebras") Then
                        NormalizeParagraphBreaks = False
                        Exit Function
                    End If
                End If
                If guardCounter > 50000 Then Exit Do
            Loop
        End With
    End If

    If InStr(doc.Content.text, vbLf) > 0 Then
        Set rng = doc.Content
        guardCounter = 0
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = vbLf
            .Replacement.text = PARAGRAPH_BREAK
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False

            Do While .Execute(Replace:=True)
                replacedLineFeeds = replacedLineFeeds + 1
                guardCounter = guardCounter + 1
                If guardCounter Mod 100 = 0 Then
                    DoEvents
                    If ShouldAbortForWordHang("normalizaÃ§Ã£o de quebras") Then
                        NormalizeParagraphBreaks = False
                        Exit Function
                    End If
                End If
                If guardCounter > 50000 Then Exit Do
            Loop
        End With
    End If

    If replacedPairs > 0 Or replacedLineFeeds > 0 Then
        LogMessage "Quebras normalizadas: " & replacedPairs & " CR+LF convertidos e " & replacedLineFeeds & " LF isolados tratados", LOG_LEVEL_INFO
    End If

    NormalizeParagraphBreaks = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao normalizar quebras de parÃ¡grafo: " & Err.Description, LOG_LEVEL_WARNING
    NormalizeParagraphBreaks = False
End Function

'================================================================================
' CLEAN DOCUMENT STRUCTURE - FUNCIONALIDADES 2, 6, 7
'================================================================================
Private Function CleanDocumentStructure(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim i As Long
    Dim firstTextParaIndex As Long
    Dim emptyLinesRemoved As Long
    Dim leadingSpacesRemoved As Long
    Dim paraCount As Long
    
    ' Cache da contagem total de parÃ¡grafos
    paraCount = doc.Paragraphs.count
    
    ' OTIMIZADO: Funcionalidade 2 - Remove linhas em branco acima do tÃ­tulo
    ' Busca otimizada do primeiro parÃ¡grafo com texto
    firstTextParaIndex = -1
    For i = 1 To paraCount
        If i > doc.Paragraphs.count Then Exit For ' ProteÃ§Ã£o dinÃ¢mica
        
        Set para = doc.Paragraphs(i)
        If ShouldAbortForWordHang("limpeza de estrutura") Then
            CleanDocumentStructure = False
            Exit Function
        End If
        Dim paraTextCheck As String
        paraTextCheck = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Encontra o primeiro parÃ¡grafo com texto real
        If paraTextCheck <> "" Then
            firstTextParaIndex = i
            Exit For
        End If
        
        ' ProteÃ§Ã£o contra documentos muito grandes
        If i > 50 Then Exit For ' Limita busca aos primeiros 50 parÃ¡grafos
    Next i
    
    ' OTIMIZADO: Remove linhas vazias ANTES do primeiro texto em uma Ãºnica passada
    If firstTextParaIndex > 1 Then
        ' Processa de trÃ¡s para frente para evitar problemas com Ã­ndices
        For i = firstTextParaIndex - 1 To 1 Step -1
            If i > doc.Paragraphs.count Or i < 1 Then Exit For ' ProteÃ§Ã£o dinÃ¢mica
            
            Set para = doc.Paragraphs(i)
            If ShouldAbortForWordHang("remoÃ§Ã£o de linhas em branco iniciais") Then
                CleanDocumentStructure = False
                Exit Function
            End If
            Dim paraTextEmpty As String
            paraTextEmpty = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            
            ' OTIMIZADO: VerificaÃ§Ã£o visual sÃ³ se necessÃ¡rio
            If paraTextEmpty = "" Then
                If Not HasVisualContent(para) Then
                    para.Range.Delete
                    emptyLinesRemoved = emptyLinesRemoved + 1
                    ' Atualiza cache apÃ³s remoÃ§Ã£o
                    paraCount = paraCount - 1
                End If
            End If
        Next i
    End If
    
    ' SUPER OTIMIZADO: Funcionalidade 7 - Remove espaÃ§os iniciais com regex
    ' Usa Find/Replace que Ã© muito mais rÃ¡pido que loop por parÃ¡grafo
    Dim rng As Range
    Set rng = doc.Range
    
    ' Remove espaÃ§os no inÃ­cio de linhas usando Find/Replace
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        
        ' Remove espaÃ§os/tabs no inÃ­cio de linhas usando Find/Replace simples
        .text = "^p "  ' Quebra seguida de espaÃ§o
        .Replacement.text = "^p"
        
        Do While .Execute(Replace:=True)
            leadingSpacesRemoved = leadingSpacesRemoved + 1
            ' ProteÃ§Ã£o contra loop infinito
            If leadingSpacesRemoved > 1000 Then Exit Do
        Loop
        
        ' Remove tabs no inÃ­cio de linhas
        .text = "^p^t"  ' Quebra seguida de tab
        .Replacement.text = "^p"
        
        Do While .Execute(Replace:=True)
            leadingSpacesRemoved = leadingSpacesRemoved + 1
            If leadingSpacesRemoved > 1000 Then Exit Do
        Loop
    End With
    
    ' Segunda passada para espaÃ§os no inÃ­cio do documento (sem ^p precedente)
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False  ' NÃ£o usa wildcards nesta seÃ§Ã£o
        
        ' Posiciona no inÃ­cio do documento
        rng.Start = 0
        rng.End = 1
        
        ' Remove espaÃ§os/tabs no inÃ­cio absoluto do documento
        If rng.text = " " Or rng.text = vbTab Then
            ' Expande o range para pegar todos os espaÃ§os iniciais usando mÃ©todo seguro
            Do While rng.End <= doc.Range.End And (SafeGetLastCharacter(rng) = " " Or SafeGetLastCharacter(rng) = vbTab)
                rng.End = rng.End + 1
                leadingSpacesRemoved = leadingSpacesRemoved + 1
                If leadingSpacesRemoved > 100 Then Exit Do ' ProteÃ§Ã£o
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
' SAFE CHECK FOR VISUAL CONTENT - VERIFICAÃ‡ÃƒO SEGURA DE CONTEÃšDO VISUAL
'================================================================================
Private Function HasVisualContent(para As Paragraph) As Boolean
    ' Usa a funÃ§Ã£o segura implementada para compatibilidade total
    HasVisualContent = SafeHasVisualContent(para)
End Function

'================================================================================
' VALIDATE PROPOSITION TYPE - FUNCIONALIDADE 3
'================================================================================
Private Function ValidatePropositionType(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim firstPara As Paragraph
    Dim firstWord As String
    Dim paraText As String
    Dim i As Long
    Dim userResponse As VbMsgBoxResult
    
    ' Encontra o primeiro parÃ¡grafo com texto
    For i = 1 To doc.Paragraphs.count
        Set firstPara = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(firstPara.Range.text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then
            Exit For
        End If
    Next i
    
    If paraText = "" Then
        LogMessage "Documento nÃ£o possui texto para validaÃ§Ã£o", LOG_LEVEL_WARNING
        ValidatePropositionType = True
        Exit Function
    End If
    
    ' Extrai a primeira palavra
    Dim words() As String
    words = Split(paraText, " ")
    If UBound(words) >= 0 Then
        firstWord = LCase(Trim(words(0)))
    End If
    
    ' Verifica se Ã© uma das proposituras vÃ¡lidas
    If firstWord = "indicaÃ§Ã£o" Or firstWord = "requerimento" Or firstWord = "moÃ§Ã£o" Then
        LogMessage "Tipo de proposiÃ§Ã£o validado: " & firstWord, LOG_LEVEL_INFO
        ValidatePropositionType = True
    Else
        ' Informa sobre documento nÃ£o-padrÃ£o e continua automaticamente
        LogMessage "Primeira palavra nÃ£o reconhecida como proposiÃ§Ã£o padrÃ£o: " & firstWord & " - continuando processamento", LOG_LEVEL_WARNING
        Application.StatusBar = "Aviso: Documento nÃ£o Ã© IndicaÃ§Ã£o/Requerimento/MoÃ§Ã£o - processando mesmo assim"
        
        ' Pequena pausa para o usuÃ¡rio visualizar a mensagem
        Dim startTime As Double
        startTime = Timer
        Do While Timer < startTime + 2  ' 2 segundos
            DoEvents
        Loop
        
        LogMessage "Processamento de documento nÃ£o-padrÃ£o autorizado automaticamente: " & firstWord, LOG_LEVEL_INFO
        ValidatePropositionType = True
    End If
    
    Exit Function

ErrorHandler:
    LogMessage "Erro na validaÃ§Ã£o do tipo de proposiÃ§Ã£o: " & Err.Description, LOG_LEVEL_ERROR
    ValidatePropositionType = False
End Function

'================================================================================
' FORMAT DOCUMENT TITLE - FUNCIONALIDADES 4 e 5
'================================================================================
Private Function FormatDocumentTitle(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim firstPara As Paragraph
    Dim paraText As String
    Dim words() As String
    Dim i As Long
    Dim newText As String
    
    ' Encontra o primeiro parÃ¡grafo com texto (apÃ³s exclusÃ£o de linhas em branco)
    For i = 1 To doc.Paragraphs.count
        Set firstPara = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(firstPara.Range.text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then
            Exit For
        End If
    Next i
    
    If paraText = "" Then
        LogMessage "Nenhum texto encontrado para formataÃ§Ã£o do tÃ­tulo", LOG_LEVEL_WARNING
        FormatDocumentTitle = True
        Exit Function
    End If
    
    ' Remove ponto final se existir
    If Right(paraText, 1) = "." Then
        paraText = Left(paraText, Len(paraText) - 1)
    End If
    
    ' Verifica se Ã© uma proposiÃ§Ã£o (para aplicar substituiÃ§Ã£o $NUMERO$/$ANO$)
    Dim isProposition As Boolean
    Dim firstWord As String
    
    words = Split(paraText, " ")
    If UBound(words) >= 0 Then
        firstWord = LCase(Trim(words(0)))
        If firstWord = "indicaÃ§Ã£o" Or firstWord = "requerimento" Or firstWord = "moÃ§Ã£o" Then
            isProposition = True
        End If
    End If
    
    ' Se for proposiÃ§Ã£o, substitui a Ãºltima palavra por $NUMERO$/$ANO$
    If isProposition And UBound(words) >= 0 Then
        ' ReconstrÃ³i o texto substituindo a Ãºltima palavra
        newText = ""
        For i = 0 To UBound(words) - 1
            If i > 0 Then newText = newText & " "
            newText = newText & words(i)
        Next i
        
        ' Adiciona $NUMERO$/$ANO$ no lugar da Ãºltima palavra
        If newText <> "" Then newText = newText & " "
        newText = newText & "$NUMERO$/$ANO$"
    Else
        ' Se nÃ£o for proposiÃ§Ã£o, mantÃ©m o texto original
        newText = paraText
    End If
    
    ' SEMPRE aplica formataÃ§Ã£o de tÃ­tulo: caixa alta, negrito, sublinhado
    firstPara.Range.text = UCase(newText) & PARAGRAPH_BREAK
    
    ' FormataÃ§Ã£o completa do tÃ­tulo (primeira linha)
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
        .SpaceAfter = 6  ' Pequeno espaÃ§o apÃ³s o tÃ­tulo
    End With
    
    If isProposition Then
        LogMessage "TÃ­tulo de proposiÃ§Ã£o formatado: " & newText & " (centralizado, caixa alta, negrito, sublinhado)", LOG_LEVEL_INFO
    Else
        LogMessage "Primeira linha formatada como tÃ­tulo: " & newText & " (centralizado, caixa alta, negrito, sublinhado)", LOG_LEVEL_INFO
    End If
    
    FormatDocumentTitle = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formataÃ§Ã£o do tÃ­tulo: " & Err.Description, LOG_LEVEL_ERROR
    FormatDocumentTitle = False
End Function

'================================================================================
' FORMAT CONSIDERANDO PARAGRAPHS - OTIMIZADO E SIMPLIFICADO - FUNCIONALIDADE 8
'================================================================================
Private Function FormatConsiderandoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim rng As Range
    Dim totalFormatted As Long
    Dim i As Long
    
    ' Percorre todos os parÃ¡grafos procurando por "considerando" no inÃ­cio
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Verifica se o parÃ¡grafo comeÃ§a com "considerando" (ignorando maiÃºsculas/minÃºsculas)
        If Len(paraText) >= 12 And LCase(Left(paraText, 12)) = "considerando" Then
            ' Verifica se apÃ³s "considerando" vem espaÃ§o, vÃ­rgula, ponto-e-vÃ­rgula ou fim da linha
            Dim nextChar As String
            If Len(paraText) > 12 Then
                nextChar = Mid(paraText, 13, 1)
                If nextChar = " " Or nextChar = "," Or nextChar = ";" Or nextChar = ":" Then
                    ' Ã‰ realmente "considerando" no inÃ­cio do parÃ¡grafo
                    Set rng = para.Range
                    
                    ' CORREÃ‡ÃƒO: Usa Find/Replace para preservar espaÃ§amento
                    With rng.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .text = "considerando"
                        .Replacement.text = "CONSIDERANDO"
                        .Replacement.Font.Bold = True
                        .MatchCase = False
                        .MatchWholeWord = False  ' CORREÃ‡ÃƒO: False para nÃ£o exigir palavra completa
                        .Forward = True
                        .Wrap = wdFindStop
                        
                        ' Limita a busca ao inÃ­cio do parÃ¡grafo
                        rng.End = rng.Start + 15  ' Seleciona apenas o inÃ­cio para evitar mÃºltiplas substituiÃ§Ãµes
                        
                        If .Execute(Replace:=True) Then
                            totalFormatted = totalFormatted + 1
                        End If
                    End With
                End If
            Else
                ' ParÃ¡grafo contÃ©m apenas "considerando"
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
    
    LogMessage "FormataÃ§Ã£o 'considerando' aplicada: " & totalFormatted & " ocorrÃªncias em negrito e caixa alta", LOG_LEVEL_INFO
    FormatConsiderandoParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formataÃ§Ã£o 'considerando': " & Err.Description, LOG_LEVEL_ERROR
    FormatConsiderandoParagraphs = False
End Function

'================================================================================
' APPLY TEXT REPLACEMENTS - FUNCIONALIDADES 10 e 11
'================================================================================
Private Function ApplyTextReplacements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Dim replacementCount As Long
    
    Set rng = doc.Range
    
    ' Funcionalidade 10: Substitui variantes de "d'Oeste"
    Dim dOesteVariants() As String
    Dim i As Long
    
    ' Define as variantes possÃ­veis dos 3 primeiros caracteres de "d'Oeste"
    ReDim dOesteVariants(0 To 15)
    dOesteVariants(0) = "d'O"   ' Original
    dOesteVariants(1) = "dÂ´O"   ' Acento agudo
    dOesteVariants(2) = "d`O"   ' Acento grave
    dOesteVariants(3) = "d" & ChrW(8220) & "O"   ' Aspas curvas esquerda
    dOesteVariants(4) = "d'o"   ' MinÃºscula
    dOesteVariants(5) = "dÂ´o"
    dOesteVariants(6) = "d`o"
    dOesteVariants(7) = "d" & ChrW(8220) & "o"
    dOesteVariants(8) = "D'O"   ' MaiÃºscula no D
    dOesteVariants(9) = "DÂ´O"
    dOesteVariants(10) = "D`O"
    dOesteVariants(11) = "D" & ChrW(8220) & "O"
    dOesteVariants(12) = "D'o"
    dOesteVariants(13) = "DÂ´o"
    dOesteVariants(14) = "D`o"
    dOesteVariants(15) = "D" & ChrW(8220) & "o"
    
    For i = 0 To UBound(dOesteVariants)
        With rng.Find
            .ClearFormatting
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
            
            Do While .Execute(Replace:=True)
                replacementCount = replacementCount + 1
            Loop
        End With
    Next i
    
    ' Funcionalidade 11: Substitui variantes de "- Vereador -"
    Set rng = doc.Range
    Dim vereadorVariants() As String
    ReDim vereadorVariants(0 To 7)
    
    ' Variantes dos caracteres inicial e final
    vereadorVariants(0) = "- Vereador -"    ' Original
    vereadorVariants(1) = "â€“ Vereador â€“"    ' TravessÃ£o
    vereadorVariants(2) = "â€” Vereador â€”"    ' Em dash
    vereadorVariants(3) = "- vereador -"    ' MinÃºscula
    vereadorVariants(4) = "â€“ vereador â€“"
    vereadorVariants(5) = "â€” vereador â€”"
    vereadorVariants(6) = "-Vereador-"      ' Sem espaÃ§os
    vereadorVariants(7) = "â€“Vereadorâ€“"
    
    For i = 0 To UBound(vereadorVariants)
        If vereadorVariants(i) <> "- Vereador -" Then
            With rng.Find
                .ClearFormatting
                .text = vereadorVariants(i)
                .Replacement.text = "- Vereador -"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                
                Do While .Execute(Replace:=True)
                    replacementCount = replacementCount + 1
                Loop
            End With
        End If
    Next i
    
    LogMessage "SubstituiÃ§Ãµes de texto aplicadas: " & replacementCount & " substituiÃ§Ãµes realizadas", LOG_LEVEL_INFO
    ApplyTextReplacements = True
    Exit Function

ErrorHandler:
    LogMessage "Erro nas substituiÃ§Ãµes de texto: " & Err.Description, LOG_LEVEL_ERROR
    ApplyTextReplacements = False
End Function

'================================================================================
' FORMAT JUSTIFICATIVA/ANEXO PARAGRAPHS - FORMATAÃ‡ÃƒO ESPECÃFICA
'================================================================================
Private Function FormatJustificativaAnexoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim cleanText As String
    Dim normalizedHeading As String
    Dim originalEnd As String
    Dim previousAlerts As WdAlertLevel
    Dim justificativaLabel As String
    Dim anexoEnd As String
    Dim anexoText As String
    Dim finalAnexoHeading As String
    Dim i As Long
    Dim formattedCount As Long
    Dim vereadorCount As Long
    
    ' Percorre todos os parÃ¡grafos do documento
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        
        ' NÃ£o processa parÃ¡grafos com conteÃºdo visual
        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            normalizedHeading = NormalizeHeadingKey(paraText)
            
            ' Remove pontuaÃ§Ã£o final para anÃ¡lise mais precisa
            cleanText = paraText
            ' Remove pontos, vÃ­rgulas, dois-pontos, ponto-e-vÃ­rgula do final
            Do While Len(cleanText) > 0 And (Right(cleanText, 1) = "." Or Right(cleanText, 1) = "," Or Right(cleanText, 1) = ":" Or Right(cleanText, 1) = ";")
                cleanText = Left(cleanText, Len(cleanText) - 1)
            Loop
            cleanText = Trim(LCase(cleanText))
            
            ' REQUISITO 1: FormataÃ§Ã£o de "Justificativa:" (case insensitive)
            If normalizedHeading = "justificativa" Then
                ' Padroniza o texto mantendo pontuaÃ§Ã£o original se houver
                originalEnd = ""
                If Len(paraText) > Len(cleanText) Then
                    originalEnd = Right(paraText, Len(paraText) - Len(cleanText))
                End If

                previousAlerts = Application.DisplayAlerts
                Application.DisplayAlerts = wdAlertsNone  ' Evita prompts visuais indesejados
                para.Range.text = "Justificativa" & originalEnd & PARAGRAPH_BREAK
                Application.DisplayAlerts = previousAlerts
                
                ' Aplica formataÃ§Ã£o especÃ­fica para Justificativa:
                With para.Format
                    .leftIndent = 0               ' Recuo Ã  esquerda = 0
                    .firstLineIndent = 0          ' Recuo da 1Âª linha = 0
                    .RightIndent = 0              ' Recuo Ã  direita = 0
                    .alignment = wdAlignParagraphCenter  ' Alinhamento centralizado
                    .SpaceBefore = 12
                    .SpaceAfter = 6
                End With
                
                With para.Range.Font
                    .Bold = True                  ' Negrito
                End With
                
                justificativaLabel = "Justificativa" & originalEnd
                LogMessage "ParÃ¡grafo '" & justificativaLabel & "' formatado (centralizado, negrito, sem recuos)", LOG_LEVEL_INFO
                formattedCount = formattedCount + 1
                
            ' REQUISITO 1: FormataÃ§Ã£o de variaÃ§Ãµes de "vereador"
            ElseIf IsVereadorPattern(cleanText) Then
                ' REQUISITO 2: Formatar parÃ¡grafo ANTERIOR a "vereador" PRIMEIRO
                If i > 1 Then
                    Dim paraPrev As Paragraph
                    Set paraPrev = doc.Paragraphs(i - 1)
                    
                    ' Verifica se o parÃ¡grafo anterior nÃ£o tem conteÃºdo visual
                    If Not HasVisualContent(paraPrev) Then
                        Dim prevText As String
                        prevText = Trim(Replace(Replace(paraPrev.Range.text, vbCr, ""), vbLf, ""))
                        
                        ' SÃ³ formata se o parÃ¡grafo anterior tem conteÃºdo textual
                        If prevText <> "" Then
                            ' FormataÃ§Ã£o COMPLETA do parÃ¡grafo anterior
                            With paraPrev.Format
                                .leftIndent = 0                      ' Recuo Ã  esquerda = 0
                                .firstLineIndent = 0                 ' Recuo da 1Âª linha = 0
                                .RightIndent = 0                     ' Recuo Ã  direita = 0
                                .alignment = wdAlignParagraphCenter  ' Alinhamento centralizado
                                .SpaceBefore = 12
                                .SpaceAfter = 6
                            End With
                            
                            ' FORÃ‡A os recuos zerados com chamadas individuais para garantia
                            paraPrev.Format.leftIndent = 0
                            paraPrev.Format.firstLineIndent = 0
                            paraPrev.Format.RightIndent = 0
                            paraPrev.Format.alignment = wdAlignParagraphCenter
                            
                            With paraPrev.Range.Font
                                .Bold = True                         ' Negrito
                            End With
                            
                            ' Aplica caixa alta ao parÃ¡grafo anterior
                            paraPrev.Range.text = UCase(prevText) & PARAGRAPH_BREAK
                            
                            LogMessage "ParÃ¡grafo anterior a '- Vereador -' formatado (centralizado, caixa alta, negrito, sem recuos): " & Left(UCase(prevText), 30) & "...", LOG_LEVEL_INFO
                        End If
                    End If
                End If
                
                ' Agora formata o parÃ¡grafo "- Vereador -"
                With para.Format
                    .leftIndent = 0               ' Recuo Ã  esquerda = 0
                    .firstLineIndent = 0          ' Recuo da 1Âª linha = 0
                    .RightIndent = 0              ' Recuo Ã  direita = 0
                    .alignment = wdAlignParagraphCenter  ' Alinhamento centralizado
                    .SpaceBefore = 12
                    .SpaceAfter = 6
                End With
                
                ' FORÃ‡A os recuos zerados com chamadas individuais para garantia
                para.Format.leftIndent = 0
                para.Format.firstLineIndent = 0
                para.Format.RightIndent = 0
                para.Format.alignment = wdAlignParagraphCenter
                
                With para.Range.Font
                    .Bold = True                  ' Negrito
                End With
                
                ' Padroniza o texto
                para.Range.text = "- Vereador -" & PARAGRAPH_BREAK
                
                LogMessage "ParÃ¡grafo '- Vereador -' formatado (centralizado, negrito, sem recuos)", LOG_LEVEL_INFO
                vereadorCount = vereadorCount + 1
                formattedCount = formattedCount + 1
                
            ' REQUISITO 3: FormataÃ§Ã£o de variaÃ§Ãµes de "anexo" ou "anexos"
            ElseIf normalizedHeading = "anexo" Or normalizedHeading = "anexos" Then
                ' Aplica formataÃ§Ã£o especÃ­fica para Anexo/Anexos
                With para.Format
                    .leftIndent = 0               ' Recuo Ã  esquerda = 0
                    .firstLineIndent = 0          ' Recuo da 1Âª linha = 0
                    .RightIndent = 0              ' Recuo Ã  direita = 0
                    .alignment = wdAlignParagraphLeft    ' Alinhamento Ã  esquerda
                    .SpaceBefore = 12
                    .SpaceAfter = 6
                End With
                
                ' FORÃ‡A os recuos zerados com chamadas individuais para garantia
                para.Format.leftIndent = 0
                para.Format.firstLineIndent = 0
                para.Format.RightIndent = 0
                para.Format.alignment = wdAlignParagraphLeft
                
                With para.Range.Font
                    .Bold = True                  ' Negrito
                End With
                
                ' Padroniza o texto mantendo pontuaÃ§Ã£o original se houver
                anexoEnd = ""
                If Len(paraText) > Len(cleanText) Then
                    anexoEnd = Right(paraText, Len(paraText) - Len(cleanText))
                End If
                
                If normalizedHeading = "anexo" Then
                    anexoText = "Anexo"
                Else
                    anexoText = "Anexos"
                End If
                finalAnexoHeading = anexoText & anexoEnd
                para.Range.text = finalAnexoHeading & PARAGRAPH_BREAK
                
                LogMessage "ParÃ¡grafo '" & finalAnexoHeading & "' formatado (alinhado Ã  esquerda, negrito, sem recuos)", LOG_LEVEL_INFO
                formattedCount = formattedCount + 1
            End If
        End If
    Next i
    
    LogMessage "FormataÃ§Ã£o especial concluÃ­da: " & formattedCount & " parÃ¡grafos formatados (incluindo " & vereadorCount & " '- Vereador -')", LOG_LEVEL_INFO
    FormatJustificativaAnexoParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formataÃ§Ã£o de parÃ¡grafos especiais: " & Err.Description, LOG_LEVEL_ERROR
    FormatJustificativaAnexoParagraphs = False
End Function

'================================================================================
' FUNÃ‡Ã•ES AUXILIARES PARA DETECÃ‡ÃƒO DE PADRÃ•ES
'================================================================================
Private Function IsVereadorPattern(text As String) As Boolean
    ' Remove espaÃ§os extras para anÃ¡lise
    Dim cleanText As String
    cleanText = Trim(text)
    
    ' Remove hifens/travessÃµes do inÃ­cio e fim e espaÃ§os adjacentes
    cleanText = Trim(cleanText)
    If Left(cleanText, 1) = "-" Or Left(cleanText, 1) = "â€“" Or Left(cleanText, 1) = "â€”" Then
        cleanText = Trim(Mid(cleanText, 2))
    End If
    If Right(cleanText, 1) = "-" Or Right(cleanText, 1) = "â€“" Or Right(cleanText, 1) = "â€”" Then
        cleanText = Trim(Left(cleanText, Len(cleanText) - 1))
    End If
    
    ' Verifica se o que sobrou Ã© alguma variaÃ§Ã£o de "vereador"
    cleanText = LCase(Trim(cleanText))
    IsVereadorPattern = (cleanText = "vereador" Or cleanText = "vereadora")
End Function

Private Function NormalizeHeadingKey(text As String) As String
    Dim normalized As String

    normalized = Replace(Replace(text, vbCr, ""), vbLf, "")
    normalized = Trim$(normalized)

    Do While Len(normalized) > 0 And InStr(":;.,", Right$(normalized, 1)) > 0
        normalized = Left$(normalized, Len(normalized) - 1)
    Loop

    NormalizeHeadingKey = LCase$(normalized)
End Function

Private Function IsJustificativaHeading(text As String) As Boolean
    IsJustificativaHeading = (NormalizeHeadingKey(text) = "justificativa")
End Function

Private Function IsAnexoPattern(text As String) As Boolean
    Dim normalized As String
    normalized = NormalizeHeadingKey(text)
    IsAnexoPattern = (normalized = "anexo" Or normalized = "anexos")
End Function

'================================================================================
' SUBROTINA PÃšBLICA: ABRIR PASTA DE LOGS
'================================================================================
Public Sub AbrirPastaLogs()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim logsFolder As String
    
    ' Tenta obter documento ativo
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler
    
    ' Define pasta de logs no diretÃ³rio do usuÃ¡rio
    logsFolder = EnsureUserDataDirectory(LOG_FOLDER_NAME)
    If Len(logsFolder) = 0 Then
        logsFolder = Environ("TEMP")
    End If
    
    ' Abre a pasta no Windows Explorer
    shell "explorer.exe """ & logsFolder & """", vbNormalFocus
    
    Application.StatusBar = "Pasta de logs aberta: " & logsFolder
    
    ' Log da operaÃ§Ã£o se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Pasta de logs aberta pelo usuÃ¡rio: " & logsFolder, LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir pasta de logs"
    
    ' Fallback: tenta abrir pasta temporÃ¡ria
    On Error Resume Next
    shell "explorer.exe """ & Environ("TEMP") & """", vbNormalFocus
    If Err.Number = 0 Then
        Application.StatusBar = "Pasta temporÃ¡ria aberta como alternativa"
    Else
        Application.StatusBar = "NÃ£o foi possÃ­vel abrir pasta de logs"
    End If
End Sub

'================================================================================
' SUBROTINA PÃšBLICA: ABRIR REPOSITÃ“RIO GITHUB - FUNCIONALIDADE 9
'================================================================================
Public Sub AbrirRepositorioGitHub()
    On Error GoTo ErrorHandler
    
    Dim repoURL As String
    Dim shellResult As Long
    
    ' URL do repositÃ³rio do projeto
    repoURL = "https://github.com/chrmsantos/chainsaw-fprops"
    
    ' Abre o link no navegador padrÃ£o
    shellResult = shell("rundll32.exe url.dll,FileProtocolHandler " & repoURL, vbNormalFocus)
    
    If shellResult > 0 Then
        Application.StatusBar = "RepositÃ³rio GitHub aberto no navegador"
        
        ' Log da operaÃ§Ã£o se sistema de log estiver ativo
        If loggingEnabled Then
            LogMessage "RepositÃ³rio GitHub aberto pelo usuÃ¡rio: " & repoURL, LOG_LEVEL_INFO
        End If
    Else
        GoTo ErrorHandler
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir repositÃ³rio GitHub"
    
    ' Log do erro se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Erro ao abrir repositÃ³rio GitHub: " & Err.Description, LOG_LEVEL_ERROR
    End If
    
    ' Fallback: tenta copiar URL para a Ã¡rea de transferÃªncia
    On Error Resume Next
    Dim dataObj As Object
    Set dataObj = CreateObject("htmlfile").parentWindow.clipboardData
    dataObj.setData "text", repoURL
    
    If Err.Number = 0 Then
        Application.StatusBar = "URL copiada para Ã¡rea de transferÃªncia: " & repoURL
    Else
        Application.StatusBar = "NÃ£o foi possÃ­vel abrir o repositÃ³rio. URL: " & repoURL
    End If
End Sub

'================================================================================
' SISTEMA DE BACKUP - FUNCIONALIDADE DE SEGURANÃ‡A
'================================================================================
Private Function CreateDocumentBackup(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ' NÃ£o faz backup se documento nÃ£o foi salvo
    If doc.Path = "" Then
        LogMessage "Backup ignorado - documento nÃ£o salvo", LOG_LEVEL_INFO
        CreateDocumentBackup = True
        Exit Function
    End If
    
    Dim backupFolder As String
    Dim fso As Object
    Dim docName As String
    Dim docExtension As String
    Dim timeStamp As String
    Dim backupFileName As String
    Dim backupFilePath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Garante a existÃªncia da pasta de backup e obtÃ©m o caminho final
    backupFolder = EnsureBackupDirectory(doc)
    If Len(backupFolder) = 0 Then
        LogMessage "Backup nÃ£o criado - pasta de backup indisponÃ­vel", LOG_LEVEL_ERROR
        CreateDocumentBackup = False
        Exit Function
    End If
    
    ' Extrai nome e extensÃ£o do documento
    docName = fso.GetBaseName(doc.Name)
    docExtension = fso.GetExtensionName(doc.Name)
    
    ' Cria timestamp para o backup
    timeStamp = Format(Now, "yyyy-mm-dd_HHmmss")
    
    ' Nome do arquivo de backup
    backupFileName = docName & "_backup_" & timeStamp & "." & docExtension
    backupFilePath = backupFolder & "\" & backupFileName
    
    ' Salva uma cÃ³pia do documento como backup
    Application.StatusBar = "Criando backup do documento..."
    
    ' Salva o documento atual primeiro para garantir que estÃ¡ atualizado
    doc.Save
    
    ' Cria uma cÃ³pia do arquivo usando FileSystemObject
    fso.CopyFile doc.FullName, backupFilePath, True
    
    ' Limpa backups antigos se necessÃ¡rio
    CleanOldBackups backupFolder, docName
    
    LogMessage "Backup criado com sucesso: " & backupFileName, LOG_LEVEL_INFO
    Application.StatusBar = "Backup criado - processando documento..."
    
    CreateDocumentBackup = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao criar backup: " & Err.Description, LOG_LEVEL_ERROR
    CreateDocumentBackup = False
End Function

'================================================================================
' LIMPEZA DE BACKUPS ANTIGOS - SIMPLIFICADO
'================================================================================
Private Sub CleanOldBackups(backupFolder As String, docBaseName As String)
    On Error Resume Next
    
    ' Limpeza simplificada - sÃ³ remove se houver muitos arquivos
    Dim fso As Object
    Dim folder As Object
    Dim filesCount As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(backupFolder)
    
    filesCount = folder.Files.count
    
    ' Se hÃ¡ mais de 15 arquivos na pasta de backup, registra aviso
    If filesCount > 15 Then
        LogMessage "Muitos backups na pasta (" & filesCount & " arquivos) - considere limpeza manual", LOG_LEVEL_WARNING
    End If
End Sub

'================================================================================
' SUBROTINA PÃšBLICA: ABRIR PASTA DE BACKUPS
'================================================================================
Public Sub AbrirPastaBackups()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim backupFolder As String
    
    ' Tenta obter documento ativo
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler
    
    ' Define pasta de backup na Ã¡rea de documentos do usuÃ¡rio
    backupFolder = EnsureUserDataDirectory(BACKUP_FOLDER_NAME)
    If Len(backupFolder) = 0 Then
        Application.StatusBar = "NÃ£o foi possÃ­vel localizar pasta de backups"
        If loggingEnabled Then
            LogMessage "Pasta de backups indisponÃ­vel", LOG_LEVEL_WARNING
        End If
        Exit Sub
    End If
    
    ' Abre a pasta no Windows Explorer
    shell "explorer.exe """ & backupFolder & """", vbNormalFocus
    
    Application.StatusBar = "Pasta de backups aberta: " & backupFolder
    
    ' Log da operaÃ§Ã£o se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Pasta de backups aberta pelo usuÃ¡rio: " & backupFolder, LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir pasta de backups"
    LogMessage "Erro ao abrir pasta de backups: " & Err.Description, LOG_LEVEL_ERROR
    
    ' Fallback: tenta abrir pasta do documento
    On Error Resume Next
    Dim userDocs As String
    userDocs = GetUserDocumentsPath()
    If Len(userDocs) > 0 Then
        shell "explorer.exe """ & userDocs & """", vbNormalFocus
        Application.StatusBar = "Pasta Documentos aberta como alternativa"
    Else
        Application.StatusBar = "NÃ£o foi possÃ­vel abrir pasta de backups"
    End If
End Sub

'================================================================================
' CLEAN MULTIPLE SPACES - LIMPEZA FINAL DE ESPAÃ‡OS MÃšLTIPLOS
'================================================================================
Private Function CleanMultipleSpaces(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Limpando espaÃ§os mÃºltiplos..."
    
    Dim rng As Range
    Dim spacesRemoved As Long
    Dim totalOperations As Long
    Dim hasDoubleSpaces As Boolean
    Dim hasTabs As Boolean
    Dim docSnapshot As String

    If doc Is Nothing Then
        CleanMultipleSpaces = True
        Exit Function
    End If

    docSnapshot = doc.Content.text
    hasDoubleSpaces = (InStr(docSnapshot, "  ") > 0)
    hasTabs = (InStr(docSnapshot, vbTab) > 0)
    If Not hasDoubleSpaces And Not hasTabs Then
        CleanMultipleSpaces = True
        Exit Function
    End If
    docSnapshot = ""
    
    ' SUPER OTIMIZADO: OperaÃ§Ãµes consolidadas em uma Ãºnica configuraÃ§Ã£o Find
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
        
        ' OTIMIZAÃ‡ÃƒO 1: Remove espaÃ§os mÃºltiplos (2 ou mais) em uma Ãºnica operaÃ§Ã£o
        ' Usa um loop otimizado que reduz progressivamente os espaÃ§os
        Do
            .text = "  "  ' Dois espaÃ§os
            .Replacement.text = " "  ' Um espaÃ§o
            
            Dim currentReplaceCount As Long
            currentReplaceCount = 0
            
            ' Executa atÃ© nÃ£o encontrar mais duplos
            Do While .Execute(Replace:=True)
                currentReplaceCount = currentReplaceCount + 1
                spacesRemoved = spacesRemoved + 1
                ' ProteÃ§Ã£o otimizada - verifica a cada 200 operaÃ§Ãµes
                If currentReplaceCount Mod 200 = 0 Then
                    DoEvents
                    If ShouldAbortForWordHang("limpeza de espaÃ§os") Then
                        CleanMultipleSpaces = False
                        Exit Function
                    End If
                End If
                If spacesRemoved > 2000 Then Exit Do
            Loop
            
            totalOperations = totalOperations + 1
            ' Se nÃ£o encontrou mais duplos ou atingiu limite, para
            If currentReplaceCount = 0 Or totalOperations > 10 Then Exit Do
        Loop
    End With
    
    ' OTIMIZAÃ‡ÃƒO 2: OperaÃ§Ãµes de limpeza de quebras de linha consolidadas
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False  ' Usar Find/Replace simples para compatibilidade
        
        ' Remove mÃºltiplos espaÃ§os antes de quebras - mÃ©todo iterativo
        .text = "  ^p"  ' 2 espaÃ§os seguidos de quebra
        .Replacement.text = " ^p"  ' 1 espaÃ§o seguido de quebra
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved Mod 200 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("limpeza de espaÃ§os") Then
                    CleanMultipleSpaces = False
                    Exit Function
                End If
            End If
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
        ' Segunda passada para garantir limpeza completa
        .text = " ^p"  ' EspaÃ§o antes de quebra
        .Replacement.text = "^p"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved Mod 200 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("limpeza de espaÃ§os") Then
                    CleanMultipleSpaces = False
                    Exit Function
                End If
            End If
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
        ' Remove mÃºltiplos espaÃ§os depois de quebras - mÃ©todo iterativo
        .text = "^p  "  ' Quebra seguida de 2 espaÃ§os
        .Replacement.text = "^p "  ' Quebra seguida de 1 espaÃ§o
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved Mod 200 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("limpeza de espaÃ§os") Then
                    CleanMultipleSpaces = False
                    Exit Function
                End If
            End If
            If spacesRemoved > 2000 Then Exit Do
        Loop
    End With
    
    ' OTIMIZAÃ‡ÃƒO 3: Limpeza de tabs consolidada e otimizada
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False  ' Usar Find/Replace simples
        
        ' Remove mÃºltiplos tabs iterativamente
        .text = "^t^t"  ' 2 tabs
        .Replacement.text = "^t"  ' 1 tab
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved Mod 200 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("limpeza de espaÃ§os") Then
                    CleanMultipleSpaces = False
                    Exit Function
                End If
            End If
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
        ' Converte tabs para espaÃ§os
        .text = "^t"
        .Replacement.text = " "
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved Mod 200 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("limpeza de espaÃ§os") Then
                    CleanMultipleSpaces = False
                    Exit Function
                End If
            End If
            If spacesRemoved > 2000 Then Exit Do
        Loop
    End With
    
    ' OTIMIZAÃ‡ÃƒO 4: VerificaÃ§Ã£o final ultra-rÃ¡pida de espaÃ§os duplos remanescentes
    Set rng = doc.Range
    With rng.Find
        .text = "  "
        .Replacement.text = " "
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop  ' Mais rÃ¡pido que wdFindContinue
        
        Dim finalCleanCount As Long
        Do While .Execute(Replace:=True) And finalCleanCount < 100
            finalCleanCount = finalCleanCount + 1
            spacesRemoved = spacesRemoved + 1
            If finalCleanCount Mod 50 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("limpeza de espaÃ§os") Then
                    CleanMultipleSpaces = False
                    Exit Function
                End If
            End If
        Loop
    End With
    
    ' PROTEÃ‡ÃƒO ESPECÃFICA: Garante espaÃ§o apÃ³s CONSIDERANDO
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchCase = False
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        
        ' Corrige CONSIDERANDO grudado com a prÃ³xima palavra
        .text = "CONSIDERANDOa"
        .Replacement.text = "CONSIDERANDO a"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
            If spacesRemoved Mod 200 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("limpeza de espaÃ§os") Then
                    CleanMultipleSpaces = False
                    Exit Function
                End If
            End If
        Loop
        
        .text = "CONSIDERANDOe"
        .Replacement.text = "CONSIDERANDO e"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
            If spacesRemoved Mod 200 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("limpeza de espaÃ§os") Then
                    CleanMultipleSpaces = False
                    Exit Function
                End If
            End If
        Loop
        
        .text = "CONSIDERANDOo"
        .Replacement.text = "CONSIDERANDO o"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
            If spacesRemoved Mod 200 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("limpeza de espaÃ§os") Then
                    CleanMultipleSpaces = False
                    Exit Function
                End If
            End If
        Loop
        
        .text = "CONSIDERANDOq"
        .Replacement.text = "CONSIDERANDO q"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
            If spacesRemoved Mod 200 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("limpeza de espaÃ§os") Then
                    CleanMultipleSpaces = False
                    Exit Function
                End If
            End If
        Loop
    End With
    
    LogMessage "Limpeza de espaÃ§os concluÃ­da: " & spacesRemoved & " correÃ§Ãµes aplicadas (com proteÃ§Ã£o CONSIDERANDO)", LOG_LEVEL_INFO
    CleanMultipleSpaces = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na limpeza de espaÃ§os mÃºltiplos: " & Err.Description, LOG_LEVEL_WARNING
    CleanMultipleSpaces = False ' NÃ£o falha o processo por isso
End Function

'================================================================================
' LIMIT SEQUENTIAL EMPTY LINES - CONTROLA LINHAS VAZIAS SEQUENCIAIS
'================================================================================
Private Function LimitSequentialEmptyLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Controlando linhas em branco sequenciais..."
    
    Dim docSnapshot As String

    If doc Is Nothing Then
        LimitSequentialEmptyLines = True
        Exit Function
    End If

    docSnapshot = doc.Content.text
    If InStr(docSnapshot, PARAGRAPH_BREAK & PARAGRAPH_BREAK & PARAGRAPH_BREAK) = 0 Then
        LimitSequentialEmptyLines = True
        Exit Function
    End If
    docSnapshot = ""

    ' IDENTIFICAÃ‡ÃƒO DO SEGUNDO PARÃGRAFO PARA PROTEÃ‡ÃƒO
    Dim secondParaIndex As Long
    secondParaIndex = GetSecondParagraphIndex(doc)
    
    ' SUPER OTIMIZADO: Usa Find/Replace com wildcard para operaÃ§Ã£o muito mais rÃ¡pida
    Dim rng As Range
    Dim linesRemoved As Long
    Dim totalReplaces As Long
    Dim passCount As Long
    
    passCount = 1 ' Inicializa contador de passadas
    
    Set rng = doc.Range
    
    ' MÃ‰TODO ULTRA-RÃPIDO: Remove mÃºltiplas quebras consecutivas usando wildcard
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False  ' Usar Find/Replace simples para compatibilidade
        
        ' Remove mÃºltiplas quebras consecutivas iterativamente
        .text = "^p^p^p^p"  ' 4 quebras
        .Replacement.text = "^p^p"  ' 2 quebras
        
        Do While .Execute(Replace:=True)
            linesRemoved = linesRemoved + 1
            totalReplaces = totalReplaces + 1
            If totalReplaces > 500 Then Exit Do
            If linesRemoved Mod 50 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("controle de linhas vazias") Then
                    LimitSequentialEmptyLines = False
                    Exit Function
                End If
            End If
        Loop
        
        ' Remove 3 quebras -> 2 quebras
        .text = "^p^p^p"  ' 3 quebras
        .Replacement.text = "^p^p"  ' 2 quebras
        
        Do While .Execute(Replace:=True)
            linesRemoved = linesRemoved + 1
            totalReplaces = totalReplaces + 1
            If totalReplaces > 500 Then Exit Do
            If linesRemoved Mod 50 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("controle de linhas vazias") Then
                    LimitSequentialEmptyLines = False
                    Exit Function
                End If
            End If
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
            If secondPassCount Mod 50 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("controle de linhas vazias") Then
                    LimitSequentialEmptyLines = False
                    Exit Function
                End If
            End If
        Loop
    End With
    
    ' VERIFICAÃ‡ÃƒO FINAL: Garantir que nÃ£o hÃ¡ mais de 1 linha vazia consecutiva
    If secondPassCount > 0 Then passCount = passCount + 1
    
    ' MÃ©todo hÃ­brido: Find/Replace para casos simples + loop apenas se necessÃ¡rio
    Set rng = doc.Range
    With rng.Find
        .text = "^p^p^p"  ' 3 quebras (2 linhas vazias + conteÃºdo)
        .Replacement.text = "^p^p"  ' 2 quebras (1 linha vazia + conteÃºdo)
        .MatchWildcards = False
        
        Dim finalPassCount As Long
        Do While .Execute(Replace:=True) And finalPassCount < 100
            finalPassCount = finalPassCount + 1
            linesRemoved = linesRemoved + 1
            If finalPassCount Mod 50 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("controle de linhas vazias") Then
                    LimitSequentialEmptyLines = False
                    Exit Function
                End If
            End If
        Loop
    End With
    
    If finalPassCount > 0 Then passCount = passCount + 1
    
    ' FALLBACK OTIMIZADO: Se ainda hÃ¡ problemas, usa mÃ©todo tradicional limitado
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
            
            ' Verifica se o parÃ¡grafo estÃ¡ vazio
            If paraText = "" And Not HasVisualContent(para) Then
                emptyLineCount = emptyLineCount + 1
                
                ' Se jÃ¡ temos mais de 1 linha vazia consecutiva, remove esta
                If emptyLineCount > 1 Then
                    para.Range.Delete
                    fallbackRemoved = fallbackRemoved + 1
                    linesRemoved = linesRemoved + 1
                    ' NÃ£o incrementa i pois removemos um parÃ¡grafo
                Else
                    i = i + 1
                End If
            Else
                ' Se encontrou conteÃºdo, reseta o contador
                emptyLineCount = 0
                i = i + 1
            End If
            
            ' Responsividade e proteÃ§Ã£o otimizadas
            If fallbackRemoved Mod 10 = 0 Then
                DoEvents
                If ShouldAbortForWordHang("controle de linhas vazias") Then
                    LimitSequentialEmptyLines = False
                    Exit Function
                End If
            End If
            If i > 500 Then Exit Do ' ProteÃ§Ã£o adicional
        Loop
    End If
    
    LogMessage "Controle de linhas vazias concluÃ­do em " & passCount & " passada(s): " & linesRemoved & " linhas excedentes removidas (mÃ¡ximo 1 sequencial)", LOG_LEVEL_INFO
    LimitSequentialEmptyLines = True
    Exit Function

ErrorHandler:
    LogMessage "Erro no controle de linhas vazias: " & Err.Description, LOG_LEVEL_WARNING
    LimitSequentialEmptyLines = False ' NÃ£o falha o processo por isso
End Function

'================================================================================
' CONFIGURE DOCUMENT VIEW - CONFIGURAÃ‡ÃƒO DE VISUALIZAÃ‡ÃƒO - #MODIFIED
'================================================================================
Private Function ConfigureDocumentView(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Configurando visualizaÃ§Ã£o do documento..."
    
    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow
    
    ' Configura APENAS o zoom para 110% - todas as outras configuraÃ§Ãµes sÃ£o preservadas
    With docWindow.View
        .Zoom.Percentage = 110
        ' NÃƒO altera mais o tipo de visualizaÃ§Ã£o - preserva o original
    End With
    
    ' Remove configuraÃ§Ãµes que alteravam configuraÃ§Ãµes globais do Word
    ' Estas configuraÃ§Ãµes sÃ£o agora preservadas do estado original
    
    LogMessage "VisualizaÃ§Ã£o configurada: zoom definido para 110%, demais configuraÃ§Ãµes preservadas"
    ConfigureDocumentView = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao configurar visualizaÃ§Ã£o: " & Err.Description, LOG_LEVEL_WARNING
    ConfigureDocumentView = False ' NÃ£o falha o processo por isso
End Function

'================================================================================
' SALVAR E SAIR - SUBROTINA PÃšBLICA PROFISSIONAL E ROBUSTA
'================================================================================
Public Sub SalvarESair()
    On Error GoTo CriticalErrorHandler
    
    Dim startTime As Date
    startTime = Now
    
    Application.StatusBar = "Verificando documentos abertos..."
    LogMessage "Iniciando processo de salvar e sair - verificaÃ§Ã£o de documentos", LOG_LEVEL_INFO
    
    ' Verifica se hÃ¡ documentos abertos
    If Application.Documents.count = 0 Then
        Application.StatusBar = "Nenhum documento aberto - encerrando Word"
        LogMessage "Nenhum documento aberto - encerrando aplicaÃ§Ã£o", LOG_LEVEL_INFO
        Application.Quit SaveChanges:=wdDoNotSaveChanges
        Exit Sub
    End If
    
    ' Coleta informaÃ§Ãµes sobre documentos nÃ£o salvos
    Dim unsavedDocs As Collection
    Set unsavedDocs = New Collection
    
    Dim doc As Document
    Dim i As Long
    
    ' Verifica cada documento aberto
    For i = 1 To Application.Documents.count
        Set doc = Application.Documents(i)
        
        On Error Resume Next
        ' Verifica se o documento tem alteraÃ§Ãµes nÃ£o salvas
        If doc.Saved = False Or doc.Path = "" Then
            unsavedDocs.Add doc.Name
            LogMessage "Documento nÃ£o salvo detectado: " & doc.Name
        End If
        On Error GoTo CriticalErrorHandler
    Next i
    
    ' Se nÃ£o hÃ¡ documentos nÃ£o salvos, encerra diretamente
    If unsavedDocs.count = 0 Then
        Application.StatusBar = "Todos os documentos salvos - encerrando Word"
        LogMessage "Todos os documentos estÃ£o salvos - encerrando aplicaÃ§Ã£o"
        Application.Quit SaveChanges:=wdDoNotSaveChanges
        Exit Sub
    End If
    
    ' ConstrÃ³i mensagem detalhada sobre documentos nÃ£o salvos
    Dim message As String
    Dim docList As String
    
    For i = 1 To unsavedDocs.count
        docList = docList & "â€¢ " & unsavedDocs(i) & PARAGRAPH_BREAK
    Next i
    
    message = "ATENÃ‡ÃƒO: HÃ¡ " & unsavedDocs.count & " documento(s) com alteraÃ§Ãµes nÃ£o salvas:" & PARAGRAPH_BREAK & PARAGRAPH_BREAK
    message = message & docList & PARAGRAPH_BREAK
    message = message & "Deseja salvar todos os documentos antes de sair?" & PARAGRAPH_BREAK & PARAGRAPH_BREAK
    message = message & "â€¢ SIM: Salva todos e fecha o Word" & PARAGRAPH_BREAK
    message = message & "â€¢ NÃƒO: Fecha sem salvar (PERDE as alteraÃ§Ãµes)" & PARAGRAPH_BREAK
    message = message & "â€¢ CANCELAR: Cancela a operaÃ§Ã£o"
    
    ' Apresenta opÃ§Ãµes ao usuÃ¡rio
    Application.StatusBar = "Aguardando decisÃ£o do usuÃ¡rio sobre documentos nÃ£o salvos..."
    
    Dim userChoice As VbMsgBoxResult
    userChoice = MsgBox(message, vbYesNoCancel + vbExclamation + vbDefaultButton1, _
                        "Chainsaw - Salvar e Sair (" & unsavedDocs.count & " documentos nÃ£o salvos)")
    
    Select Case userChoice
        Case vbYes
            ' UsuÃ¡rio escolheu salvar todos
            Application.StatusBar = "Salvando todos os documentos..."
            LogMessage "UsuÃ¡rio optou por salvar todos os documentos antes de sair"
            
            If SalvarTodosDocumentos() Then
                Application.StatusBar = "Documentos salvos com sucesso - encerrando Word"
                LogMessage "Todos os documentos salvos com sucesso - encerrando aplicaÃ§Ã£o"
                Application.Quit SaveChanges:=wdDoNotSaveChanges
            Else
                Application.StatusBar = "Erro ao salvar documentos - operaÃ§Ã£o cancelada"
                LogMessage "Falha ao salvar alguns documentos - operaÃ§Ã£o de sair cancelada", LOG_LEVEL_ERROR
          MsgBox "Erro ao salvar um ou mais documentos." & PARAGRAPH_BREAK & _
              "A operaÃ§Ã£o foi cancelada por seguranÃ§a." & PARAGRAPH_BREAK & PARAGRAPH_BREAK & _
              "Verifique os documentos e tente novamente.", _
                       vbCritical, "Chainsaw - Erro ao Salvar"
            End If
            
        Case vbNo
            ' UsuÃ¡rio escolheu nÃ£o salvar
            Dim confirmMessage As String
            confirmMessage = "CONFIRMAÃ‡ÃƒO FINAL:" & PARAGRAPH_BREAK & PARAGRAPH_BREAK
            confirmMessage = confirmMessage & "VocÃª estÃ¡ prestes a FECHAR O WORD SEM SALVAR " & unsavedDocs.count & " documento(s)." & PARAGRAPH_BREAK & PARAGRAPH_BREAK
            confirmMessage = confirmMessage & "TODAS AS ALTERAÃ‡Ã•ES NÃƒO SALVAS SERÃƒO PERDIDAS!" & PARAGRAPH_BREAK & PARAGRAPH_BREAK
            confirmMessage = confirmMessage & "Tem certeza absoluta?"
            
            Dim finalConfirm As VbMsgBoxResult
            finalConfirm = MsgBox(confirmMessage, vbYesNo + vbCritical + vbDefaultButton2, _
                                  "Chainsaw - CONFIRMAÃ‡ÃƒO FINAL")
            
            If finalConfirm = vbYes Then
                Application.StatusBar = "Fechando Word sem salvar alteraÃ§Ãµes..."
                LogMessage "UsuÃ¡rio confirmou fechamento sem salvar - encerrando aplicaÃ§Ã£o", LOG_LEVEL_WARNING
                Application.Quit SaveChanges:=wdDoNotSaveChanges
            Else
                Application.StatusBar = "OperaÃ§Ã£o cancelada pelo usuÃ¡rio"
                LogMessage "UsuÃ¡rio cancelou fechamento sem salvar"
                MsgBox "OperaÃ§Ã£o cancelada." & PARAGRAPH_BREAK & "Os documentos permanecem abertos.", _
                       vbInformation, "Chainsaw - OperaÃ§Ã£o Cancelada"
            End If
            
        Case vbCancel
            ' UsuÃ¡rio cancelou
            Application.StatusBar = "OperaÃ§Ã£o de sair cancelada pelo usuÃ¡rio"
            LogMessage "UsuÃ¡rio cancelou operaÃ§Ã£o de salvar e sair"
            MsgBox "OperaÃ§Ã£o cancelada." & PARAGRAPH_BREAK & "Os documentos permanecem abertos.", _
                   vbInformation, "Chainsaw - OperaÃ§Ã£o Cancelada"
    End Select
    
    Application.StatusBar = False
    LogMessage "Processo de salvar e sair concluÃ­do em " & Format(Now - startTime, "hh:mm:ss")
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "ERRO CRÃTICO na operaÃ§Ã£o Salvar e Sair #" & Err.Number & ": " & Err.Description
    
    LogMessage errDesc, LOG_LEVEL_ERROR
    Application.StatusBar = "Erro crÃ­tico - operaÃ§Ã£o cancelada"
    
    MsgBox "Erro crÃ­tico durante a operaÃ§Ã£o Salvar e Sair:" & PARAGRAPH_BREAK & PARAGRAPH_BREAK & _
        Err.Description & PARAGRAPH_BREAK & PARAGRAPH_BREAK & _
        "A operaÃ§Ã£o foi cancelada por seguranÃ§a." & PARAGRAPH_BREAK & _
           "Salve manualmente os documentos importantes.", _
           vbCritical, "Chainsaw - Erro CrÃ­tico"
End Sub

'================================================================================
' SALVAR TODOS DOCUMENTOS - FUNÃ‡ÃƒO AUXILIAR PRIVADA
'================================================================================
Private Function SalvarTodosDocumentos() As Boolean
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim i As Long
    Dim savedCount As Long
    Dim errorCount As Long
    Dim totalDocs As Long
    
    totalDocs = Application.Documents.count
    
    ' Salva cada documento individualmente
    For i = 1 To totalDocs
        Set doc = Application.Documents(i)
        
        Application.StatusBar = "Salvando documento " & i & " de " & totalDocs & ": " & doc.Name
        
        On Error Resume Next
        
        ' Se o documento nunca foi salvo (sem caminho), abre dialog
        If doc.Path = "" Then
            Dim saveDialog As Object
            Set saveDialog = Application.FileDialog(msoFileDialogSaveAs)
            
            With saveDialog
                .Title = "Salvar documento: " & doc.Name
                .InitialFileName = doc.Name
                
                If .Show = -1 Then
                    doc.SaveAs2 .SelectedItems(1)
                    If Err.Number = 0 Then
                        savedCount = savedCount + 1
                        LogMessage "Documento salvo como novo arquivo: " & doc.Name
                    Else
                        errorCount = errorCount + 1
                        LogMessage "Erro ao salvar documento como novo: " & doc.Name & " - " & Err.Description, LOG_LEVEL_ERROR
                    End If
                Else
                    errorCount = errorCount + 1
                    LogMessage "Salvamento cancelado pelo usuÃ¡rio: " & doc.Name, LOG_LEVEL_WARNING
                End If
            End With
        Else
            ' Documento jÃ¡ tem caminho, apenas salva
            doc.Save
            If Err.Number = 0 Then
                savedCount = savedCount + 1
                LogMessage "Documento salvo: " & doc.Name
            Else
                errorCount = errorCount + 1
                LogMessage "Erro ao salvar documento: " & doc.Name & " - " & Err.Description, LOG_LEVEL_ERROR
            End If
        End If
        
        On Error GoTo ErrorHandler
    Next i
    
    ' Verifica resultado
    If errorCount = 0 Then
        LogMessage "Todos os documentos salvos com sucesso: " & savedCount & " de " & totalDocs
        SalvarTodosDocumentos = True
    Else
        LogMessage "Falha parcial no salvamento: " & savedCount & " salvos, " & errorCount & " erros", LOG_LEVEL_WARNING
        SalvarTodosDocumentos = False
    End If
    
    Exit Function

ErrorHandler:
    LogMessage "Erro crÃ­tico ao salvar documentos: " & Err.Description, LOG_LEVEL_ERROR
    SalvarTodosDocumentos = False
End Function

'================================================================================
' VIEW SETTINGS PROTECTION SYSTEM - SISTEMA DE PROTEÃ‡ÃƒO DAS CONFIGURAÃ‡Ã•ES DE VISUALIZAÃ‡ÃƒO
'================================================================================

'================================================================================
' BACKUP VIEW SETTINGS - Faz backup das configuraÃ§Ãµes de visualizaÃ§Ã£o originais
'================================================================================
Private Function BackupViewSettings(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Fazendo backup das configuraÃ§Ãµes de visualizaÃ§Ã£o..."
    
    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow
    
    ' Backup das configuraÃ§Ãµes de visualizaÃ§Ã£o
    With originalViewSettings
        .ViewType = docWindow.View.Type
        ' RÃ©guas sÃ£o controladas pelo Window, nÃ£o pelo View
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
        ' .ShowAnimation removida - pode nÃ£o existir em todas as versÃµes
        .DraftFont = docWindow.View.Draft
        .WrapToWindow = docWindow.View.WrapToWindow
        .ShowPicturePlaceHolders = docWindow.View.ShowPicturePlaceHolders
        .ShowFieldShading = docWindow.View.FieldShading
        .TableGridlines = docWindow.View.TableGridlines
        ' .EnlargeFontsLessThan removida - pode nÃ£o existir em todas as versÃµes
    End With
    
    LogMessage "Backup das configuraÃ§Ãµes de visualizaÃ§Ã£o concluÃ­do"
    BackupViewSettings = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao fazer backup das configuraÃ§Ãµes de visualizaÃ§Ã£o: " & Err.Description, LOG_LEVEL_WARNING
    BackupViewSettings = False
End Function

'================================================================================
' RESTORE VIEW SETTINGS - Restaura as configuraÃ§Ãµes de visualizaÃ§Ã£o originais
'================================================================================
Private Function RestoreViewSettings(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Restaurando configuraÃ§Ãµes de visualizaÃ§Ã£o originais..."
    
    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow
    
    ' Restaura todas as configuraÃ§Ãµes originais, EXCETO o zoom
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
        
        ' ZOOM Ã© mantido em 110% - Ãºnica configuraÃ§Ã£o que permanece alterada
        .Zoom.Percentage = 110
    End With
    
    ' ConfiguraÃ§Ãµes especÃ­ficas do Window (para rÃ©guas)
    docWindow.DisplayRulers = originalViewSettings.ShowHorizontalRuler
    docWindow.DisplayVerticalRuler = originalViewSettings.ShowVerticalRuler
    
    LogMessage "ConfiguraÃ§Ãµes de visualizaÃ§Ã£o originais restauradas (zoom mantido em 110%)"
    RestoreViewSettings = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao restaurar configuraÃ§Ãµes de visualizaÃ§Ã£o: " & Err.Description, LOG_LEVEL_WARNING
    RestoreViewSettings = False
End Function

'================================================================================
' CLEANUP VIEW SETTINGS - Limpeza das variÃ¡veis de configuraÃ§Ãµes de visualizaÃ§Ã£o
'================================================================================
Private Sub CleanupViewSettings()
    On Error Resume Next
    
    ' Reinicializa a estrutura de configuraÃ§Ãµes
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
    
    LogMessage "VariÃ¡veis de configuraÃ§Ãµes de visualizaÃ§Ã£o limpas"
End Sub

'================================================================================
' REPLACE STANDARD LOCATION AND DATE PARAGRAPH
'================================================================================
Private Sub ReplacePlenarioDateParagraph(doc As Document)
    On Error GoTo ErrorHandler
    
    If doc Is Nothing Then Exit Sub
    
    Dim para As Paragraph
    Dim rawText As String
    Dim normalizedText As String
    Dim lowerText As String
    Dim monthTerms As Variant
    Dim locationTerms As Variant
    Dim monthTerm As Variant
    Dim locationTerm As Variant
    Dim monthFound As Boolean
    Dim locationFound As Boolean
    Dim targetRange As Range
    Dim replaced As Boolean
    
    monthTerms = Array("de janeiro de", "de fevereiro de", "de marÃ§o de", "de abril de", _
                       "de maio de", "de junho de", "de julho de", "de agosto de", _
                       "de setembro de", "de outubro de", "de novembro de", "de dezembro de")
    locationTerms = Array("palÃ¡cio 15 de junho", "palacio 15 de junho", "plenÃ¡rio", "plenario")
    
    Dim i As Long
    Dim j As Long
    Dim paraIndex As Long
    Dim beforeSpacing As Long
    Dim afterSpacing As Long
    Dim formattedBelow As Long
    Dim subsequentPara As Paragraph
    Dim subsequentText As String
    
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        If ShouldAbortForWordHang("substituiÃ§Ã£o de data do plenÃ¡rio") Then
            Exit Sub
        End If
        rawText = para.Range.text
        rawText = Replace(rawText, vbCr, "")
        rawText = Replace(rawText, vbLf, "")
        rawText = Trim$(rawText)
        
        If Len(rawText) = 0 Then GoTo NextParagraph
        
        normalizedText = Replace(rawText, ChrW(8220), Chr$(34))
        normalizedText = Replace(normalizedText, ChrW(8221), Chr$(34))
        lowerText = LCase$(normalizedText)
        
        monthFound = False
        For Each monthTerm In monthTerms
            If InStr(1, lowerText, CStr(monthTerm), vbBinaryCompare) > 0 Then
                monthFound = True
                Exit For
            End If
        Next monthTerm
        If Not monthFound Then GoTo NextParagraph
        
        locationFound = False
        For Each locationTerm In locationTerms
            If InStr(1, lowerText, CStr(locationTerm), vbBinaryCompare) > 0 Then
                locationFound = True
                Exit For
            End If
        Next locationTerm
        If Not locationFound Then GoTo NextParagraph
        
        If Len(lowerText) > 180 Then GoTo NextParagraph
        
        paraIndex = i
        If Not EnsureBlankLinesAroundParagraphIndex(doc, paraIndex, 2, 0, beforeSpacing, afterSpacing) Then
            beforeSpacing = CountBlankLinesBefore(doc, paraIndex)
            afterSpacing = CountBlankLinesAfter(doc, paraIndex)
        End If
        If paraIndex < 1 Or paraIndex > doc.Paragraphs.count Then GoTo NextParagraph
        Set para = doc.Paragraphs(paraIndex)
        Set targetRange = para.Range
        targetRange.text = "PlenÃ¡rio ""Dr. Tancredo Neves"", $DATAATUALEXTENSO$." & vbCr
        
        With targetRange.ParagraphFormat
            .leftIndent = 0
            .firstLineIndent = 0
            .RightIndent = 0
            .alignment = wdAlignParagraphCenter
        End With
        LogMessage "ParÃ¡grafo de plenÃ¡rio substituÃ­do e formatado (linhas em branco antes: " & beforeSpacing & ", depois: " & afterSpacing & ")", LOG_LEVEL_INFO

        formattedBelow = 0
        For j = paraIndex + 1 To doc.Paragraphs.count
            If formattedBelow >= 4 Then Exit For

            Set subsequentPara = doc.Paragraphs(j)
            subsequentText = subsequentPara.Range.text
            subsequentText = Replace(subsequentText, vbCr, "")
            subsequentText = Replace(subsequentText, vbLf, "")
            subsequentText = Trim$(subsequentText)

            With subsequentPara.Format
                .leftIndent = 0
                .firstLineIndent = 0
                .RightIndent = 0
                .alignment = wdAlignParagraphCenter
            End With

            If Len(subsequentText) > 0 Or HasVisualContent(subsequentPara) Then
                formattedBelow = formattedBelow + 1
            End If
        Next j

        If formattedBelow > 0 Then
            LogMessage "ParÃ¡grafos subsequentes Ã  data centralizados (total tratados: " & formattedBelow & ")", LOG_LEVEL_INFO
        End If
        replaced = True
        Exit For
        
NextParagraph:
    Next i
    
    If Not replaced Then
        LogMessage "ParÃ¡grafo de plenÃ¡rio nÃ£o encontrado para substituiÃ§Ã£o", LOG_LEVEL_WARNING
    End If
    
    Exit Sub
    
ErrorHandler:
    LogMessage "Erro ao processar parÃ¡grafos: " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' BACKUP DIRECTORY MANAGEMENT
'================================================================================
Private Function EnsureBackupDirectory(doc As Document) As String
    On Error GoTo ErrorHandler
    
    Dim backupPath As String
    Dim fso As Object
    
    backupPath = EnsureUserDataDirectory(BACKUP_FOLDER_NAME)
    If Len(backupPath) = 0 Then GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(backupPath) Then
        fso.CreateFolder backupPath
    End If
    
    EnsureBackupDirectory = backupPath
    Exit Function
    
ErrorHandler:
    EnsureBackupDirectory = ""
End Function
