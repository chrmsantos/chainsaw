' filepath: c:\Users\csantos\Meu Drive\Câmara\Câmara 2025\Legislativo\Projects\chainsaw\src\src\modChainsaw1.bas

'================================================================================
' MELHORIA #1: SISTEMA DE LOGGING CENTRALIZADO
'================================================================================

Private Const LOG_FILE_PATH As String = "C:\Temp\chainsaw_log.txt"
Private Const MAX_LOG_SIZE_MB As Long = 10

Private Type LogEntry
    Timestamp As Date
    Level As String
    FunctionName As String
    Message As String
    ErrorNumber As Long
    ErrorSource As String
    Context As String
    ElapsedMs As Long
End Type

Private Function InitializeLogging() As Boolean
    On Error GoTo ErrorHandler
    
    ' Check and rotate log file if needed
    If Dir(LOG_FILE_PATH) <> "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim logFile As Object
        Set logFile = fso.GetFile(LOG_FILE_PATH)
        
        ' Rotate log if > MAX_LOG_SIZE_MB
        If logFile.size > (MAX_LOG_SIZE_MB * 1024 * 1024) Then
            Dim backupPath As String
            backupPath = LOG_FILE_PATH & "." & Format(Now, "yyyyMMdd_hhmmss")
            fso.MoveFile LOG_FILE_PATH, backupPath
        End If
        
        Set logFile = Nothing
        Set fso = Nothing
    End If
    
    InitializeLogging = True
    Exit Function
    
ErrorHandler:
    InitializeLogging = False
End Function

Private Sub LogEvent(functionName As String, level As String, message As String, Optional errorNum As Long = 0, Optional context As String = "")
    On Error Resume Next
    
    Dim entry As LogEntry
    With entry
        .Timestamp = Now
        .Level = level
        .FunctionName = functionName
        .Message = message
        .ErrorNumber = errorNum
        .ErrorSource = Err.Source
        .Context = context
        .ElapsedMs = CLng((Timer - processingStartTime) * 1000)
    End With
    
    ' Write to file
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open LOG_FILE_PATH For Append As fileNum
    Print #fileNum, FormatLogEntry(entry)
    Close fileNum
    
    On Error GoTo 0
End Sub

Private Function FormatLogEntry(entry As LogEntry) As String
    Dim output As String
    output = Format(entry.Timestamp, "yyyy-MM-dd HH:mm:ss.000") & " | " & _
             "[" & entry.Level & "] | " & _
             entry.FunctionName & " | " & _
             entry.Message
    
    If entry.ErrorNumber <> 0 Then
        output = output & " | Err#" & entry.ErrorNumber & ": " & entry.ErrorSource
    End If
    
    If entry.Context <> "" Then
        output = output & " | Context: " & entry.Context
    End If
    
    output = output & " | Elapsed: " & entry.ElapsedMs & "ms"
    
    FormatLogEntry = output
End Function

Public Sub ViewLog()
    On Error Resume Next
    Shell "notepad.exe " & LOG_FILE_PATH
    On Error GoTo 0
End Sub

'================================================================================
' MELHORIA #2: CACHE DE OBJETOS WORD
'================================================================================

Private Type CachedParagraph
    Index As Long
    Text As String
    TextNormalized As String
    InlineShapesCount As Long
    HasVisualContent As Boolean
    IsBlank As Boolean
    WordCount As Long
    CacheTime As Single
End Type

Private paraCache As Collection
Private cacheTimestamp As Single
Private cacheValid As Boolean

Private Function InitializeParagraphCache(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Set paraCache = New Collection
    cacheTimestamp = Timer
    cacheValid = False
    
    Dim para As Paragraph
    Dim cached As CachedParagraph
    Dim i As Long
    Dim cacheCount As Long
    
    For i = 1 To doc.Paragraphs.count
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextCachePara
        End If
        On Error GoTo ErrorHandler
        
        If Not para Is Nothing Then
            With cached
                .Index = i
                .Text = para.Range.text
                .TextNormalized = NormalizeForMatching(.Text)
                .InlineShapesCount = para.Range.InlineShapes.count
                .HasVisualContent = SafeHasVisualContent(para)
                .IsBlank = IsParagraphEffectivelyBlank(para)
                .WordCount = CountWordsForStamp(.Text)
                .CacheTime = Timer
            End With
            
            paraCache.Add cached
            cacheCount = cacheCount + 1
        End If
        
NextCachePara:
    Next i
    
    cacheValid = True
    LogEvent "InitializeParagraphCache", "INFO", "Cached " & cacheCount & " paragraphs", , "Cache initialization"
    InitializeParagraphCache = True
    Exit Function
    
ErrorHandler:
    LogEvent "InitializeParagraphCache", "ERROR", Err.Description, Err.Number, "Failed to cache paragraphs"
    InitializeParagraphCache = False
End Function

Private Function GetCachedParagraph(index As Long) As CachedParagraph
    On Error Resume Next
    
    If cacheValid And index > 0 And index <= paraCache.count Then
        GetCachedParagraph = paraCache.Item(index)
    End If
End Function

Private Sub InvalidateParagraphCache()
    On Error Resume Next
    Set paraCache = Nothing
    cacheValid = False
    cacheTimestamp = 0
    On Error GoTo 0
End Sub

'================================================================================
' MELHORIA #3: VALIDAÇÃO AVANÇADA DE REGEX COM CONFIDENCE SCORING
'================================================================================

Private Type SensitiveDataPattern
    Name As String
    Pattern As String
    MinConfidence As Long
    ContextKeywords As String
    FalsePositivePatterns As String
End Type

Private Function GetSensitivePatterns() As Collection
    Dim patterns As Collection
    Set patterns = New Collection
    
    Dim cpfPattern As SensitiveDataPattern
    With cpfPattern
        .Name = "CPF"
        .Pattern = "(\d{3}\.\d{3}\.\d{3}-\d{2}|\d{11})"
        .MinConfidence = 85
        .ContextKeywords = "CPF,CADASTRO,PESSOA,FÍSICA,CONTRIBUINTE"
        .FalsePositivePatterns = "000.000.000-00,111.111.111-11"
    End With
    patterns.Add cpfPattern
    
    Dim rgPattern As SensitiveDataPattern
    With rgPattern
        .Name = "RG"
        .Pattern = "(\d{1,2}\.\d{3}\.\d{3}-?\d{1}|\d{7,8}-?\d{1})"
        .MinConfidence = 75
        .ContextKeywords = "RG,IDENTIDADE,REGISTRO,GERAL"
        .FalsePositivePatterns = ""
    End With
    patterns.Add rgPattern
    
    Set GetSensitivePatterns = patterns
End Function

Private Function CalculateSensitiveDataConfidence(matchText As String, pattern As SensitiveDataPattern, contextText As String) As Long
    Dim confidence As Long
    confidence = pattern.MinConfidence
    
    ' Boost confidence if context keywords found
    Dim keywords As Variant
    keywords = Split(pattern.ContextKeywords, ",")
    
    Dim i As Long
    For i = LBound(keywords) To UBound(keywords)
        If InStr(1, contextText, Trim(keywords(i)), vbTextCompare) > 0 Then
            confidence = confidence + 10
        End If
    Next i
    
    ' Reduce confidence if matches known false positive pattern
    If InStr(pattern.FalsePositivePatterns, matchText) > 0 Then
        confidence = 0
    End If
    
    ' Cap at 100
    If confidence > 100 Then confidence = 100
    
    CalculateSensitiveDataConfidence = confidence
End Function

'================================================================================
' MELHORIA #4: BARRA DE PROGRESSO VISUAL
'================================================================================

Private Type ProgressTracker
    TotalItems As Long
    ProcessedItems As Long
    CurrentPhase As String
    StartTime As Single
    LastUpdateTime As Single
    EstimatedTimeRemainingSec As Long
End Type

Private currentProgress As ProgressTracker

Private Function InitializeProgress(totalItems As Long, phase As String) As Boolean
    With currentProgress
        .TotalItems = totalItems
        .ProcessedItems = 0
        .CurrentPhase = phase
        .StartTime = Timer
        .LastUpdateTime = Timer
        .EstimatedTimeRemainingSec = 0
    End With
    
    UpdateProgressBar 0
    InitializeProgress = True
End Function

Private Sub UpdateProgressBar(Optional incrementBy As Long = 1)
    On Error Resume Next
    
    With currentProgress
        If incrementBy > 0 Then
            .ProcessedItems = .ProcessedItems + incrementBy
        End If
        
        ' Update UI only every 0.5 seconds to avoid excessive calls
        If (Timer - .LastUpdateTime) < 0.5 And .ProcessedItems < .TotalItems Then
            Exit Sub
        End If
        
        .LastUpdateTime = Timer
        
        ' Calculate estimate
        Dim elapsedSec As Long
        elapsedSec = CLng(Timer - .StartTime)
        
        If .ProcessedItems > 0 And elapsedSec > 0 Then
            Dim ratePerSec As Double
            ratePerSec = .ProcessedItems / elapsedSec
            .EstimatedTimeRemainingSec = CLng((.TotalItems - .ProcessedItems) / ratePerSec)
        End If
        
        ' Format status bar message
        Dim percentComplete As Long
        percentComplete = IIf(.TotalItems > 0, CLng((.ProcessedItems / .TotalItems) * 100), 0)
        
        Dim statusMsg As String
        statusMsg = "CHAINSAW: " & .CurrentPhase & " (" & percentComplete & "%) | " & _
                    .ProcessedItems & "/" & .TotalItems & " | " & _
                    "Tempo: " & FormatSeconds(.EstimatedTimeRemainingSec)
        
        Application.StatusBar = statusMsg
    End With
    
    On Error GoTo 0
End Sub

Private Function FormatSeconds(seconds As Long) As String
    Dim mins As Long, secs As Long
    mins = seconds \ 60
    secs = seconds Mod 60
    
    If mins > 60 Then
        FormatSeconds = CLng(mins / 60) & "h " & (mins Mod 60) & "m"
    ElseIf mins > 0 Then
        FormatSeconds = mins & "m " & secs & "s"
    Else
        FormatSeconds = secs & "s"
    End If
End Function

'================================================================================
' MELHORIA #5: GESTÃO INTELIGENTE DE MEMÓRIA
'================================================================================

Private Function CleanupMemory() As Boolean
    On Error GoTo ErrorHandler
    
    ' Clear paragraph cache
    Call InvalidateParagraphCache()
    
    LogEvent "CleanupMemory", "INFO", "Memory cleanup completed", , "Post-processing"
    CleanupMemory = True
    Exit Function
    
ErrorHandler:
    LogEvent "CleanupMemory", "ERROR", Err.Description, Err.Number, "Memory cleanup failed"
    CleanupMemory = False
End Function

'================================================================================
' MELHORIA #6: VALIDAÇÃO DE INTEGRIDADE PÓS-PROCESSAMENTO
'================================================================================

Private Type ValidationResult
    IsValid As Boolean
    ErrorCount As Long
    WarningCount As Long
    ChecksPerformed As Long
    ChecksPassed As Long
End Type

Private Function ValidatePostProcessing(doc As Document) As ValidationResult
    Dim result As ValidationResult
    result.ChecksPerformed = 0
    result.ChecksPassed = 0
    result.ErrorCount = 0
    result.WarningCount = 0
    
    ' Check #1: All paragraphs use standard font
    result.ChecksPerformed = result.ChecksPerformed + 1
    If ValidateAllParagraphsHaveStandardFont(doc) Then
        result.ChecksPassed = result.ChecksPassed + 1
    Else
        result.WarningCount = result.WarningCount + 1
    End If
    
    ' Check #2: Page setup matches specification
    result.ChecksPerformed = result.ChecksPerformed + 1
    If ValidatePageSetupCorrect(doc) Then
        result.ChecksPassed = result.ChecksPassed + 1
    Else
        result.ErrorCount = result.ErrorCount + 1
    End If
    
    ' Check #3: No spacing violations
    result.ChecksPerformed = result.ChecksPerformed + 1
    If ValidateNoExcessiveSpacing(doc) Then
        result.ChecksPassed = result.ChecksPassed + 1
    Else
        result.WarningCount = result.WarningCount + 1
    End If
    
    ' Determine overall validity
    result.IsValid = (result.ErrorCount = 0)
    
    LogEvent "ValidatePostProcessing", "INFO", _
             "Validation: " & result.ChecksPassed & "/" & result.ChecksPerformed & " checks passed", , _
             "Errors: " & result.ErrorCount & ", Warnings: " & result.WarningCount
    
    ValidatePostProcessing = result
End Function

Private Function ValidateAllParagraphsHaveStandardFont(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim i As Long
    Dim violationCount As Long
    
    For i = 1 To doc.Paragraphs.count
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextValPara
        End If
        On Error GoTo ErrorHandler
        
        If Not para Is Nothing Then
            If para.Range.Font.name <> STANDARD_FONT Then
                violationCount = violationCount + 1
                If violationCount > 5 Then Exit For
            End If
        End If
        
NextValPara:
    Next i
    
    ValidateAllParagraphsHaveStandardFont = (violationCount = 0)
    Exit Function
    
ErrorHandler:
    ValidateAllParagraphsHaveStandardFont = False
End Function

Private Function ValidatePageSetupCorrect(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim tolerance As Double
    tolerance = 0.2 ' cm tolerance
    
    With doc.PageSetup
        Dim topOk As Boolean, bottomOk As Boolean, leftOk As Boolean, rightOk As Boolean
        
        topOk = Abs(CmFromPoints(.TopMargin) - TOP_MARGIN_CM) < tolerance
        bottomOk = Abs(CmFromPoints(.BottomMargin) - BOTTOM_MARGIN_CM) < tolerance
        leftOk = Abs(CmFromPoints(.LeftMargin) - LEFT_MARGIN_CM) < tolerance
        rightOk = Abs(CmFromPoints(.RightMargin) - RIGHT_MARGIN_CM) < tolerance
        
        ValidatePageSetupCorrect = topOk And bottomOk And leftOk And rightOk
    End With
    
    Exit Function
    
ErrorHandler:
    ValidatePageSetupCorrect = False
End Function

Private Function ValidateNoExcessiveSpacing(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim i As Long
    Dim blankCount As Long
    
    blankCount = 0
    For i = 1 To doc.Paragraphs.count
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextBlankCheck
        End If
        On Error GoTo ErrorHandler
        
        If Not para Is Nothing Then
            If Trim(para.Range.text) = "" Then
                blankCount = blankCount + 1
                If blankCount > 3 Then
                    ValidateNoExcessiveSpacing = False
                    Exit Function
                End If
            Else
                blankCount = 0
            End If
        End If
        
NextBlankCheck:
    Next i
    
    ValidateNoExcessiveSpacing = True
    Exit Function
    
ErrorHandler:
    ValidateNoExcessiveSpacing = False
End Function

'================================================================================
' MELHORIA #7: OTIMIZAÇÃO DO ALGORITMO DE DETECÇÃO DE CARIMBO (2-PASS)
'================================================================================

Private Function FindSessionStampParagraphOptimized(doc As Document) As Paragraph
    On Error GoTo ErrorHandler
    
    If doc Is Nothing Then
        Set FindSessionStampParagraphOptimized = Nothing
        Exit Function
    End If
    
    Dim para As Paragraph
    Dim i As Long
    Dim searchLimit As Long
    
    ' Strategy: Stamps typically appear in first 10-20% of document
    searchLimit = CLng(doc.Paragraphs.count * 0.2)
    If searchLimit > 1000 Then searchLimit = 1000
    If searchLimit < 50 Then searchLimit = doc.Paragraphs.count
    
    ' Two-pass strategy:
    ' Pass 1: Look only for centered paragraphs (faster)
    For i = 1 To searchLimit
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextPassOne
        End If
        On Error GoTo ErrorHandler
        
        If Not para Is Nothing Then
            ' Center-aligned = likely candidate
            If para.alignment = wdAlignParagraphCenter Then
                If Not IsParagraphEffectivelyBlank(para) Then
                    Dim rawText As String
                    rawText = ParagraphTextWithoutBreaks(para)
                    
                    ' Quick length check
                    If CountWordsForStamp(rawText) <= MAX_SESSION_STAMP_WORDS Then
                        If IsLikelySessionStamp(NormalizeForMatching(rawText), para.Range.text) Then
                            If HasBlankPadding(para) Then
                                Set FindSessionStampParagraphOptimized = para
                                LogEvent "FindSessionStampParagraphOptimized", "INFO", "Found stamp at paragraph " & i & " (Pass 1)", , "2-pass optimization"
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
NextPassOne:
    Next i
    
    ' Pass 2: If not found, broaden search (slower fallback)
    LogEvent "FindSessionStampParagraphOptimized", "WARNING", "Pass 1 failed, attempting Pass 2 full scan", , "Fallback search"
    
    For i = searchLimit + 1 To doc.Paragraphs.count
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextPassTwo
        End If
        On Error GoTo ErrorHandler
        
        If Not para Is Nothing Then
            If Not IsParagraphEffectivelyBlank(para) Then
                Dim rawTextPass2 As String
                rawTextPass2 = ParagraphTextWithoutBreaks(para)
                
                If CountWordsForStamp(rawTextPass2) <= MAX_SESSION_STAMP_WORDS Then
                    If IsLikelySessionStamp(NormalizeForMatching(rawTextPass2), para.Range.text) Then
                        If HasBlankPadding(para) Then
                            Set FindSessionStampParagraphOptimized = para
                            LogEvent "FindSessionStampParagraphOptimized", "INFO", "Found stamp at paragraph " & i & " (Pass 2)", , "Fallback search"
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        
NextPassTwo:
    Next i
    
    Set FindSessionStampParagraphOptimized = Nothing
    LogEvent "FindSessionStampParagraphOptimized", "WARNING", "Session stamp not found after 2 passes", , "Document may lack proper structure"
    Exit Function
    
ErrorHandler:
    LogEvent "FindSessionStampParagraphOptimized", "ERROR", Err.Description, Err.Number, "Stamp detection failed"
    Set FindSessionStampParagraphOptimized = Nothing
End Function

'================================================================================
' MELHORIA #8: TRATAMENTO ROBUSTO DE ERROS COM CONTEXTO
'================================================================================

Private Type ErrorContext
    FunctionName As String
    ErrorNumber As Long
    ErrorDescription As String
    ErrorSource As String
    DocumentPath As String
    CurrentOperation As String
    RetryCount As Long
    MaxRetries As Long
End Type

Private Function HandleErrorWithContext(context As ErrorContext) As Boolean
    On Error GoTo ErrorHandler
    
    ' Log detailed error
    LogEvent context.FunctionName, "ERROR", _
             "Err#" & context.ErrorNumber & ": " & context.ErrorDescription, _
             context.ErrorNumber, _
             "Operation: " & context.CurrentOperation & " | Doc: " & context.DocumentPath
    
    ' Attempt automatic recovery based on error type
    Select Case context.ErrorNumber
        Case 11 ' Division by zero
            LogEvent context.FunctionName, "RECOVERY", "Attempted recovery from division by zero", context.ErrorNumber
            HandleErrorWithContext = True
            
        Case 429 ' ActiveX object error
            LogEvent context.FunctionName, "CRITICAL", "Word ActiveX object failed", context.ErrorNumber
            HandleErrorWithContext = False
            
        Case 4605 ' Document protection
            LogEvent context.FunctionName, "WARNING", "Document protection prevented operation", context.ErrorNumber
            HandleErrorWithContext = False
            
        Case Else
            LogEvent context.FunctionName, "ERROR", "Unhandled error - retry scheduled", context.ErrorNumber
            ' Retry logic
            If context.RetryCount < context.MaxRetries Then
                context.RetryCount = context.RetryCount + 1
                HandleErrorWithContext = True
            Else
                HandleErrorWithContext = False
            End If
    End Select
    
    Exit Function
    
ErrorHandler:
    HandleErrorWithContext = False
End Function

'================================================================================
' MELHORIA #9: CONFIGURAÇÃO EXTERNALIZÁVEL
'================================================================================

Private Type ChainsawConfig
    StandardFont As String
    StandardFontSize As Long
    TopMarginCm As Double
    BottomMarginCm As Double
    LeftMarginCm As Double
    RightMarginCm As Double
    EnableLogging As Boolean
    MaxSessionStampWords As Long
    SensitiveDataMinConfidence As Long
    EnableProgressBar As Boolean
End Type

Private configPath As String

Private Function LoadConfiguration() As ChainsawConfig
    On Error GoTo ErrorHandler
    
    configPath = ThisDocument.Path & "\chainsaw.config"
    
    Dim config As ChainsawConfig
    
    ' Default values
    With config
        .StandardFont = "Arial"
        .StandardFontSize = 12
        .TopMarginCm = 4.6
        .BottomMarginCm = 2
        .LeftMarginCm = 3
        .RightMarginCm = 3
        .EnableLogging = True
        .MaxSessionStampWords = 17
        .SensitiveDataMinConfidence = 80
        .EnableProgressBar = True
    End With
    
    ' Try to load from file
    If Dir(configPath) <> "" Then
        Dim fso As Object, configFile As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set configFile = fso.OpenTextFile(configPath, 1) ' ForReading
        
        ' Parse key=value format
        Dim line As String
        While Not configFile.AtEndOfStream
            line = configFile.ReadLine
            If InStr(line, "=") > 0 And Left(line, 1) <> "#" Then
                Dim parts As Variant
                parts = Split(line, "=")
                If UBound(parts) >= 1 Then
                    Dim key As String, value As String
                    key = Trim(parts(0))
                    value = Trim(parts(1))
                    
                    Select Case LCase(key)
                        Case "standardfont"
                            config.StandardFont = value
                        Case "standardfontsize"
                            config.StandardFontSize = CLng(value)
                        Case "topmargin"
                            config.TopMarginCm = CDbl(value)
                        Case "enablelogging"
                            config.EnableLogging = (LCase(value) = "true")
                    End Select
                End If
            End If
        Wend
        configFile.Close
        
        LogEvent "LoadConfiguration", "INFO", "Configuration loaded from " & configPath, , "Settings loaded"
    Else
        Call SaveConfiguration(config)
    End If
    
    Set fso = Nothing
    Set configFile = Nothing
    LoadConfiguration = config
    Exit Function
    
ErrorHandler:
    LogEvent "LoadConfiguration", "ERROR", Err.Description, Err.Number, "Config load failed - using defaults"
    LoadConfiguration = config
End Function

Private Sub SaveConfiguration(config As ChainsawConfig)
    On Error Resume Next
    
    Dim fso As Object, configFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set configFile = fso.CreateTextFile(configPath, True) ' ForWriting, overwrite
    
    With config
        configFile.WriteLine "# CHAINSAW PROPOSITURAS - CONFIGURATION FILE"
        configFile.WriteLine "# Edit this file to customize Chainsaw behavior"
        configFile.WriteLine ""
        configFile.WriteLine "StandardFont=" & .StandardFont
        configFile.WriteLine "StandardFontSize=" & .StandardFontSize
        configFile.WriteLine "TopMargin=" & .TopMarginCm
        configFile.WriteLine "BottomMargin=" & .BottomMarginCm
        configFile.WriteLine "LeftMargin=" & .LeftMarginCm
        configFile.WriteLine "RightMargin=" & .RightMarginCm
        configFile.WriteLine "EnableLogging=" & IIf(.EnableLogging, "true", "false")
        configFile.WriteLine "MaxSessionStampWords=" & .MaxSessionStampWords
        configFile.WriteLine "EnableProgressBar=" & IIf(.EnableProgressBar, "true", "false")
    End With
    
    configFile.Close
    Set configFile = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Sub

'================================================================================
' SOLICITAÇÃO #1: REMOÇÃO DE ESPAÇOS ANTES E DEPOIS DOS PARÁGRAFOS
'================================================================================

Private Function RemoveParagraphSpacing(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim i As Long
    Dim removedCount As Long
    
    If doc Is Nothing Then
        RemoveParagraphSpacing = False
        Exit Function
    End If
    
    ' Process all paragraphs (backwards to avoid index issues)
    For i = doc.Paragraphs.count To 1 Step -1
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParaSpacing
        End If
        On Error GoTo ErrorHandler
        
        If Not para Is Nothing Then
            ' Skip paragraphs after session stamp (PROTECTION ZONE)
            If Not IsAfterSessionStamp(para, ParagraphStampLocation) Then
                ' Remove before and after spacing for paragraphs BEFORE stamp
                With para.Format
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                End With
                removedCount = removedCount + 1
            End If
        End If
        
NextParaSpacing:
    Next i
    
    RemoveParagraphSpacing = True
    Exit Function
    
ErrorHandler:
    RemoveParagraphSpacing = False
End Function

'================================================================================
' SOLICITAÇÃO #2: PROTEÇÃO DA ZONA PÓS-CARIMBO
'================================================================================

Private Function IsAfterSessionStamp(para As Paragraph, stampPara As Paragraph) As Boolean
    On Error GoTo ErrorHandler
    
    IsAfterSessionStamp = False
    
    If para Is Nothing Or stampPara Is Nothing Then
        Exit Function
    End If
    
    Dim paraIndex As Long
    Dim stampIndex As Long
    
    ' Get paragraph indices safely
    On Error Resume Next
    paraIndex = para.Range.ParagraphNumber
    stampIndex = stampPara.Range.ParagraphNumber
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Para is after stamp if its index is greater
    If paraIndex > stampIndex Then
        IsAfterSessionStamp = True
    End If
    
    Exit Function
    
ErrorHandler:
    IsAfterSessionStamp = False
End Function

'================================================================================
' SOLICITAÇÃO #3: JUSTIFICATIVA EM NEGRITO
'================================================================================

Private Function FormatJustificativaHeading(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim i As Long
    Dim paraText As String
    Dim formattedCount As Long
    
    If doc Is Nothing Then
        FormatJustificativaHeading = False
        Exit Function
    End If
    
    ' Process all paragraphs
    For i = 1 To doc.Paragraphs.count
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextJustPara
        End If
        On Error GoTo ErrorHandler
        
        If Not para Is Nothing Then
            ' Get paragraph text (normalized)
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            
            ' Check if paragraph contains ONLY "Justificativa:" (case-insensitive)
            If LCase(paraText) = "justificativa:" Or LCase(paraText) = "justificativa" Then
                
                ' Safety check: skip if paragraph has visual content (images, shapes)
                If Not SafeHasVisualContent(para) Then
                    With para.Range.Font
                        .Bold = True
                    End With
                    formattedCount = formattedCount + 1
                End If
            End If
        End If
        
NextJustPara:
    Next i
    
    FormatJustificativaHeading = True
    Exit Function
    
ErrorHandler:
    FormatJustificativaHeading = False
End Function

'================================================================================
' MELHORIA #10: INTEGRACIÓN ACTUALIZADA EN ENTRY POINT
'================================================================================

Public Sub StandardizeDocumentMain()
    On Error GoTo CriticalErrorHandler
    
    ' ========================================
    ' INITIALIZATION AND CONFIG LOAD
    ' ========================================
    
    processingStartTime = Timer
    formattingCancelled = False
    
    ' Initialize logging system (MELHORIA #1)
    If Not InitializeLogging() Then
        ' Continue without logging
    End If
    
    LogEvent "StandardizeDocumentMain", "INFO", "Document standardization started", , "v1.0.0-Beta3"
    
    ' Load configuration (MELHORIA #9)
    Dim appConfig As ChainsawConfig
    appConfig = LoadConfiguration()
    
    ' ========================================
    ' PRELIMINARY VALIDATIONS
    ' ========================================
    
    ' Word version validation (always on)
    If Not CheckWordVersion() Then
        Application.StatusBar = "Error: Word version not supported (minimum: Word " & MIN_WORD_VERSION & ")"
        Dim verMsg As String
        verMsg = ReplacePlaceholders(MSG_ERR_VERSION, _
                    "MIN", CStr(MIN_WORD_VERSION), _
                    "CUR", CStr(Application.version))
        MsgBox NormalizeForUI(verMsg), vbCritical, NormalizeForUI(TITLE_VERSION_ERROR)
        Exit Sub
    End If
        
    ' Active document validation
    Dim doc As Document
    Set doc = Nothing
    
    On Error Resume Next
    Set doc = ActiveDocument
    If doc Is Nothing Then
        On Error GoTo CriticalErrorHandler
        Application.StatusBar = "Error: No document is accessible"
        MsgBox NormalizeForUI(MSG_NO_DOCUMENT), vbExclamation, NormalizeForUI(TITLE_DOC_NOT_FOUND)
        Exit Sub
    End If
    On Error GoTo CriticalErrorHandler
    
    ' Document integrity validation (always on)
    If Not ValidateDocumentIntegrity(doc) Then GoTo CleanUp
    
    ' Ensure the document is editable
    If Not EnsureDocumentEditable(doc) Then
        Application.StatusBar = "Document is not editable - operation cancelled"
        GoTo CleanUp
    End If
    
    ' ========================================
    ' PERFORMANCE OPTIMIZATION INITIALIZATION
    ' ========================================
    
    If Not InitializePerformanceOptimization() Then
    End If
        
    ' Configure undo group
    StartUndoGroup "Document Standardization - " & doc.name
    
    ' Configure application state
    If Not SetAppState(False, "Formatting document...") Then
    End If
    
    ' Preliminary checks
    If Not PreviousChecking(doc) Then
        GoTo CleanUp
    End If
    
    ' Mandatory save for unsaved documents
    If doc.Path = "" Then
        If Not SaveDocumentFirst(doc) Then
            Application.StatusBar = "Operation cancelled: document needs to be saved"
            GoTo CleanUp
        End If
    End If
    
    ' Initialize paragraph cache (MELHORIA #2)
    If Not InitializeParagraphCache(doc) Then
        LogEvent "StandardizeDocumentMain", "WARNING", "Failed to initialize paragraph cache", , "Continuing without cache"
    End If
    
    ' Initialize progress tracking (MELHORIA #4)
    If appConfig.EnableProgressBar Then
        Call InitializeProgress(doc.Paragraphs.count, "Formatando documento...")
    End If
    
    Application.StatusBar = "Processing document structure..."

    If Not PreviousFormatting(doc) Then
        GoTo CleanUp
    End If

    If formattingCancelled Then
        GoTo CleanUp
    End If
    
    ' Post-processing validation (MELHORIA #6)
    Application.StatusBar = "Validating formatted document..."
    Dim valResult As ValidationResult
    valResult = ValidatePostProcessing(doc)
    
    If Not valResult.IsValid Then
        LogEvent "StandardizeDocumentMain", "ERROR", "Post-processing validation failed", , _
                 "Errors: " & valResult.ErrorCount & ", Warnings: " & valResult.WarningCount
    End If

    Application.StatusBar = "Document standardized successfully!"
    LogEvent "StandardizeDocumentMain", "INFO", "Document standardization completed successfully", , _
             "Validation: " & valResult.ChecksPassed & "/" & valResult.ChecksPerformed & " checks passed"

CleanUp:
    ' Memory cleanup (MELHORIA #5)
    Call CleanupMemory()
    
    ' Restore performance settings
    If Not RestorePerformanceSettings() Then
    End If
  
    If Not SetAppState(True, "Document standardized successfully!") Then
    End If
    
    ' Final status
    Dim elapsedSec As Long
    elapsedSec = ElapsedSeconds()
    Application.StatusBar = "Chainsaw: concluído em " & elapsedSec & "s"
    
    LogEvent "StandardizeDocumentMain", "INFO", "Processing completed in " & elapsedSec & " seconds", , "Total execution time"
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "CRITICAL ERROR #" & Err.Number & ": " & Err.Description & _
              " in " & Err.Source & " (Line: " & Erl & ")"
    
    LogEvent "StandardizeDocumentMain", "CRITICAL", Err.Description, Err.Number, "Critical failure in main entry point"
    Application.StatusBar = "Critical error during processing"
    
    EmergencyRecovery
End Sub

'================================================================================
' UPDATED PREVIOUS FORMATTING WITH ALL SOLICITAÇÕES
'================================================================================

Private Function PreviousFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ' Apply page setup (always on)
    Call ApplyPageSetup(doc)
    
    ' Clean document structure (remove blank lines above the first text and leading spaces)
    Call CleanDocumentStructure(doc)

    ' Validate proposition type (informational)
    Call ValidatePropositionType(doc)

    ' Validate content consistency (may cancel if user chooses so)
    If Not ValidateContentConsistency(doc) Then
        PreviousFormatting = False
        Exit Function
    End If
    
    ' Locate session stamp with fuzzy matching (stores result in module-level variable)
    Set ParagraphStampLocation = FindSessionStampParagraphOptimized(doc)
    
    ' Title formatting (uppercased, bold, underlined, centered)
    Call FormatDocumentTitle(doc)

    ' Apply standard font (always on)
    Call ApplyStdFont(doc)
    
    ' Apply standard paragraphs (always on)
    Call ApplyStdParagraphs(doc)

    ' Format first and second paragraphs
    Call FormatFirstParagraph(doc)
    Call FormatSecondParagraph(doc)
    
    ' Apply CONSIDERANDO uppercase/bold at paragraph start
    Call FormatConsiderandoParagraphs(doc)
    
    ' Apply text replacements (always on)
    Call ApplyTextReplacements(doc)
    
    ' Apply specific paragraph replacements (always on)
    Call ApplySpecificParagraphReplacements(doc)
    
    ' Normalize numbered paragraphs
    Call FormatNumberedParagraphs(doc)
    
    ' Format "Justificativa:" heading as bold (SOLICITAÇÃO #3)
    Call FormatJustificativaHeading(doc)
    
    ' Justificativa/Anexo formatting
    Call FormatJustificativaAnexoParagraphs(doc)

    ' Hyphenation and watermark
    Call RemoveWatermark(doc)

    ' Insert header image (always enabled)
    InsertHeaderstamp doc
    
    ' Insert page numbers in footer (restored feature)
    Call InsertFooterstamp(doc)
    
    ' Remove spacing before and after paragraphs (SOLICITAÇÃO #1)
    Call RemoveParagraphSpacing(doc)
    
    ' Final spacing and separation controls (with PROTECTION ZONE support)
    Call CleanMultipleSpaces(doc)
    Call LimitSequentialEmptyLines(doc)
    Call EnsureParagraphSeparation(doc)
    Call EnsureSecondParagraphBlankLines(doc)
    Call FormatJustificativaAnexoParagraphs(doc)
    
    ' Replace session stamp with standardized format (final format routine)
    Call ReplaceSessionStampParagraph()
    
    ' Configure view (keeps user zoom)
    Call ConfigureDocumentView(doc)
    
    PreviousFormatting = True
    Exit Function

ErrorHandler:
    PreviousFormatting = False
End Function

'================================================================================
' UPDATED FUNCTIONS WITH PROTECTION ZONE SUPPORT (SOLICITAÇÃO #2)
'================================================================================

Private Function CleanMultipleSpaces(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim i As Long
    Dim cleanText As String
    
    If doc Is Nothing Then
        CleanMultipleSpaces = False
        Exit Function
    End If
    
    For i = doc.Paragraphs.count To 1 Step -1
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextCleanPara
        End If
        On Error GoTo ErrorHandler
        
        If Not para Is Nothing Then
            ' PROTECTION: Skip paragraphs after session stamp
            If Not IsAfterSessionStamp(para, ParagraphStampLocation) Then
                cleanText = para.Range.text
                
                ' Only clean if has multiple spaces
                If InStr(cleanText, "  ") > 0 Then
                    Do While InStr(cleanText, "  ") > 0
                        cleanText = Replace(cleanText, "  ", " ")
                    Loop
                    
                    ' Apply only if no visual content
                    If Not SafeHasVisualContent(para) Then
                        para.Range.text = cleanText
                    End If
                End If
            End If
        End If
        
NextCleanPara:
    Next i
    
    CleanMultipleSpaces = True
    Exit Function
    
ErrorHandler:
    CleanMultipleSpaces = False
End Function

Private Function LimitSequentialEmptyLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim blankCount As Long
    Dim i As Long
    Dim paraText As String
    
    If doc Is Nothing Then
        LimitSequentialEmptyLines = False
        Exit Function
    End If
    
    blankCount = 0
    i = 1
    
    Do While i <= doc.Paragraphs.count
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            blankCount = 0
            i = i + 1
            GoTo NextLimitPara
        End If
        On Error GoTo ErrorHandler
        
        If Not para Is Nothing Then
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            
            ' Check if paragraph is blank
            If paraText = "" And Not SafeHasVisualContent(para) Then
                blankCount = blankCount + 1
                
                ' PROTECTION: Do not remove blank lines after session stamp
                If IsAfterSessionStamp(para, ParagraphStampLocation) Then
                    ' Skip removal - preserve structure after stamp
                    i = i + 1
                    GoTo NextLimitPara
                End If
                
                ' Remove if more than 2 consecutive blanks (before stamp)
                If blankCount > 2 Then
                    para.Range.Delete
                    ' Do NOT increment i, check same position again
                    GoTo NextLimitPara
                End If
            Else
                blankCount = 0
            End If
        End If
        
NextLimitPara:
        i = i + 1
    Loop
    
    LimitSequentialEmptyLines = True
    Exit Function
    
ErrorHandler:
    LimitSequentialEmptyLines = False
End Function

Private Function EnsureParagraphSeparation(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim i As Long
    
    If doc Is Nothing Then
        EnsureParagraphSeparation = False
        Exit Function
    End If
    
    ' Process paragraphs in FORWARD order (to avoid index shifts)
    For i = 1 To doc.Paragraphs.count
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextSepPara
        End If
        On Error GoTo ErrorHandler
        
        If Not para Is Nothing Then
            ' PROTECTION: Skip paragraphs after session stamp
            If Not IsAfterSessionStamp(para, ParagraphStampLocation) Then
                ' Original logic: ensure separation between paragraphs
                If para.Format.spacing After < 6 Then
                    para.Format.spacing After = 6
                End If
            End If
        End If
        
NextSepPara:
    Next i
    
    EnsureParagraphSeparation = True
    Exit Function
    
ErrorHandler:
    EnsureParagraphSeparation = False
End Function

Private Function EnsureSecondParagraphBlankLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim secondParaIndex As Long
    
    If doc Is Nothing Then
        EnsureSecondParagraphBlankLines = False
        Exit Function
    End If
    
    ' Find second paragraph with content
    actualParaIndex = 0
    secondParaIndex = 0
    
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        If paraText <> "" Or SafeHasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1
            
            If actualParaIndex = 2 Then
                secondParaIndex = i
                Exit For
            End If
        End If
        
        If i > 10 Then Exit For
    Next i
    
    ' Apply formatting only to second paragraph (before stamp zone)
    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(secondParaIndex)
        
        ' PROTECTION: Only process if NOT in protected zone (after stamp)
        If Not IsAfterSessionStamp(para, ParagraphStampLocation) Then
            ' Original formatting logic for 2nd paragraph
            With para.Format
                .leftIndent = CentimetersToPoints(9)
                .firstLineIndent = 0
                .RightIndent = 0
                .alignment = wdAlignParagraphJustify
            End With
        End If
    End If
    
    EnsureSecondParagraphBlankLines = True
    Exit Function
    
ErrorHandler:
    EnsureSecondParagraphBlankLines = False
End Function

' ...existing code...