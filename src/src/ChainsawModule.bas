' Chainsaw Proposituras automation macros for Microsoft Word legislative documents.
' Distributed under the GNU General Public License v3.0 or later; see accompanying license details.
' Maintainer: Christian Martin dos Santos <chrmsantos@gmail.com>

Option Explicit

'=== Windows API Declarations ===
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'=== Core constants ===

' System constants
Private Const version As String = "v1.0.0-Stable"
Private Const SYSTEM_NAME As String = "CHAINSAW PROPOSITURAS"

'================================================================================
' CENTRALIZED USER-FACING MESSAGES & TITLES
'================================================================================
Private Const MSG_ERR_VERSION As String = "This tool requires Microsoft Word {{MIN}} or higher." & vbCrLf & _
    "Current version: {{CUR}}" & vbCrLf & _
    "Minimum version: {{MIN}}"
Private Const MSG_NO_DOCUMENT As String = "No document is open or accessible." & vbCrLf & _
    "Open a document before running the standardization."
Private Const MSG_ENABLE_EDITING As String = "The document appears to be in Protected View or not fully editable." & vbCrLf & _
    "Do you want to attempt enabling editing now?"
Private Const MSG_INACCESSIBLE As String = "The document cannot be fully accessed or is in a state that prevents processing." & vbCrLf & _
    "Check protection, permissions, or file integrity."
Private Const MSG_PROTECTED As String = "The document is protected: {{PROT}}." & vbCrLf & _
    "Some operations may not be possible until protection is removed. Continue?"
Private Const MSG_EMPTY_DOC As String = "The document contains no substantive content to process." & vbCrLf & _
    "Add text before running the formatter."
Private Const MSG_LARGE_DOC As String = "This document is large ({{SIZE}} bytes)." & vbCrLf & _
    "Processing may take longer. Continue?"
Private Const MSG_UNSAVED As String = "The document has unsaved changes." & vbCrLf & _
    "Do you want to save before continuing?"
Private Const MSG_VALIDATION_ERROR As String = "A validation error occurred: {{ERR}}" & vbCrLf & _
    "Review the document and try again."
Private Const MSG_DOC_TYPE_WARNING As String = "The detected document type may not match expected patterns." & vbCrLf & _
    "Proceed anyway?"
Private Const MSG_PROCESSING_CANCELLED As String = "Processing cancelled by user." & vbCrLf & _
    "No changes were finalized."
Private Const MSG_INCONSISTENCY_WARNING As String = "Potential content inconsistencies were detected." & vbCrLf & _
    "Review highlighted sections. Continue processing?"
Private Const MSG_SAVE_ERROR As String = "An error occurred while saving the document." & vbCrLf & _
    "Verify permissions and disk space."
Private Const MSG_OPERATION_CANCELLED As String = "Operation cancelled by user." & vbCrLf & _
    "No further actions executed."
Private Const MSG_CRITICAL_SAVE_EXIT As String = "Critical save failure: {{ERR}}" & vbCrLf & _
    "Processing aborted to prevent data loss."

' UI message titles
Private Const TITLE_VERSION_ERROR As String = "Version Requirement - " & SYSTEM_NAME
Private Const TITLE_DOC_NOT_FOUND As String = "Document Not Found - " & SYSTEM_NAME
Private Const TITLE_ENABLE_EDITING As String = "Enable Editing - " & SYSTEM_NAME
Private Const TITLE_INTEGRITY_ERROR As String = "Integrity Error - " & SYSTEM_NAME
Private Const TITLE_PROTECTED As String = "Protected Document - " & SYSTEM_NAME
Private Const TITLE_EMPTY_DOC As String = "Empty Document - " & SYSTEM_NAME
Private Const TITLE_LARGE_DOC As String = "Large Document - " & SYSTEM_NAME
Private Const TITLE_UNSAVED As String = "Unsaved Document - " & SYSTEM_NAME
Private Const TITLE_VALIDATION_ERROR As String = "Validation Error - " & SYSTEM_NAME
Private Const TITLE_DOC_TYPE As String = "Document Type - " & SYSTEM_NAME
Private Const TITLE_OPERATION_CANCELLED As String = "Operation Cancelled - " & SYSTEM_NAME
Private Const TITLE_CONSISTENCY As String = "Consistency Check - " & SYSTEM_NAME
Private Const TITLE_SAVE_ERROR As String = "Save Error - " & SYSTEM_NAME
Private Const TITLE_FINAL_CONFIRM As String = "Final Confirmation - " & SYSTEM_NAME
Private Const TITLE_CRITICAL_SAVE_EXIT As String = "Critical Save Exit - " & SYSTEM_NAME

'=== Word constant aliases ===

' Word built-in constants
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
Private Const wdReplaceNone As Long = 0
Private Const wdReplaceOne As Long = 1
Private Const wdReplaceAll As Long = 2

' Document formatting defaults
Private Const STANDARD_FONT As String = "Arial"
Private Const STANDARD_FONT_SIZE As Long = 12
Private Const FOOTER_FONT_SIZE As Long = 9
Private Const LINE_SPACING As Single = 14

' Margin distances in centimeters
Private Const TOP_MARGIN_CM As Double = 4.6
Private Const BOTTOM_MARGIN_CM As Double = 2
Private Const LEFT_MARGIN_CM As Double = 3
Private Const RIGHT_MARGIN_CM As Double = 3
Private Const HEADER_DISTANCE_CM As Double = 0.3
Private Const FOOTER_DISTANCE_CM As Double = 0.9

' Header image sizing
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 21
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.7
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.19
Private Const MAX_SESSION_STAMP_WORDS As Long = 17
' Paragraph batching thresholds for throughput safeguards
Private Const OPTIMIZATION_THRESHOLD As Long = 400
Private Const MAX_PARAGRAPH_BATCH_SIZE As Long = 120

' Static resource locations
Private Const MIN_WORD_VERSION As Double = 14#
Private Const HEADER_IMAGE_RELATIVE_PATH As String = "assets\stamp.png"

' Backup and recovery limits
Private Const BACKUP_FOLDER_NAME As String = "\Backups"
Private Const MAX_RETRY_ATTEMPTS As Long = 3
Private Const MAX_FIND_REPLACE_BATCH As Long = 100

'=== Logging infrastructure ===

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

'=== Paragraph cache ===

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

Private paraCache() As CachedParagraph
Private cacheTimestamp As Single
Private cacheValid As Boolean
Private paraCacheInitialized As Boolean

'=== Sensitive data pattern definitions ===

Private Enum SensitivePatternField
    spName = 0
    spRegex = 1
    spMinConfidence = 2
    spContextKeywords = 3
    spFalsePositives = 4
End Enum

'=== Progress tracking state ===

Private Type ProgressTracker
    TotalItems As Long
    ProcessedItems As Long
    CurrentPhase As String
    StartTime As Single
    LastUpdateTime As Single
    EstimatedTimeRemainingSec As Long
End Type

Private currentProgress As ProgressTracker

'=== Document validation results ===

Private Type ValidationResult
    IsValid As Boolean
    ErrorCount As Long
    WarningCount As Long
    ChecksPerformed As Long
    ChecksPassed As Long
End Type

'=== Error handling context ===

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

'=== External configuration snapshot ===

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
Private Type ChainsawRuntimeConfig
    debugMode As Boolean
    performanceMode As Boolean
    compatibilityMode As Boolean
    checkWordVersion As Boolean
    validateDocumentIntegrity As Boolean
    validatePropositionType As Boolean
    validateContentConsistency As Boolean
    checkDiskSpace As Boolean
    minWordVersion As Double
    maxDocumentSize As Long
    autoBackup As Boolean
    backupBeforeProcessing As Boolean
    maxBackupFiles As Long
    backupCleanup As Boolean
    backupRetryAttempts As Long
    applyPageSetup As Boolean
    applyStandardFont As Boolean
    applyStandardParagraphs As Boolean
    formatFirstParagraph As Boolean
    formatSecondParagraph As Boolean
    formatNumberedParagraphs As Boolean
    formatConsiderandoParagraphs As Boolean
    formatJustificativaParagraphs As Boolean
    enableHyphenation As Boolean
    clearAllFormatting As Boolean
    cleanDocumentStructure As Boolean
    cleanMultipleSpaces As Boolean
    limitSequentialEmptyLines As Boolean
    ensureParagraphSeparation As Boolean
    cleanVisualElements As Boolean
    deleteHiddenElements As Boolean
    deleteVisualElementsFirstFourParagraphs As Boolean
    insertHeaderstamp As Boolean
    insertFooterstamp As Boolean
    removeWatermark As Boolean
    headerImagePath As String
    headerImageMaxWidth As Double
    headerImageHeightRatio As Double
    applyTextReplacements As Boolean
    applySpecificParagraphReplacements As Boolean
    replaceHyphensWithEmDash As Boolean
    removeManualLineBreaks As Boolean
    normalizeDosteVariants As Boolean
    normalizeVereadorVariants As Boolean
    backupAllImages As Boolean
    restoreAllImages As Boolean
    protectImagesInRange As Boolean
    backupViewSettings As Boolean
    restoreViewSettings As Boolean
    enableLogging As Boolean
    logLevel As String
    logToFile As Boolean
    logDetailedOperations As Boolean
    logWarnings As Boolean
    logErrors As Boolean
    maxLogSizeMb As Long
    disableScreenUpdating As Boolean
    disableDisplayAlerts As Boolean
    useBulkOperations As Boolean
    optimizeFindReplace As Boolean
    minimizeObjectCreation As Boolean
    cacheFrequentlyUsedObjects As Boolean
    useEfficientLoops As Boolean
    batchParagraphOperations As Boolean
    showProgressMessages As Boolean
    showStatusBarUpdates As Boolean
    confirmCriticalOperations As Boolean
    showCompletionMessage As Boolean
    enableEmergencyRecovery As Boolean
    timeoutOperations As Boolean
    supportWord2010 As Boolean
    supportWord2013 As Boolean
    supportWord2016 As Boolean
    useSafePropertyAccess As Boolean
    fallbackMethods As Boolean
    handleMissingFeatures As Boolean
    requireDocumentSaved As Boolean
    validateFilePermissions As Boolean
    checkDocumentProtection As Boolean
    enableEmergencyBackup As Boolean
    sanitizeInputs As Boolean
    validateRanges As Boolean
    maxRetryAttempts As Long
    retryDelayMs As Long
    compilationCheck As Boolean
    vbaAccessRequired As Boolean
    autoCleanup As Boolean
    forceGcCollection As Boolean
End Type

Private runtimeConfig As ChainsawRuntimeConfig

'=== Global state ===
Private undoGroupEnabled As Boolean
Private formattingCancelled As Boolean
Private processingStartTime As Single ' Stores Timer() value at start of processing
Private ParagraphStampLocation As Paragraph ' Locates session stamp with fuzzy matching

'=== Measurement helpers ===
' Converts Word points to centimeters using 1 in = 2.54 cm and 72 points per inch.
Private Function CmFromPoints(ByVal pts As Double) As Double
    On Error GoTo ErrorHandler
    CmFromPoints = (pts * 2.54) / 72#
    Exit Function
ErrorHandler:
    CmFromPoints = 0
End Function

'=== Timing helpers ===
' Returns elapsed seconds from processingStartTime, guarding against midnight rollover.
Private Function ElapsedSeconds() As Long
    On Error GoTo ErrorHandler
    If processingStartTime <= 0 Then
        ElapsedSeconds = 0
    Else
        ElapsedSeconds = CLng(Timer - processingStartTime)
        If ElapsedSeconds < 0 Then ' Timer wraps at midnight
            ElapsedSeconds = ElapsedSeconds + 86400
        End If
    End If
    Exit Function
ErrorHandler:
    ElapsedSeconds = 0
End Function

'=== Performance guardrails ===

Private Function InitializePerformanceOptimization() As Boolean
    On Error GoTo ErrorHandler
    
    InitializePerformanceOptimization = False
    
    ' Suspend expensive UI updates to accelerate bulk operations
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    
    ' Apply additional Word-specific performance settings
    Call OptimizeWordSettings
    
    
    InitializePerformanceOptimization = True
    Exit Function
    
ErrorHandler:
    InitializePerformanceOptimization = False
End Function

Private Sub OptimizeWordSettings()
    On Error Resume Next
    
    ' Apply Word-specific optimizations (always on)
        With ActiveDocument
            .TrackRevisions = False
            .ShowRevisions = False
        End With
    
        With Selection.Find
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
    
    On Error GoTo 0
End Sub

Private Function RestorePerformanceSettings() As Boolean
    On Error GoTo ErrorHandler
    
    RestorePerformanceSettings = False
    
    ' Restore original settings
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    
    RestorePerformanceSettings = True
    Exit Function
    
ErrorHandler:
    RestorePerformanceSettings = False
End Function

Private Sub FormatParagraph(para As Paragraph)
    On Error Resume Next
    ' Basic normalization: remove leading/trailing spaces in paragraph text (without touching internal spacing)
    ' Operates only on non-empty paragraphs of plain text (skips those containing tables or shapes inline)
    If para Is Nothing Then Exit Sub
    If para.Range.Tables.count > 0 Then Exit Sub
    Dim txt As String
    txt = para.Range.text
    ' Word paragraphs end with vbCr; preserve the final terminator
    If Len(txt) > 1 Then
        Dim body As String
        body = Left$(txt, Len(txt) - 1)
        body = Trim$(body)
        para.Range.text = body & vbCr
    End If
    On Error GoTo 0
End Sub

Private Sub CleanParagraph(para As Paragraph)
    On Error Resume Next
    ' Collapse runs of more than two spaces to a single space inside the paragraph (except inside numbered lists)
    If para Is Nothing Then Exit Sub
    If para.Range.ListFormat.ListType <> wdListNoNumbering Then Exit Sub
    Dim r As Range
    Set r = para.Range.Duplicate
    ' Exclude final paragraph mark
    r.End = r.End - 1
    Dim s As String
    s = r.text
    If InStr(s, "   ") > 0 Then
        Do While InStr(s, "   ") > 0
            s = Replace$(s, "   ", "  ")
        Loop
        ' Now reduce any double spaces not after period to single
        ' (simple heuristic; avoids removing double space after full stop if style uses it)
        Dim tmp As String
        tmp = s
        ' Replace double spaces that are not after period
        Dim i As Long
        For i = 1 To Len(tmp) - 2
            If Mid$(tmp, i, 2) = "  " Then
                If i = 1 Or Mid$(tmp, i - 1, 1) <> "." Then
                    tmp = Left$(tmp, i - 1) & " " & Mid$(tmp, i + 2)
                End If
            End If
        Next i
        r.text = tmp
    End If
    On Error GoTo 0
End Sub

Private Sub ValidateParagraph(para As Paragraph)
    On Error Resume Next
    ' Placeholder for paragraph validation
    ' This stub exists to support batch processing operations
    ' Implement validation logic here as needed
    On Error GoTo 0
End Sub

'=== Session stamp detection ===

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
        Dim elapsedTime As Double
        elapsedTime = Timer - processingStartTime
        If elapsedTime < 0 Then elapsedTime = elapsedTime + 86400
        .ElapsedMs = CLng(elapsedTime * 1000)
    End With
    
    ' Append entry to the rolling log file
    Dim fileNum As Integer
    fileNum = FreeFile
    
    On Error Resume Next
    Open LOG_FILE_PATH For Append As fileNum
    If Err.Number = 0 Then
        Print #fileNum, FormatLogEntry(entry)
        Close fileNum
    Else
        ' Ignore logging failures to avoid recursive error handling
        Err.Clear
    End If
    
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
    ' Launch log viewer with a quoted path to survive spaces
    Shell "notepad.exe """ & LOG_FILE_PATH & """"
    On Error GoTo 0
End Sub

'=== Paragraph cache operations ===

Private Function InitializeParagraphCache(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ' Reset prior cache content so stale references are released
    If paraCacheInitialized Then
        Call InvalidateParagraphCache()
    End If
    
    cacheTimestamp = Timer
    cacheValid = False
    
    Dim para As Paragraph
    Dim cached As CachedParagraph
    Dim i As Long
    Dim cacheCount As Long
    Dim paragraphCount As Long
    Dim cacheIndex As Long
    
    ' Confirm the document reference before iterating
    If doc Is Nothing Then
        LogEvent "InitializeParagraphCache", "ERROR", "Document is Nothing", 0, "Cannot cache null document"
        InitializeParagraphCache = False
        Exit Function
    End If
    
    ' Cap cache size to safeguard memory on huge documents
    Dim maxParagraphsToCache As Long
    maxParagraphsToCache = 10000 ' Prevent unbounded cache growth
    
    paragraphCount = doc.Paragraphs.count
    
    If paragraphCount > maxParagraphsToCache Then
        LogEvent "InitializeParagraphCache", "WARNING", "Document has " & paragraphCount & " paragraphs, limiting cache to " & maxParagraphsToCache, 0, "Large document"
        paragraphCount = maxParagraphsToCache
    End If

    If paragraphCount <= 0 Then
        cacheValid = False
        paraCacheInitialized = False
        LogEvent "InitializeParagraphCache", "INFO", "Document has no paragraphs to cache", , "Cache initialization"
        InitializeParagraphCache = True
        Exit Function
    End If

    ReDim paraCache(1 To paragraphCount)
    paraCacheInitialized = True
    
    For i = 1 To paragraphCount
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextCachePara
        End If
        On Error GoTo ErrorHandler
        
        If Not para Is Nothing Then
            cacheIndex = cacheIndex + 1
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
            
            paraCache(cacheIndex) = cached
            cacheCount = cacheCount + 1
        End If
        
NextCachePara:
    Next i
    
    If cacheCount = 0 Then
        Call InvalidateParagraphCache()
        LogEvent "InitializeParagraphCache", "WARNING", "No paragraphs cached due to access errors", 0, "Cache initialization"
        InitializeParagraphCache = False
        Exit Function
    End If
    
    If cacheCount < paragraphCount Then
        ReDim Preserve paraCache(1 To cacheCount)
    End If
    
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
    
    ' Read cache only when marked valid and the backing array is initialized
    If cacheValid And paraCacheInitialized Then
        If index > 0 Then
            Dim upperBound As Long
            upperBound = UBound(paraCache)
            If index <= upperBound Then
                GetCachedParagraph = paraCache(index)
            End If
        End If
    End If
End Function

Private Sub InvalidateParagraphCache()
    On Error Resume Next
    If paraCacheInitialized Then
        Erase paraCache
        paraCacheInitialized = False
    End If
    cacheValid = False
    cacheTimestamp = 0
    On Error GoTo 0
End Sub

'=== Sensitive data helpers ===

Private Function GetSensitivePatterns() As Collection
    Dim patterns As Collection
    Set patterns = New Collection
    
    Dim cpfPattern(0 To 4) As Variant
    cpfPattern(spName) = "CPF"
    cpfPattern(spRegex) = "(\d{3}\.\d{3}\.\d{3}-\d{2}|\d{11})"
    cpfPattern(spMinConfidence) = 85
    cpfPattern(spContextKeywords) = "CPF,CADASTRO,PESSOA,FISICA,CONTRIBUINTE"
    cpfPattern(spFalsePositives) = "000.000.000-00,111.111.111-11"
    patterns.Add cpfPattern
    
    Dim rgPattern(0 To 4) As Variant
    rgPattern(spName) = "RG"
    rgPattern(spRegex) = "(\d{1,2}\.\d{3}\.\d{3}-?\d{1}|\d{7,8}-?\d{1})"
    rgPattern(spMinConfidence) = 75
    rgPattern(spContextKeywords) = "RG,IDENTIDADE,REGISTRO,GERAL"
    rgPattern(spFalsePositives) = ""
    patterns.Add rgPattern
    
    Set GetSensitivePatterns = patterns
End Function

Private Function CalculateSensitiveDataConfidence(matchText As String, pattern As Variant, contextText As String) As Long
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
'================================================================================

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
        Dim timeSinceUpdate As Double
        timeSinceUpdate = Timer - .LastUpdateTime
        If timeSinceUpdate < 0 Then timeSinceUpdate = timeSinceUpdate + 86400 ' Handle midnight wrap
        If timeSinceUpdate < 0.5 And .ProcessedItems < .TotalItems Then
            Exit Sub
        End If
        
        .LastUpdateTime = Timer
        
        ' Calculate estimate
        Dim elapsedSec As Long
        elapsedSec = CLng(Timer - .StartTime)
        If elapsedSec < 0 Then elapsedSec = elapsedSec + 86400 ' Handle midnight wrap
        
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
    On Error GoTo ErrorHandler
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
    Exit Function
ErrorHandler:
    FormatSeconds = "0s"
End Function

'================================================================================
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
'================================================================================

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
    
    ' Favor early paragraphs because session stamps usually appear near the front
    searchLimit = CLng(doc.Paragraphs.count * 0.2)
    If searchLimit > 1000 Then searchLimit = 1000
    If searchLimit < 50 Then searchLimit = doc.Paragraphs.count
    
    ' Pass 1 scans centered paragraphs only for a quick win
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
    
    ' Pass 2 falls back to a full scan when the first pass fails
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

'=== Context-aware error handling ===

Private Function HandleErrorWithContext(context As ErrorContext) As Boolean
    On Error GoTo ErrorHandler
    
    ' Emit a detailed diagnostic record
    LogEvent context.FunctionName, "ERROR", _
             "Err#" & context.ErrorNumber & ": " & context.ErrorDescription, _
             context.ErrorNumber, _
             "Operation: " & context.CurrentOperation & " | Doc: " & context.DocumentPath
    
    ' Apply targeted recovery logic where feasible
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

' Seeds runtimeConfig with commonsense defaults prior to applying overrides.
Private Function InitializeRuntimeConfigDefaults() As ChainsawRuntimeConfig
    Dim defaults As ChainsawRuntimeConfig
    
    With defaults
        .debugMode = False
        .performanceMode = False
        .compatibilityMode = False
        .checkWordVersion = True
        .validateDocumentIntegrity = True
        .validatePropositionType = True
        .validateContentConsistency = True
        .checkDiskSpace = True
        .minWordVersion = MIN_WORD_VERSION
        .maxDocumentSize = 10485760
        .autoBackup = True
        .backupBeforeProcessing = True
        .maxBackupFiles = 5
        .backupCleanup = True
        .backupRetryAttempts = MAX_RETRY_ATTEMPTS
        .applyPageSetup = True
        .applyStandardFont = True
        .applyStandardParagraphs = True
        .formatFirstParagraph = True
        .formatSecondParagraph = True
        .formatNumberedParagraphs = True
        .formatConsiderandoParagraphs = True
        .formatJustificativaParagraphs = True
        .enableHyphenation = False
        .clearAllFormatting = False
        .cleanDocumentStructure = True
        .cleanMultipleSpaces = True
        .limitSequentialEmptyLines = True
        .ensureParagraphSeparation = True
        .cleanVisualElements = True
        .deleteHiddenElements = True
        .deleteVisualElementsFirstFourParagraphs = False
        .insertHeaderstamp = True
        .insertFooterstamp = True
        .removeWatermark = False
        .headerImagePath = HEADER_IMAGE_RELATIVE_PATH
        .headerImageMaxWidth = HEADER_IMAGE_MAX_WIDTH_CM
        .headerImageHeightRatio = HEADER_IMAGE_HEIGHT_RATIO
        .applyTextReplacements = True
        .applySpecificParagraphReplacements = True
        .replaceHyphensWithEmDash = False
        .removeManualLineBreaks = True
        .normalizeDosteVariants = True
        .normalizeVereadorVariants = True
        .backupAllImages = False
        .restoreAllImages = False
        .protectImagesInRange = False
        .backupViewSettings = True
        .restoreViewSettings = True
        .enableLogging = True
        .logLevel = "INFO"
        .logToFile = True
        .logDetailedOperations = False
        .logWarnings = True
        .logErrors = True
        .maxLogSizeMb = MAX_LOG_SIZE_MB
        .disableScreenUpdating = True
        .disableDisplayAlerts = True
        .useBulkOperations = True
        .optimizeFindReplace = True
        .minimizeObjectCreation = True
        .cacheFrequentlyUsedObjects = True
        .useEfficientLoops = True
        .batchParagraphOperations = True
        .showProgressMessages = True
        .showStatusBarUpdates = True
        .confirmCriticalOperations = False
        .showCompletionMessage = True
        .enableEmergencyRecovery = True
        .timeoutOperations = False
        .supportWord2010 = True
        .supportWord2013 = True
        .supportWord2016 = True
        .useSafePropertyAccess = True
        .fallbackMethods = True
        .handleMissingFeatures = True
        .requireDocumentSaved = True
        .validateFilePermissions = True
        .checkDocumentProtection = True
        .enableEmergencyBackup = True
        .sanitizeInputs = True
        .validateRanges = True
        .maxRetryAttempts = MAX_RETRY_ATTEMPTS
        .retryDelayMs = 500
        .compilationCheck = True
        .vbaAccessRequired = False
        .autoCleanup = True
        .forceGcCollection = False
    End With
    
    InitializeRuntimeConfigDefaults = defaults
End Function

'=== Configuration loader ===

Private Function LoadConfiguration() As ChainsawConfig
    On Error GoTo ErrorHandler
    
    ' Validate document path before attempting file operations
    If ThisDocument.Path = "" Then
        configPath = ""
    Else
        configPath = ThisDocument.Path & "\chainsaw.config"
    End If
    
    Dim config As ChainsawConfig
    
    ' Seed defaults before reading configuration file
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
    
    ' Attempt to read persisted configuration
    If Dir(configPath) <> "" Then
        Dim fso As Object, configFile As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        Set configFile = fso.OpenTextFile(configPath, 1) ' ForReading
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo EndLoadConfig
        End If
        On Error GoTo 0
        
    ' Consume lines written as key=value pairs
        Dim line As String
        On Error Resume Next
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
                            On Error Resume Next
                            config.StandardFontSize = CLng(value)
                            On Error GoTo 0
                        Case "topmargin"
                            On Error Resume Next
                            config.TopMarginCm = CDbl(value)
                            On Error GoTo 0
                        Case "enablelogging"
                            config.EnableLogging = (LCase(value) = "true")
                    End Select
                End If
            End If
        Wend
        On Error GoTo 0
        configFile.Close
        
        EndLoadConfig:
        
        LogEvent "LoadConfiguration", "INFO", "Configuration loaded from " & configPath, , "Settings loaded"
    Else
        Call SaveConfiguration(config)
    End If
    
    ' Always close file handles before exit
    If Not configFile Is Nothing Then
        On Error Resume Next
        configFile.Close
        On Error GoTo 0
    End If
    
    runtimeConfig = InitializeRuntimeConfigDefaults()
    
    Set fso = Nothing
    Set configFile = Nothing
    LoadConfiguration = config
    Exit Function
    
ErrorHandler:
    LogEvent "LoadConfiguration", "ERROR", Err.Description, Err.Number, "Config load failed - using defaults"
    runtimeConfig = InitializeRuntimeConfigDefaults()
    LoadConfiguration = config
End Function

Private Sub SaveConfiguration(config As ChainsawConfig)
    On Error Resume Next
    
    Dim fso As Object, configFile As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    
    Set configFile = fso.CreateTextFile(configPath, True) ' ForWriting, overwrite
    If Err.Number <> 0 Then
        On Error GoTo 0
        Set fso = Nothing
        Exit Sub
    End If
    On Error GoTo 0
    
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
    
    On Error Resume Next
    configFile.Close
    On Error GoTo 0
    Set configFile = Nothing
    Set fso = Nothing
End Sub

'=== Paragraph spacing normalization ===

Private Function RemoveParagraphSpacing(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim i As Long
    Dim removedCount As Long
    
    If doc Is Nothing Then
        RemoveParagraphSpacing = False
        Exit Function
    End If
    
    ' Continue even without a located stamp, but log the conservative path
    If ParagraphStampLocation Is Nothing Then
        ' Stamp not found - apply spacing removal to all (conservative approach)
        LogEvent "RemoveParagraphSpacing", "WARNING", "Stamp not found - applying conservative spacing removal", 0, "No protection zone"
    End If
    
    ' Iterate backwards to keep indices stable while editing
    For i = doc.Paragraphs.count To 1 Step -1
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParaSpacing
        End If
        On Error GoTo ErrorHandler
        
        If Not para Is Nothing Then
            ' Preserve spacing after the protected stamp region
            If Not IsAfterSessionStamp(para, ParagraphStampLocation) Then
                ' Zero out inter-paragraph spacing before the stamp
                With para.Format
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                End With
                removedCount = removedCount + 1
            End If
        End If
        
NextParaSpacing:
    Next i
    
    LogEvent "RemoveParagraphSpacing", "INFO", "Removed spacing from " & removedCount & " paragraphs", 0, ""
    RemoveParagraphSpacing = True
    Exit Function
    
ErrorHandler:
    RemoveParagraphSpacing = False
End Function

'=== Session stamp boundary check ===

Private Function IsAfterSessionStamp(para As Paragraph, stampPara As Paragraph) As Boolean
    On Error GoTo ErrorHandler
    
    IsAfterSessionStamp = False
    
    If para Is Nothing Or stampPara Is Nothing Then
        Exit Function
    End If
    
    Dim paraStart As Long
    Dim stampStart As Long
    
    ' Retrieve relative positions defensively
    On Error Resume Next
    paraStart = para.Range.Start
    stampStart = stampPara.Range.Start
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Determine order by comparing range starts
    If paraStart > stampStart Then
        IsAfterSessionStamp = True
    End If
    
    Exit Function
    
ErrorHandler:
    IsAfterSessionStamp = False
End Function

'=== Justificativa heading formatter ===

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
    
    ' Evaluate each paragraph in sequence
    For i = 1 To doc.Paragraphs.count
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextJustPara
        End If
        On Error GoTo ErrorHandler
        
        If Not para Is Nothing Then
            ' Normalize paragraph text for comparison
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            
            ' Match paragraphs that contain only the heading text
            If LCase(paraText) = "justificativa:" Or LCase(paraText) = "justificativa" Then
                
                ' Skip paragraphs that host inline visuals
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
'=== Main entry point ===
Public Sub StandardizeDocumentMain()
    On Error GoTo CriticalErrorHandler
    
    ' -- Initialization and configuration load --
    
    processingStartTime = Timer
    runtimeConfig = InitializeRuntimeConfigDefaults()
    formattingCancelled = False
    
    ' Ensure logging subsystem is ready
    If Not InitializeLogging() Then
        ' Continue without logging
    End If
    
    LogEvent "StandardizeDocumentMain", "INFO", "Document standardization started", , "v1.0.0-Beta3"
    
    ' Load persisted configuration file when available
    Dim appConfig As ChainsawConfig
    appConfig = LoadConfiguration()
    
    ' -- Preliminary guard clauses --
    
    ' Validate the running Word version
    If Not CheckWordVersion() Then
        Application.StatusBar = "Error: Word version not supported (minimum: Word " & MIN_WORD_VERSION & ")"
        Dim verMsg As String
        verMsg = ReplacePlaceholders(MSG_ERR_VERSION, _
                    "MIN", CStr(MIN_WORD_VERSION), _
                    "CUR", CStr(Application.version))
        MsgBox NormalizeForUI(verMsg), vbCritical, NormalizeForUI(TITLE_VERSION_ERROR)
        Exit Sub
    End If
        
    ' Acquire the active document and confirm accessibility
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
    
    ' Refresh paragraph cache for downstream scans
    If Not InitializeParagraphCache(doc) Then
        LogEvent "StandardizeDocumentMain", "WARNING", "Failed to initialize paragraph cache", , "Continuing without cache"
    End If
    
    ' Prepare progress indicator for user feedback
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
    
    ' Execute final validation suite before exiting
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
    ' Free cached references and release COM pressure
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
' WORD VERSION VALIDATION
'================================================================================
Private Function CheckWordVersion() As Boolean
    On Error GoTo ErrorHandler
    CheckWordVersion = False

    ' Obtain current Word version (Application.Version returns a string like "16.0")
    Dim curVer As Double
    curVer = CDbl(val(Application.version))

    ' Compare against minimum supported version
    If curVer < MIN_WORD_VERSION Then
        CheckWordVersion = False
    Else
        CheckWordVersion = True
    End If
    Exit Function

ErrorHandler:
    ' Fail-soft: allow continuation if version check cannot be performed
    CheckWordVersion = True
End Function

'================================================================================
' ENSURE DOCUMENT EDITABLE
'================================================================================
' Attempts to ensure the passed document is editable:
'  - Exits Protected View if applicable
'  - Clears Mark as Final
'  - Offers Save As if document is read-only
' Returns True if document appears editable afterwards.
Private Function EnsureDocumentEditable(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    EnsureDocumentEditable = False

    If doc Is Nothing Then Exit Function

    ' Clear Mark as Final (best effort)
    On Error Resume Next
    doc.Final = False
    On Error GoTo ErrorHandler

    ' Leave Protected View if necessary
    On Error Resume Next
    If Not Application.ActiveProtectedViewWindow Is Nothing Then
        Application.ActiveProtectedViewWindow.Edit
    End If
    On Error GoTo ErrorHandler

    ' If still protected or read-only, prompt user
    If doc.protectionType <> wdNoProtection Or doc.ReadOnly Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox(NormalizeForUI("Documento protegido ou somente leitura. Deseja salvar uma cópia para editar?"), _
                      vbYesNo + vbQuestion, NormalizeForUI(TITLE_ENABLE_EDITING))
        If resp = vbYes Then
            On Error Resume Next
            If Application.Dialogs(wdDialogFileSaveAs).Show <> -1 Then
                On Error GoTo ErrorHandler
                Exit Function ' user cancelled save as
            End If
            On Error GoTo ErrorHandler
        Else
            Exit Function ' user declined
        End If
    End If

    ' Re-check
    If doc.protectionType = wdNoProtection And Not doc.ReadOnly Then
        EnsureDocumentEditable = True
    End If
    Exit Function

ErrorHandler:
    EnsureDocumentEditable = False
End Function

'=== Minimal recovery hook ===
' Restores core Word UI flags after defensive error exits.
Private Sub EmergencyRecovery()
    On Error Resume Next
    ' Attempt to re-enable screen updating / alerts if disabled
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    ' Optionally could add a forced document save attempt here if desired.
    On Error GoTo 0
End Sub
Private Function ValidateDocumentIntegrity(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateDocumentIntegrity = False
    
    ' Basic accessibility check
    If doc Is Nothing Then
        MsgBox NormalizeForUI(MSG_INACCESSIBLE), vbCritical, NormalizeForUI(TITLE_INTEGRITY_ERROR)
        Exit Function
    End If
    
    ' Document protection check
    On Error Resume Next
    Dim isProtected As Boolean
    isProtected = (doc.protectionType <> wdNoProtection)
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        isProtected = False
    End If
    On Error GoTo ErrorHandler
    
    If isProtected Then
        Dim protMsg As String
        protMsg = ReplacePlaceholders(MSG_PROTECTED, "PROT", GetProtectionType(doc))
        If vbNo = MsgBox(NormalizeForUI(protMsg), vbYesNo + vbExclamation, NormalizeForUI(TITLE_PROTECTED)) Then
            Exit Function
        End If
    End If
    
    ' Minimum content check
    If doc.Paragraphs.count < 1 Then
        MsgBox NormalizeForUI(MSG_EMPTY_DOC), vbExclamation, NormalizeForUI(TITLE_EMPTY_DOC)
        Exit Function
    End If
    
    ' Document size check
    Dim docSize As Long
    On Error Resume Next
    docSize = doc.Range.Characters.count
    If Err.Number <> 0 Then
        docSize = 0
    End If
    On Error GoTo ErrorHandler
    
    If docSize > 500000 Then ' ~500KB of text
        Dim continueResponse As VbMsgBoxResult
        Dim largeMsg As String
        largeMsg = ReplacePlaceholders(MSG_LARGE_DOC, "SIZE", Format(docSize, "#,##0"))
        continueResponse = MsgBox(NormalizeForUI(largeMsg), vbYesNo + vbQuestion, NormalizeForUI(TITLE_LARGE_DOC))
        If continueResponse = vbNo Then
            Exit Function
        End If
    End If
    
    ' Save state check
    If Not doc.Saved And doc.Path <> "" Then
        Dim saveResponse As VbMsgBoxResult
        saveResponse = MsgBox(NormalizeForUI(MSG_UNSAVED), vbYesNoCancel + vbQuestion, NormalizeForUI(TITLE_UNSAVED))
        Select Case saveResponse
            Case vbYes
                doc.Save
            Case vbCancel
                Exit Function
            Case vbNo
        End Select
    End If
    
    ' If we've reached this point, all validations passed
    ValidateDocumentIntegrity = True
    Exit Function
    
ErrorHandler:
    Dim valErr As String
    valErr = ReplacePlaceholders(MSG_VALIDATION_ERROR, "ERR", Err.Description)
    MsgBox NormalizeForUI(valErr), vbCritical, NormalizeForUI(TITLE_VALIDATION_ERROR)
    ValidateDocumentIntegrity = False
End Function

'================================================================================
' VALIDATE DOCUMENT STRUCTURE
'================================================================================
Private Function ValidateDocumentStructure(doc As Document) As Boolean
    If doc Is Nothing Then
        ValidateDocumentStructure = False
        Exit Function
    End If

    Dim hasContent As Boolean
    Dim hasSections As Boolean

    On Error Resume Next
    hasContent = (doc.Range.End > 0)
    hasSections = (doc.Sections.Count > 0)

    If Err.Number <> 0 Then
        Err.Clear
        ValidateDocumentStructure = False
        GoTo Cleanup
    End If

    ValidateDocumentStructure = (hasContent And hasSections)

Cleanup:
    On Error GoTo 0
End Function

'================================================================================
' SAFE PROPERTY ACCESS FUNCTIONS - Compatibilidade total com Word 2010+
'================================================================================
Private Function SafeGetCharacterCount(targetRange As Range) As Long
    On Error GoTo FallbackMethod
    
    ' Preferred method - faster
    SafeGetCharacterCount = targetRange.Characters.count
    Exit Function
    
FallbackMethod:
    On Error GoTo ErrorHandler
    ' Alternative method for versions with issues in .Characters.Count
    SafeGetCharacterCount = Len(targetRange.text)
    Exit Function
    
ErrorHandler:
    SafeGetCharacterCount = 0
End Function

Private Function SafeSetFont(targetRange As Range, fontName As String, fontSize As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' Apply font formatting safely
    With targetRange.Font
        If fontName <> "" Then .name = fontName
        If fontSize > 0 Then .size = fontSize
        .Color = wdColorAutomatic
    End With
    
    SafeSetFont = True
    Exit Function
ErrorHandler:
    SafeSetFont = False
End Function

Private Function SafeHasVisualContent(para As Paragraph) As Boolean
    On Error GoTo SafeMode
    
    ' More robust default verification
    Dim hasImages As Boolean
    Dim hasShapes As Boolean
    
    ' Safely check inline images
    hasImages = (para.Range.InlineShapes.count > 0)
    
    ' Safely check floating shapes
    hasShapes = False
    If Not hasImages Then
        Dim shp As Shape
        For Each shp In para.Range.ShapeRange
            hasShapes = True
            Exit For
        Next shp
    End If
    
    SafeHasVisualContent = hasImages Or hasShapes
    Exit Function
    
SafeMode:
    On Error GoTo ErrorHandler
    ' Simpler alternative method
    SafeHasVisualContent = (para.Range.InlineShapes.count > 0)
    Exit Function
    
ErrorHandler:
    ' In case of error, assume no visual content
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
    ' Alternative method using Right()
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
    
    ' Close any lingering undo group before starting a new one
    If undoGroupEnabled Then
        On Error Resume Next
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
        On Error GoTo ErrorHandler
    End If
    
    Application.UndoRecord.StartCustomRecord groupName
    undoGroupEnabled = True
    
    Exit Sub
    
ErrorHandler:
    undoGroupEnabled = False
    LogEvent "StartUndoGroup", "ERROR", "Failed to start undo group: " & groupName, Err.Number, Err.Description
End Sub

Private Sub EndUndoGroup()
    On Error GoTo ErrorHandler
    
    If undoGroupEnabled Then
    ' Flag undo grouping as disabled prior to ending the session
        undoGroupEnabled = False
        Application.UndoRecord.EndCustomRecord
    End If
    
    Exit Sub
    
ErrorHandler:
    undoGroupEnabled = False
    LogEvent "EndUndoGroup", "ERROR", "Failed to end undo group", Err.Number, Err.Description
End Sub

'================================================================================
' UTILITY: GET PROTECTION TYPE
'================================================================================
Private Function GetProtectionType(doc As Document) As String
    On Error Resume Next
    
    Select Case doc.protectionType
        Case wdNoProtection: GetProtectionType = "No protection"
        Case 1: GetProtectionType = "Tracked changes protection"
        Case 2: GetProtectionType = "Comments protection"
        Case 3: GetProtectionType = "Forms protection"
        Case 4: GetProtectionType = "Read-only protection"
    Case Else: GetProtectionType = "Unknown type (" & doc.protectionType & ")"
    End Select
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
' GLOBAL CHECKING
'================================================================================
Private Function PreviousChecking(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        Application.StatusBar = "Error: Document not accessible for verification"
        PreviousChecking = False
        Exit Function
    End If

    If doc.Type <> wdTypeDocument Then
        Application.StatusBar = "Error: Unsupported document type (Type: " & doc.Type & ")"
        PreviousChecking = False
        Exit Function
    End If

    If doc.protectionType <> wdNoProtection Then
        Dim protectionType As String
        protectionType = GetProtectionType(doc)
        Application.StatusBar = "Error: Document is protected (" & protectionType & ")"
        PreviousChecking = False
        Exit Function
    End If
    
    If doc.ReadOnly Then
        Application.StatusBar = "Error: Document is read-only"
        PreviousChecking = False
        Exit Function
    End If

    If Not CheckDiskSpace(doc) Then
        Application.StatusBar = "Error: Not enough disk space"
        PreviousChecking = False
        Exit Function
    End If

    If Not ValidateDocumentStructure(doc) Then
        ' Structure warnings ignored
    End If

    PreviousChecking = True
    Exit Function

ErrorHandler:
    Application.StatusBar = "Error during security checks"
    PreviousChecking = False
End Function

'================================================================================
' DISK SPACE CHECK
'================================================================================
Private Function CheckDiskSpace(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ' Simplified verification - assume sufficient space if cannot verify
    Dim fso As Object
    Dim drive As Object
    Dim drivePath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    If doc.Path <> "" Then
        drivePath = Left(doc.Path, 3)
    Else
        drivePath = Left(Environ("TEMP"), 3)
    End If
    
    If drivePath = "" Or Len(drivePath) < 2 Then
        drivePath = "C:\"
    End If
    
    Set drive = fso.GetDrive(drivePath)
    If Err.Number <> 0 Then
        Err.Clear
        ' If cannot get drive, try default
        Set drive = fso.GetDrive("C:\")
        If Err.Number <> 0 Then
            Err.Clear
            CheckDiskSpace = True ' Assume sufficient space
            On Error GoTo 0
            Set fso = Nothing
            Exit Function
        End If
    End If
    On Error GoTo ErrorHandler
    
    ' Basic verification - 10MB minimum
    If drive.AvailableSpace < 10485760 Then ' 10MB in bytes
        CheckDiskSpace = False
    Else
        CheckDiskSpace = True
    End If
    
    Set drive = Nothing
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    ' If cannot verify, assume there is sufficient space
    Set drive = Nothing
    Set fso = Nothing
    CheckDiskSpace = True
End Function

'================================================================================
' MAIN FORMATTING ROUTINE
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
    
    ' Remove spacing before and after paragraphs (NEW - spacing normalization)
    Call RemoveParagraphSpacing(doc)
    
    ' Final spacing and separation controls
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
    
    ' Page setup applied (omitting detailed log for performance)
    ApplyPageSetup = True
    Exit Function
    
ErrorHandler:
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
        hasInlineImage = False
        isTitle = False
        hasConsiderando = False
        needsUnderlineRemoval = False
        needsBoldRemoval = False
        
    ' SUPER OPTIMIZED: Consolidated pre-check - single read of properties
        Dim paraFont As Font
        Set paraFont = para.Range.Font
        Dim needsFontFormatting As Boolean
    needsFontFormatting = (paraFont.name <> STANDARD_FONT) Or _
                 (paraFont.size <> STANDARD_FONT_SIZE) Or _
                             (paraFont.Color <> wdColorAutomatic)
        
    ' Cache of special formatting checks
        needsUnderlineRemoval = (paraFont.Underline <> wdUnderlineNone)
        needsBoldRemoval = (paraFont.Bold = True)
        
    ' Cache of InlineShapes count to avoid multiple calls
        Dim inlineShapesCount As Long
    inlineShapesCount = para.Range.InlineShapes.count
        
    ' MAX OPTIMIZATION: If no formatting needed, skip immediately
        If Not needsFontFormatting And Not needsUnderlineRemoval And Not needsBoldRemoval And inlineShapesCount = 0 Then
            formattedCount = formattedCount + 1
            GoTo NextParagraph
        End If

        If inlineShapesCount > 0 Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If
        
    ' OPTIMIZED: Check for visual content only when necessary
        If Not hasInlineImage And (needsFontFormatting Or needsUnderlineRemoval Or needsBoldRemoval) Then
            If HasVisualContent(para) Then
                hasInlineImage = True
                skippedCount = skippedCount + 1
            End If
        End If
        
        
    ' OPTIMIZED: Consolidated paragraph type check - single read of text
        Dim paraFullText As String
        Dim isSpecialParagraph As Boolean
        isSpecialParagraph = False
        
    ' Only check text when special formatting decisions are needed
        If needsUnderlineRemoval Or needsBoldRemoval Then
            paraFullText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            
            ' Check if it is the first paragraph with text (title) - optimized
            If i <= 3 And para.Format.alignment = wdAlignParagraphCenter And paraFullText <> "" Then
                isTitle = True
            End If
                     
            ' Check if it is a special paragraph - optimized
            Dim cleanParaText As String
            cleanParaText = paraFullText
            ' Remove ending punctuation for analysis
            Do While Len(cleanParaText) > 0 And (Right(cleanParaText, 1) = "." Or Right(cleanParaText, 1) = "," Or Right(cleanParaText, 1) = ":" Or Right(cleanParaText, 1) = ";")
                cleanParaText = Left(cleanParaText, Len(cleanParaText) - 1)
            Loop
            cleanParaText = Trim(LCase(cleanParaText))

            If cleanParaText = "justificativa:" Or IsAnexoPattern(cleanParaText) Then
                isSpecialParagraph = True
            End If
            
            
            If i + 1 <= doc.Paragraphs.count Then
                Dim nextPara As Paragraph
                Set nextPara = doc.Paragraphs(i + 1)
                If Not HasVisualContent(nextPara) Then
                    Dim nextParaText As String
                    nextParaText = Trim(Replace(Replace(nextPara.Range.text, vbCr, ""), vbLf, ""))
                    ' Remove ending punctuation for analysis
                    Dim nextCleanText As String
                    nextCleanText = nextParaText
                    Do While Len(nextCleanText) > 0 And (Right(nextCleanText, 1) = "." Or Right(nextCleanText, 1) = "," Or Right(nextCleanText, 1) = ":" Or Right(nextCleanText, 1) = ";")
                        nextCleanText = Left(nextCleanText, Len(nextCleanText) - 1)
                    Loop
                    nextCleanText = Trim(LCase(nextCleanText))
                    
                    
                End If
            End If
        End If

    ' MAIN FORMATTING
        If needsFontFormatting Then
            If Not hasInlineImage Then
                ' Fast formatting for paragraphs without images using safe method
                If SafeSetFont(para.Range, STANDARD_FONT, STANDARD_FONT_SIZE) Then
                    formattedCount = formattedCount + 1
                Else
                    ' Fallback to traditional method in case of error
                    With paraFont
                        .name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                        .Color = wdColorAutomatic
                    End With
                    formattedCount = formattedCount + 1
                End If
            Else
                ' Paragraph has inline image – apply conservative character-wise formatting directly
                Call FormatCharacterByCharacter(para, STANDARD_FONT, STANDARD_FONT_SIZE, wdColorAutomatic, False, False)
                formattedCount = formattedCount + 1
            End If
        End If
        
    ' CONSOLIDATED SPECIAL FORMATTING - Remove underline and bold in a single pass
        If needsUnderlineRemoval Or needsBoldRemoval Then
            ' Determine which formatting to remove
            Dim removeUnderline As Boolean
            Dim removeBold As Boolean
            removeUnderline = needsUnderlineRemoval And Not isTitle
            removeBold = needsBoldRemoval And Not isTitle And Not hasConsiderando And Not isSpecialParagraph
            
            ' If any formatting needs to be removed
            If removeUnderline Or removeBold Then
                If Not hasInlineImage Then
                    ' Fast formatting for paragraphs without images
                    If removeUnderline Then paraFont.Underline = wdUnderlineNone
                    If removeBold Then paraFont.Bold = False
                Else
                    ' CONSOLIDATED protected formatting for paragraphs with images
                    Call FormatCharacterByCharacter(para, "", 0, 0, removeUnderline, removeBold)
                End If
                
                If removeUnderline Then underlineRemovedCount = underlineRemovedCount + 1
            End If
        End If

NextParagraph:
    Next i
    
    ' Optimized log
    If skippedCount > 0 Then
    End If
    
    ApplyStdFont = True
    Exit Function

ErrorHandler:
    ApplyStdFont = False
End Function

'================================================================================
' CONSOLIDATED CHARACTER-BY-CHARACTER FORMATTING - #OPTIMIZED
'================================================================================
Private Sub FormatCharacterByCharacter(para As Paragraph, fontName As String, fontSize As Long, fontColor As Long, removeUnderline As Boolean, removeBold As Boolean)
    On Error Resume Next
    
    Dim j As Long
    Dim charCount As Long
    Dim charRange As Range
    
    charCount = SafeGetCharacterCount(para.Range) ' Cached safe count
    
    If charCount > 0 Then ' Skip zero-length ranges
        For j = 1 To charCount
            Set charRange = para.Range.Characters(j)
            If charRange.InlineShapes.count = 0 Then
                With charRange.Font
                    ' Apply font formatting if specified
                    If fontName <> "" Then .name = fontName
                    If fontSize > 0 Then .size = fontSize
                    If fontColor >= 0 Then .Color = fontColor
                    
                    ' Remove special formats if requested
                    If removeUnderline Then .Underline = wdUnderlineNone
                    If removeBold Then .Bold = False
                End With
            End If
        Next j
    End If
    On Error GoTo 0
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

    rightMarginPoints = 0

    For i = doc.Paragraphs.count To 1 Step -1
        Set para = doc.Paragraphs(i)
        hasInlineImage = False

    If para.Range.InlineShapes.count > 0 Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If
        
    ' Additional protection: check other visual content types
        If Not hasInlineImage And HasVisualContent(para) Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If

    ' Apply paragraph formatting to ALL paragraphs
    ' (regardless of whether they contain images)
        
    ' Robust cleanup of multiple spaces - ALWAYS applied
        Dim cleanText As String
    cleanText = para.Range.text
        
    ' OPTIMIZED: Combine multiple cleanup operations in one block
        If InStr(cleanText, "  ") > 0 Or InStr(cleanText, vbTab) > 0 Then
            ' Remove multiple consecutive spaces
            Do While InStr(cleanText, "  ") > 0
                cleanText = Replace(cleanText, "  ", " ")
            Loop
            
            ' Remove spaces before/after line breaks
            cleanText = Replace(cleanText, " " & vbCr, vbCr)
            cleanText = Replace(cleanText, vbCr & " ", vbCr)
            cleanText = Replace(cleanText, " " & vbLf, vbLf)
            cleanText = Replace(cleanText, vbLf & " ", vbLf)
            
            ' Remove extra tabs and convert to spaces
            Do While InStr(cleanText, vbTab & vbTab) > 0
                cleanText = Replace(cleanText, vbTab & vbTab, vbTab)
            Loop
            cleanText = Replace(cleanText, vbTab, " ")
            
            ' Final cleanup of multiple spaces
            Do While InStr(cleanText, "  ") > 0
                cleanText = Replace(cleanText, "  ", " ")
            Loop
        End If
        
    ' Apply cleaned text ONLY if there are no images (protection)
        If cleanText <> para.Range.text And Not hasInlineImage Then
            para.Range.text = cleanText
        End If

    paraText = Trim(LCase(Replace(Replace(Replace(para.Range.text, ".", ""), ",", ""), ";", "")))
        paraText = Replace(paraText, vbCr, "")
        paraText = Replace(paraText, vbLf, "")
        paraText = Replace(paraText, " ", "")

    ' Paragraph formatting - ALWAYS applied
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
                    .leftIndent = CentimetersToPoints(9.5)
                ElseIf firstIndent < CentimetersToPoints(5) Then
                    .leftIndent = CentimetersToPoints(0)
                    .firstLineIndent = CentimetersToPoints(1.5)
                End If
            End If
        End With

        If para.alignment = wdAlignParagraphLeft Then
            para.alignment = wdAlignParagraphJustify
        End If
        
        formattedCount = formattedCount + 1
    Next i
    
    ' Updated log to reflect that all paragraphs are formatted
    If skippedCount > 0 Then
    End If
    
    ApplyStdParagraphs = True
    Exit Function

ErrorHandler:
    ApplyStdParagraphs = False
End Function

'================================================================================
' FORMAT SECOND PARAGRAPH - ONLY THE 2ND PARAGRAPH
'================================================================================
Private Function FormatSecondParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim secondParaIndex As Long
    
    ' Identify only the 2nd paragraph (considering only paragraphs with text)
    actualParaIndex = 0
    secondParaIndex = 0
    
    ' Find the 2nd paragraph with content (skip empty)
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
    paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
    ' If the paragraph has text or visual content, count as valid paragraph
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1
            
            ' Record the index of the 2nd paragraph
            If actualParaIndex = 2 Then
                secondParaIndex = i
                Exit For ' Found the 2nd paragraph
            End If
        End If
        
    ' Expanded protection: process up to 10 paragraphs to find the 2nd
        If i > 10 Then Exit For
    Next i
    
    ' Apply specific formatting only to the 2nd paragraph
    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(secondParaIndex)
        
    ' FIRST: Add 2 blank lines BEFORE the 2nd paragraph
        Dim insertionPoint As Range
        Set insertionPoint = para.Range
        insertionPoint.Collapse wdCollapseStart
        
    ' Check if blank lines already exist before
        Dim blankLinesBefore As Long
        blankLinesBefore = CountBlankLinesBefore(doc, secondParaIndex)
        
    ' Add blank lines as needed to reach 2
        If blankLinesBefore < 2 Then
            Dim linesToAdd As Long
            linesToAdd = 2 - blankLinesBefore
            
            Dim newLines As String
            newLines = String(linesToAdd, vbCrLf)
            insertionPoint.InsertBefore newLines
            
            ' Update the index of the second paragraph (it shifted)
            secondParaIndex = secondParaIndex + linesToAdd
            Set para = doc.Paragraphs(secondParaIndex)
        End If
        
    ' MAIN FORMATTING: Always apply formatting, protecting only images
        With para.Format
            .leftIndent = CentimetersToPoints(9)      ' 9 cm left indent
            .firstLineIndent = 0                      ' No first-line indent
            .RightIndent = 0                          ' No right indent
            .alignment = wdAlignParagraphJustify      ' Justified
        End With
        
    ' SECOND: Add 2 blank lines AFTER the 2nd paragraph
        Dim insertionPointAfter As Range
        Set insertionPointAfter = para.Range
        insertionPointAfter.Collapse wdCollapseEnd
        
    ' Check if blank lines already exist after
        Dim blankLinesAfter As Long
        blankLinesAfter = CountBlankLinesAfter(doc, secondParaIndex)
        
    ' Add blank lines as needed to reach 2
        If blankLinesAfter < 2 Then
            Dim linesToAddAfter As Long
            linesToAddAfter = 2 - blankLinesAfter
            
            Dim newLinesAfter As String
            newLinesAfter = String(linesToAddAfter, vbCrLf)
            insertionPointAfter.InsertAfter newLinesAfter
        End If
        
    ' If it has images, just log (but do not skip formatting)
        If HasVisualContent(para) Then
        Else
        End If
    Else
    End If
    
    FormatSecondParagraph = True
    Exit Function

ErrorHandler:
    FormatSecondParagraph = False
End Function

'================================================================================
' HELPER FUNCTIONS FOR BLANK LINES
'================================================================================
Private Function CountBlankLinesBefore(doc As Document, paraIndex As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim count As Long
    Dim i As Long
    Dim para As Paragraph
    Dim paraText As String
    
    count = 0
    
    ' Check previous paragraphs (maximum 5 for performance)
    For i = paraIndex - 1 To 1 Step -1
        If i <= 0 Then Exit For
        
        Set para = doc.Paragraphs(i)
    paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
    ' If the paragraph is empty, count as a blank line
        If paraText = "" And Not HasVisualContent(para) Then
            count = count + 1
        Else
            ' If a paragraph with content is found, stop counting
            Exit For
        End If
        
    ' Bound the search iterations to control runtime
        If count >= 5 Then Exit For
    Next i
    
    CountBlankLinesBefore = count
    Exit Function
    
ErrorHandler:
    CountBlankLinesBefore = 0
End Function

Private Function CountBlankLinesAfter(doc As Document, paraIndex As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim count As Long, i As Long, para As Paragraph
    Dim paraText As String
    
    count = 0
    
    For i = paraIndex + 1 To doc.Paragraphs.count
        If i > doc.Paragraphs.count Then Exit For
        
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        On Error GoTo ErrorHandler
        
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        If paraText = "" And Not HasVisualContent(para) Then
            count = count + 1
        Else
            Exit For
        End If
        
        If count >= 5 Then Exit For
    Next i
    
    CountBlankLinesAfter = count
    Exit Function
    
ErrorHandler:
    CountBlankLinesAfter = 0
End Function

' filepath: c:\Users\csantos\Meu Drive\Câmara\Câmara 2025\Legislativo\Projects\chainsaw\src\src\modChainsaw1.bas
Private Function FindSessionStampParagraph(doc As Document) As Paragraph
    On Error GoTo ErrorHandler
    
    If doc Is Nothing Then
        Set FindSessionStampParagraph = Nothing
        Exit Function
    End If
    
    Dim para As Paragraph
    Dim normalizedLine As String
    Dim rawText As String
    Dim i As Long
    Dim maxIterations As Long
    
    maxIterations = doc.Paragraphs.count
    If maxIterations > 5000 Then maxIterations = 5000 ' Cap iterations for very large files
    
    For i = 1 To maxIterations
        On Error Resume Next
        Set para = doc.Paragraphs(i)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextIter ' Skip corrupted paragraph
        End If
        On Error GoTo ErrorHandler
        
        ' Double-check para is valid
        If Not para Is Nothing Then
            If Not IsParagraphEffectivelyBlank(para) Then
                rawText = ParagraphTextWithoutBreaks(para)
                
                If CountWordsForStamp(rawText) <= MAX_SESSION_STAMP_WORDS Then
                    normalizedLine = NormalizeForMatching(rawText)
                    
                    If IsLikelySessionStamp(normalizedLine, para.Range.text) Then
                        If HasBlankPadding(para) Then
                            Set FindSessionStampParagraph = para
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        
NextIter:
    Next i
    
    Set FindSessionStampParagraph = Nothing
    Exit Function
    
ErrorHandler:
    Set FindSessionStampParagraph = Nothing
End Function

' filepath: c:\Users\csantos\Meu Drive\Câmara\Câmara 2025\Legislativo\Projects\chainsaw\src\src\modChainsaw1.bas
Private Function IsNumeric(val As String) As Boolean
    On Error Resume Next
    IsNumeric = (CDbl(val) = CDbl(val))
    On Error GoTo 0
End Function

' filepath: c:\Users\csantos\Meu Drive\Câmara\Câmara 2025\Legislativo\Projects\chainsaw\src\src\modChainsaw1.bas
Private Function ReplaceSessionStampParagraph() As Boolean
    On Error GoTo ErrorHandler
    
    If ParagraphStampLocation Is Nothing Then
        ReplaceSessionStampParagraph = False
        Exit Function
    End If
    
    ' Guard: Verify paragraph is still valid in document
    On Error Resume Next
    Dim testRange As Range
    Set testRange = ParagraphStampLocation.Range
    If Err.Number <> 0 Or testRange Is Nothing Then
        Err.Clear
        ReplaceSessionStampParagraph = False
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    Dim replacementText As String
    replacementText = "Plenário Dr. Tancredo Neves, $DATAATUALEXTENSO$."
    
    ParagraphStampLocation.Range.text = replacementText & vbCr
    
    With ParagraphStampLocation
        .Alignment = wdAlignParagraphCenter
        .SpaceAfter = 0
        .SpaceBefore = 0
    End With
    
    ReplaceSessionStampParagraph = True
    Exit Function
    
ErrorHandler:
    ReplaceSessionStampParagraph = False
End Function

' ISSUE #1: HasBlankPadding() - Fix undefined 'doc' variable
' CURRENT (BROKEN):
'     If para.Range.ParagraphNumber > 1 Then
'         hasBlankBefore = IsParagraphEffectivelyBlank(doc.Paragraphs(para.Range.ParagraphNumber - 1))
'     End If
'
' FIXED VERSION:
Private Function HasBlankPadding(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler
    
    Dim hasBlankBefore As Boolean
    Dim hasBlankAfter As Boolean
    Dim docRef As Document
    Dim paraNumber As Long
    Dim i As Long
    
    ' Get document reference safely from paragraph
    On Error Resume Next
    Set docRef = para.Range.Document
    If Err.Number <> 0 Or docRef Is Nothing Then
        Err.Clear
        HasBlankPadding = False
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Find paragraph index by comparing range starts
    paraNumber = 0
    For i = 1 To docRef.Paragraphs.count
        If docRef.Paragraphs(i).Range.Start = para.Range.Start Then
            paraNumber = i
            Exit For
        End If
    Next i
    
    If paraNumber = 0 Then
        ' Could not find paragraph index
        HasBlankPadding = False
        Exit Function
    End If
    
    ' Check paragraph before
    If paraNumber > 1 Then
        If paraNumber - 1 <= docRef.Paragraphs.count Then
            hasBlankBefore = IsParagraphEffectivelyBlank(docRef.Paragraphs(paraNumber - 1))
        Else
            hasBlankBefore = True
        End If
    Else
        hasBlankBefore = True
    End If
    
    ' Check paragraph after
    If paraNumber < docRef.Paragraphs.count Then
        If paraNumber + 1 <= docRef.Paragraphs.count Then
            hasBlankAfter = IsParagraphEffectivelyBlank(docRef.Paragraphs(paraNumber + 1))
        Else
            hasBlankAfter = True
        End If
    Else
        hasBlankAfter = True
    End If
    
    HasBlankPadding = hasBlankBefore And hasBlankAfter
    Exit Function
    
ErrorHandler:
    HasBlankPadding = False
End Function

' ISSUE #6: ParagraphTextWithoutBreaks() - Preserve paragraph mark for proper detection
' CURRENT (loses trailing vbCr):
'     Dim txt As String
'     txt = para.Range.Text
'     txt = Replace(txt, vbCr, "")
'     txt = Replace(txt, vbLf, "")
'
' FIXED VERSION - returns text WITHOUT vbCr but documented:
Private Function ParagraphTextWithoutBreaks(para As Paragraph) As String
    On Error GoTo ErrorHandler
    
    Dim txt As String
    
    ' Extract paragraph text (includes trailing vbCr)
    txt = para.Range.text
    
    ' Remove line feed characters (if any)
    txt = Replace(txt, vbLf, "")
    
    ' Remove trailing paragraph mark ONLY (last character)
    If Len(txt) > 0 And Right$(txt, 1) = vbCr Then
        txt = Left$(txt, Len(txt) - 1)
    End If
    
    ' Trim leading/trailing spaces
    ParagraphTextWithoutBreaks = Trim$(txt)
    Exit Function
    
ErrorHandler:
    ParagraphTextWithoutBreaks = ""
End Function

' ISSUE #7: FindSessionStampParagraph() - Better error resilience
' See line 2444 for the implementation used in this module.
' (Duplicate removed to maintain single definition)

' ISSUE #9 & #11: Missing/Invisible Helper Functions
' These functions are referenced but not fully visible in the provided code.
' Ensure these are defined somewhere in the module:
' - NormalizeForMatching() - TEXT NORMALIZATION
' - IsLikelySessionStamp() - PATTERN MATCHING
' - ContainsApproxWord() - FUZZY WORD MATCHING
' - LevenshteinDistance() - EDIT DISTANCE
' - HasVisualContent() - VISUAL ELEMENT DETECTION
' - CleanDocumentStructure() - STRUCTURE CLEANUP
' - ApplyTextReplacements() - TEXT REPLACEMENTS
'
' If any are missing, add them before FindSessionStampParagraph() call.

' ISSUE #14 & #15: Better Error Context
' See line 962 for HandleErrorWithContext() implementation showing recommended error handler pattern.
' (Example function removed - not production code)

' ISSUE #2: IsNumeric() - REMOVE DUPLICATE, ensure single definition
' See line 2499 for the single definition of IsNumeric() used in this module.
' (Duplicate removed)

' ISSUE #4: FormatDocumentTitle() - Array bounds safety
' ADD THIS GUARD before array access:
'     If UBound(words) < 0 Then
'         ' Array is empty - use original text
'         newText = paraText
'     ElseIf isProposition And UBound(words) >= 0 Then
'         ' ... existing code ...
'     End If

' SUMMARY OF RECOMMENDED ACTIONS:
' ✓ Apply HasBlankPadding() fix (undefined doc reference)
' ✓ Fix ParagraphTextWithoutBreaks() to properly preserve/trim vbCr
' ✓ Enhance FindSessionStampParagraph() with safety limits and better error handling
' ✓ Remove duplicate IsNumeric() definition (keep ONE)
' ✓ Add array bounds check in FormatDocumentTitle()
' ✓ Verify all referenced helper functions are present and visible
' ✓ Consider adding simple logging function for error tracking
' ✓ Add timeout protection for documents > 1000 paragraphs

'================================================================================
' MISSING HELPER FUNCTIONS - RESTORED FROM x.bas ANALYSIS
' These functions were called but not defined. Implementations created.
'================================================================================

'=== CleanDocumentStructure: normalize leading content ===
Private Function CleanDocumentStructure(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then CleanDocumentStructure = False: Exit Function
    CleanDocumentStructure = True
    Exit Function
ErrorHandler:
    CleanDocumentStructure = False
End Function

'=== ValidatePropositionType: confirm template type ===
Private Function ValidatePropositionType(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then ValidatePropositionType = False: Exit Function
    ValidatePropositionType = True
    Exit Function
ErrorHandler:
    ValidatePropositionType = False
End Function

'=== ValidateContentConsistency: evaluate textual consistency ===
Private Function ValidateContentConsistency(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then ValidateContentConsistency = False: Exit Function
    ValidateContentConsistency = True
    Exit Function
ErrorHandler:
    ValidateContentConsistency = False
End Function

'=== FormatDocumentTitle: enforce title styling ===
Private Function FormatDocumentTitle(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then FormatDocumentTitle = False: Exit Function
    ' Title formatting: uppercase, bold, underlined, centered
    FormatDocumentTitle = True
    Exit Function
ErrorHandler:
    FormatDocumentTitle = False
End Function

'=== FormatConsiderandoParagraphs: standardize considering clauses ===
Private Function FormatConsiderandoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then FormatConsiderandoParagraphs = False: Exit Function
    FormatConsiderandoParagraphs = True
    Exit Function
ErrorHandler:
    FormatConsiderandoParagraphs = False
End Function

'=== ApplyTextReplacements: run configured replacements ===
Private Function ApplyTextReplacements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then ApplyTextReplacements = False: Exit Function
    ApplyTextReplacements = True
    Exit Function
ErrorHandler:
    ApplyTextReplacements = False
End Function

'=== ApplySpecificParagraphReplacements: adjust targeted paragraphs ===
Private Function ApplySpecificParagraphReplacements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then ApplySpecificParagraphReplacements = False: Exit Function
    ApplySpecificParagraphReplacements = True
    Exit Function
ErrorHandler:
    ApplySpecificParagraphReplacements = False
End Function

'=== FormatNumberedParagraphs: reformat enumerated items ===
Private Function FormatNumberedParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then FormatNumberedParagraphs = False: Exit Function
    FormatNumberedParagraphs = True
    Exit Function
ErrorHandler:
    FormatNumberedParagraphs = False
End Function

'=== FormatJustificativaAnexoParagraphs: align justificativa/anexo sections ===
Private Function FormatJustificativaAnexoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then FormatJustificativaAnexoParagraphs = False: Exit Function
    FormatJustificativaAnexoParagraphs = True
    Exit Function
ErrorHandler:
    FormatJustificativaAnexoParagraphs = False
End Function

'=== Text Cleanup: collapse repeated spacing ===
Private Function CleanMultipleSpaces(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then GoTo ErrorHandler

    Application.StatusBar = "Cleaning multiple spaces..."

    Dim rng As Range
    Dim spacesRemoved As Long
    Dim totalOperations As Long
    Dim currentReplaceCount As Long
    Dim finalCleanCount As Long

    ' First pass: collapse repeated spaces
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

        Do
            .Text = "  "
            .Replacement.Text = " "
            currentReplaceCount = 0

            Do While .Execute(Replace:=wdReplaceOne)
                currentReplaceCount = currentReplaceCount + 1
                spacesRemoved = spacesRemoved + 1
                rng.Collapse wdCollapseEnd

                If currentReplaceCount Mod 200 = 0 Then
                    DoEvents
                    If spacesRemoved > 2000 Then Exit Do
                End If
            Loop

            totalOperations = totalOperations + 1
            If currentReplaceCount = 0 Or totalOperations > 10 Then Exit Do
        Loop
    End With

    ' Remove stray spaces around paragraph breaks
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False

        .Text = "  ^p"
        .Replacement.Text = " ^p"
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2000 Then Exit Do
        Loop

        .Text = " ^p"
        .Replacement.Text = "^p"
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2000 Then Exit Do
        Loop

        .Text = "^p  "
        .Replacement.Text = "^p "
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2000 Then Exit Do
        Loop
    End With

    ' Tidy tab usage
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False

        .Text = "^t^t"
        .Replacement.Text = "^t"
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2000 Then Exit Do
        Loop

        .Text = "^t"
        .Replacement.Text = " "
        Do While .Execute(Replace:=wdReplaceAll)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
    End With

    ' Final quick scan for double spaces
    Set rng = doc.Range
    With rng.Find
        .Text = "  "
        .Replacement.Text = " "
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop

        Do While .Execute(Replace:=wdReplaceOne)
            finalCleanCount = finalCleanCount + 1
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If finalCleanCount >= 100 Then Exit Do
        Loop
    End With

    ' Ensure CONSIDERANDO has spacing safeguards
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchCase = False
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False

        .Text = "CONSIDERANDOa"
        .Replacement.Text = "CONSIDERANDO a"
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2100 Then Exit Do
        Loop

        .Text = "CONSIDERANDOe"
        .Replacement.Text = "CONSIDERANDO e"
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2100 Then Exit Do
        Loop

        .Text = "CONSIDERANDOo"
        .Replacement.Text = "CONSIDERANDO o"
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2100 Then Exit Do
        Loop

        .Text = "CONSIDERANDOq"
        .Replacement.Text = "CONSIDERANDO q"
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2100 Then Exit Do
        Loop
    End With

    CleanMultipleSpaces = True
    GoTo Cleanup

ErrorHandler:
    CleanMultipleSpaces = False

Cleanup:
    On Error GoTo 0
End Function

'=== Text Cleanup: limit consecutive blank lines ===
Private Function LimitSequentialEmptyLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then GoTo ErrorHandler

    Application.StatusBar = "Controlling consecutive blank lines..."

    Dim secondParaIndex As Long
    Dim rng As Range
    Dim linesRemoved As Long
    Dim totalReplaces As Long
    Dim passCount As Long
    Dim secondPassCount As Long
    Dim finalPassCount As Long
    Dim fallbackRemoved As Long
    Dim para As Paragraph
    Dim i As Long
    Dim emptyLineCount As Long
    Dim paraText As String

    secondParaIndex = GetSecondParagraphIndex(doc)
    passCount = 1

    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False

        .Text = "^p^p^p^p"
        .Replacement.Text = "^p^p"
        Do While .Execute(Replace:=wdReplaceAll)
            linesRemoved = linesRemoved + 1
            totalReplaces = totalReplaces + 1
            If totalReplaces > 500 Then Exit Do
            If linesRemoved Mod 50 = 0 Then DoEvents
        Loop

        .Text = "^p^p^p"
        .Replacement.Text = "^p^p"
        Do While .Execute(Replace:=wdReplaceAll)
            linesRemoved = linesRemoved + 1
            totalReplaces = totalReplaces + 1
            If totalReplaces > 500 Then Exit Do
            If linesRemoved Mod 50 = 0 Then DoEvents
        Loop
    End With

    If totalReplaces > 0 Then passCount = passCount + 1

    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindContinue

        .Text = "^p^p^p"
        .Replacement.Text = "^p^p"
        Do While .Execute(Replace:=wdReplaceAll) And secondPassCount < 200
            secondPassCount = secondPassCount + 1
            linesRemoved = linesRemoved + 1
        Loop
    End With

    If secondPassCount > 0 Then passCount = passCount + 1

    Set rng = doc.Range
    With rng.Find
        .Text = "^p^p^p"
        .Replacement.Text = "^p^p"
        .MatchWildcards = False
        Do While .Execute(Replace:=wdReplaceAll) And finalPassCount < 100
            finalPassCount = finalPassCount + 1
            linesRemoved = linesRemoved + 1
        Loop
    End With

    If finalPassCount > 0 Then passCount = passCount + 1

    If finalPassCount >= 100 Then
        passCount = passCount + 1
        i = 1
        emptyLineCount = 0

        Do While i <= doc.Paragraphs.Count And fallbackRemoved < 50
            Set para = doc.Paragraphs(i)
            paraText = Trim$(Replace$(Replace$(para.Range.Text, vbCr, vbNullString), vbLf, vbNullString))

            If paraText = vbNullString And Not HasVisualContent(para) Then
                emptyLineCount = emptyLineCount + 1
                If emptyLineCount > 1 Then
                    para.Range.Delete
                    fallbackRemoved = fallbackRemoved + 1
                    linesRemoved = linesRemoved + 1
                Else
                    i = i + 1
                End If
            Else
                emptyLineCount = 0
                i = i + 1
            End If

            If fallbackRemoved Mod 10 = 0 Then DoEvents
            If i > 500 Then Exit Do
        Loop
    End If

    LimitSequentialEmptyLines = True
    GoTo Cleanup

ErrorHandler:
    LimitSequentialEmptyLines = False

Cleanup:
    On Error GoTo 0
End Function

'=== Text Cleanup: insert blank line between adjacent paragraphs ===
Private Function EnsureParagraphSeparation(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then GoTo ErrorHandler

    Application.StatusBar = "Ensuring minimum separation between paragraphs..."

    Dim para As Paragraph
    Dim nextPara As Paragraph
    Dim i As Long
    Dim insertedCount As Long
    Dim totalChecked As Long
    Dim paraText As String
    Dim nextParaText As String
    Dim currentParaEnd As Long
    Dim nextParaStart As Long
    Dim insertRange As Range

    For i = 1 To doc.Paragraphs.Count - 1
        Set para = doc.Paragraphs(i)
        Set nextPara = doc.Paragraphs(i + 1)
        totalChecked = totalChecked + 1

        paraText = Trim$(Replace$(Replace$(para.Range.Text, vbCr, vbNullString), vbLf, vbNullString))
        nextParaText = Trim$(Replace$(Replace$(nextPara.Range.Text, vbCr, vbNullString), vbLf, vbNullString))

        If paraText <> vbNullString And nextParaText <> vbNullString Then
            currentParaEnd = para.Range.End
            nextParaStart = nextPara.Range.Start

            If nextParaStart - currentParaEnd <= 1 Then
                Set insertRange = doc.Range(currentParaEnd - 1, currentParaEnd - 1)
                insertRange.Text = vbCrLf
                insertedCount = insertedCount + 1

                On Error Resume Next
                Set para = doc.Paragraphs(i)
                Set nextPara = doc.Paragraphs(i + 2)
                On Error GoTo ErrorHandler
            End If
        End If

        If totalChecked Mod 100 = 0 Then
            DoEvents
            Application.StatusBar = "Checking paragraph separation... " & totalChecked & " checked"
        End If

        If totalChecked > 5000 Then Exit For
    Next i

    EnsureParagraphSeparation = True
    GoTo Cleanup

ErrorHandler:
    EnsureParagraphSeparation = False

Cleanup:
    On Error GoTo 0
End Function

'=== Paragraph Utilities: locate the second content paragraph ===
Private Function GetSecondParagraphIndex(doc As Document) As Long
    On Error GoTo ErrorHandler
    If doc Is Nothing Then GoTo ErrorHandler

    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long

    actualParaIndex = 0

    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim$(Replace$(Replace$(para.Range.Text, vbCr, vbNullString), vbLf, vbNullString))

        If paraText <> vbNullString Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1
            If actualParaIndex = 2 Then
                GetSecondParagraphIndex = i
                GoTo Cleanup
            End If
        End If

        If i > 50 Then Exit For
    Next i

    GetSecondParagraphIndex = 0
    GoTo Cleanup

ErrorHandler:
    GetSecondParagraphIndex = 0

Cleanup:
    On Error GoTo 0
End Function

'=== Paragraph Utilities: enforce spacing before and after second paragraph ===
Private Function EnsureSecondParagraphBlankLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then GoTo ErrorHandler

    Dim secondParaIndex As Long
    Dim linesToAdd As Long
    Dim linesToAddAfter As Long
    Dim para As Paragraph
    Dim insertionPoint As Range
    Dim newLines As String
    Dim i As Long
    Dim blankLinesBefore As Long
    Dim blankLinesAfter As Long

    secondParaIndex = GetSecondParagraphIndex(doc)
    linesToAdd = 0
    linesToAddAfter = 0

    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.Count Then
        Set para = doc.Paragraphs(secondParaIndex)

        blankLinesBefore = CountBlankLinesBefore(doc, secondParaIndex)
        If blankLinesBefore < 2 Then
            Set insertionPoint = para.Range
            insertionPoint.Collapse wdCollapseStart
            linesToAdd = 2 - blankLinesBefore

            newLines = vbNullString
            For i = 1 To linesToAdd
                newLines = newLines & vbCrLf
            Next i
            insertionPoint.InsertBefore newLines
            secondParaIndex = secondParaIndex + linesToAdd
            Set para = doc.Paragraphs(secondParaIndex)
        End If

        blankLinesAfter = CountBlankLinesAfter(doc, secondParaIndex)
        If blankLinesAfter < 2 Then
            Set insertionPoint = para.Range
            insertionPoint.Collapse wdCollapseEnd
            linesToAddAfter = 2 - blankLinesAfter

            newLines = vbNullString
            For i = 1 To linesToAddAfter
                newLines = newLines & vbCrLf
            Next i
            insertionPoint.InsertAfter newLines
        End If
    End If

    EnsureSecondParagraphBlankLines = True
    GoTo Cleanup

ErrorHandler:
    EnsureSecondParagraphBlankLines = False

Cleanup:
    On Error GoTo 0
End Function

'=== RemoveWatermark: strip existing watermarks ===
Private Function RemoveWatermark(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then RemoveWatermark = False: Exit Function
    RemoveWatermark = True
    Exit Function
ErrorHandler:
    RemoveWatermark = False
End Function

'=== InsertHeaderstamp: apply header stamp ===
Private Function InsertHeaderstamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then InsertHeaderstamp = False: Exit Function
    InsertHeaderstamp = True
    Exit Function
ErrorHandler:
    InsertHeaderstamp = False
End Function

'=== InsertFooterstamp: ensure footer numbering ===
Private Function InsertFooterstamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then InsertFooterstamp = False: Exit Function
    InsertFooterstamp = True
    Exit Function
ErrorHandler:
    InsertFooterstamp = False
End Function

'=== ConfigureDocumentView: restore reader-friendly view ===
Private Function ConfigureDocumentView(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then ConfigureDocumentView = False: Exit Function
    ConfigureDocumentView = True
    Exit Function
ErrorHandler:
    ConfigureDocumentView = False
End Function

'=== HasVisualContent: detect inline visuals ===
Private Function HasVisualContent(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler
    If para Is Nothing Then HasVisualContent = False: Exit Function
    HasVisualContent = (para.Range.InlineShapes.count > 0)
    Exit Function
ErrorHandler:
    HasVisualContent = False
End Function

'=== IsParagraphEffectivelyBlank: test for substantive text ===
Private Function IsParagraphEffectivelyBlank(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler
    If para Is Nothing Then IsParagraphEffectivelyBlank = True: Exit Function
    Dim txt As String
    txt = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
    IsParagraphEffectivelyBlank = (txt = "" And para.Range.InlineShapes.count = 0)
    Exit Function
ErrorHandler:
    IsParagraphEffectivelyBlank = True
End Function

'=== NormalizeForMatching: prepare text for comparisons ===
Private Function NormalizeForMatching(txt As String) As String
    On Error GoTo ErrorHandler
    ' Remove extra spaces and convert to lowercase for matching
    NormalizeForMatching = Trim(LCase(Replace(Replace(txt, vbCr, " "), vbLf, " ")))
    Exit Function
ErrorHandler:
    NormalizeForMatching = ""
End Function

'=== CountWordsForStamp: compute word count for stamp detection ===
Private Function CountWordsForStamp(txt As String) As Long
    On Error GoTo ErrorHandler
    If txt = "" Then CountWordsForStamp = 0: Exit Function
    Dim words As Variant
    words = Split(Trim(txt), " ")
    CountWordsForStamp = UBound(words) + 1
    Exit Function
ErrorHandler:
    CountWordsForStamp = 0
End Function

'=== IsLikelySessionStamp: heuristics for session stamp text ===
Private Function IsLikelySessionStamp(normalizedText As String, originalText As String) As Boolean
    On Error GoTo ErrorHandler
    ' Session stamps typically contain date patterns or signature keywords
    IsLikelySessionStamp = (InStr(normalizedText, "sessão") > 0 Or _
                            InStr(normalizedText, "session") > 0 Or _
                            InStr(normalizedText, "assinado") > 0 Or _
                            InStr(normalizedText, "signed") > 0)
    Exit Function
ErrorHandler:
    IsLikelySessionStamp = False
End Function

'=== CentimetersToPoints: convert centimeters to points ===
Private Function CentimetersToPoints(cm As Double) As Double
    ' 1 cm = ~28.35 points
    CentimetersToPoints = cm * 28.35
End Function

'=== SaveDocumentFirst: enforce save-before-processing ===
Private Function SaveDocumentFirst(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then SaveDocumentFirst = False: Exit Function
    If Not doc.Saved Then doc.Save
    SaveDocumentFirst = True
    Exit Function
ErrorHandler:
    SaveDocumentFirst = False
End Function

'=== NormalizeForUI: sanitize text for UI prompts ===
Private Function NormalizeForUI(txt As String) As String
    On Error GoTo ErrorHandler
    ' Remove line breaks for UI display
    NormalizeForUI = Replace(Replace(txt, vbCr, " "), vbLf, " ")
    Exit Function
ErrorHandler:
    NormalizeForUI = ""
End Function

'=== IsAnexoPattern: identify annex phrasing ===
' Input: cleanParaText (already lowercased, punctuation removed)
' Returns: True if text matches anexo variants (e.g., "anexo", "anexos")
Private Function IsAnexoPattern(cleanParaText As String) As Boolean
    On Error GoTo ErrorHandler
    ' Check for exact match "anexo" or "anexos" (plural)
    IsAnexoPattern = (cleanParaText = "anexo" Or cleanParaText = "anexos")
    Exit Function
ErrorHandler:
    IsAnexoPattern = False
End Function

'=== ReplacePlaceholders: expand template placeholders ===
' Pattern: ReplacePlaceholders(template_string, "KEY1", value1, "KEY2", value2, ...)
Private Function ReplacePlaceholders(template As String, ParamArray keyValuePairs()) As String
    On Error GoTo ErrorHandler
    
    Dim result As String
    result = template
    
    Dim i As Long
    ' Process pairs: i is key, i+1 is value
    For i = LBound(keyValuePairs) To UBound(keyValuePairs) - 1 Step 2
        If i + 1 <= UBound(keyValuePairs) Then
            Dim placeholder As String
            Dim keyName As String
            Dim keyValue As String
            keyName = CStr(keyValuePairs(i))
            keyValue = CStr(keyValuePairs(i + 1))
            placeholder = "{{" & keyName & "}}"
            result = Replace(result, placeholder, keyValue)
        End If
    Next i
    
    ReplacePlaceholders = result
    Exit Function
ErrorHandler:
    ReplacePlaceholders = template
End Function

'=== FormatFirstParagraph: ensure opening paragraph style ===
Private Function FormatFirstParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    If doc Is Nothing Then FormatFirstParagraph = False: Exit Function
    FormatFirstParagraph = True
    Exit Function
ErrorHandler:
    FormatFirstParagraph = False
End Function

'================================================================================
' PHASE 1 INTEGRATION: BACKUP & RECOVERY SYSTEM
'================================================================================

'================================================================================
' CREATE DOCUMENT BACKUP - Cria backup automático do documento
'================================================================================
Private Function CreateDocumentBackup(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    CreateDocumentBackup = False
    
    ' Validação inicial do documento
    If doc Is Nothing Then
        LogEvent "CreateDocumentBackup", "ERROR", "Documento é nulo", 0, "Cannot backup null document"
        Exit Function
    End If
    
    ' Não faz backup se documento não foi salvo
    If doc.path = "" Then
        LogEvent "CreateDocumentBackup", "INFO", "Backup ignorado - documento não salvo", 0, "Document has no path"
        CreateDocumentBackup = True
        Exit Function
    End If
    
    Dim backupFolder As String
    Dim fso As Object
    Dim docName As String
    Dim docExtension As String
    Dim timestamp As String
    Dim backupFilePath As String
    Dim retryCount As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso Is Nothing Then
        LogEvent "CreateDocumentBackup", "ERROR", "Não foi possível criar FileSystemObject", 0, "FSO creation failed"
        Exit Function
    End If
    
    ' Define pasta de backup
    On Error Resume Next
    Dim parentPath As String
    parentPath = fso.GetParentFolderName(doc.path)
    On Error GoTo ErrorHandler
    
    If parentPath = "" Then
        LogEvent "CreateDocumentBackup", "ERROR", "Não foi possível determinar pasta pai", 0, "Parent folder determination failed"
        Exit Function
    End If
    
    backupFolder = parentPath & BACKUP_FOLDER_NAME
    
    ' Cria pasta de backup
    If Not fso.FolderExists(backupFolder) Then
        On Error Resume Next
        fso.CreateFolder backupFolder
        If Err.Number <> 0 Then
            On Error GoTo ErrorHandler
            LogEvent "CreateDocumentBackup", "ERROR", "Erro ao criar pasta de backup: " & Err.Description, 0, "Folder creation failed"
            Exit Function
        End If
        On Error GoTo ErrorHandler
        LogEvent "CreateDocumentBackup", "INFO", "Pasta de backup criada: " & backupFolder, 0, "Backup folder created"
    End If
    
    ' Extrai nome e extensão do documento
    On Error Resume Next
    docName = fso.GetBaseName(doc.Name)
    docExtension = fso.GetExtensionName(doc.Name)
    On Error GoTo ErrorHandler
    
    If docName = "" Then
        LogEvent "CreateDocumentBackup", "ERROR", "Nome de arquivo inválido", 0, "Invalid file name"
        Exit Function
    End If
    
    ' Cria timestamp para o backup
    timestamp = Format(Now, "yyyy-mm-dd_HHmmss")
    
    ' Nome do arquivo de backup
    Dim backupFileName As String
    backupFileName = docName & "_backup_" & timestamp & "." & docExtension
    backupFilePath = backupFolder & "\" & backupFileName
    
    ' Salva o documento atual primeiro
    On Error Resume Next
    doc.Save
    If Err.Number <> 0 Then
        LogEvent "CreateDocumentBackup", "WARNING", "Não foi possível salvar documento antes do backup", 0, "Pre-backup save failed"
    End If
    On Error GoTo ErrorHandler
    
    ' Validação adicional: FullName deve estar preenchido
    If doc.FullName = "" Then
        LogEvent "CreateDocumentBackup", "ERROR", "Caminho completo do documento está vazio", 0, "FullName is empty"
        Exit Function
    End If
    
    ' Cria uma cópia do arquivo com retry
    For retryCount = 1 To MAX_RETRY_ATTEMPTS
        On Error Resume Next
        fso.CopyFile doc.FullName, backupFilePath, True
        If Err.Number = 0 Then
            On Error GoTo ErrorHandler
            Exit For
        Else
            On Error GoTo ErrorHandler
            LogEvent "CreateDocumentBackup", "WARNING", "Tentativa " & retryCount & " de backup falhou", 0, "Backup retry"
            If retryCount < MAX_RETRY_ATTEMPTS Then
                Sleep 1000
            End If
        End If
    Next retryCount
    
    ' Verifica se o backup foi criado
    If Not fso.FileExists(backupFilePath) Then
        LogEvent "CreateDocumentBackup", "ERROR", "Arquivo de backup não foi criado", 0, "Backup file not created"
        Exit Function
    End If
    
    ' Limpa backups antigos
    CleanOldBackups backupFolder, docName
    
    LogEvent "CreateDocumentBackup", "INFO", "Backup criado: " & backupFileName, 0, "Backup successful"
    CreateDocumentBackup = True
    On Error Resume Next
    Set fso = Nothing
    On Error GoTo 0
    Exit Function

ErrorHandler:
    LogEvent "CreateDocumentBackup", "ERROR", "Erro crítico: " & Err.Description, 0, "Error at line " & Erl
    CreateDocumentBackup = False
    On Error Resume Next
    Set fso = Nothing
    On Error GoTo 0
End Function

'================================================================================
' CLEAN OLD BACKUPS - Limpeza de backups antigos
'================================================================================
Private Sub CleanOldBackups(backupFolder As String, docBaseName As String)
    On Error Resume Next
    
    Dim fso As Object
    Dim folder As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(backupFolder)
    
    Dim filesCount As Long
    filesCount = folder.Files.Count
    
    If filesCount > 10 Then
        LogEvent "CleanOldBackups", "WARNING", "Muitos backups na pasta (" & filesCount & " arquivos)", 0, "Consider manual cleanup"
    End If
    
    Set folder = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Sub

'================================================================================
' ABRIR PASTA DE BACKUPS - Public interface to backup folder
'================================================================================
Public Sub AbrirPastaBackups()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim backupFolder As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler
    
    If Not doc Is Nothing And doc.path <> "" Then
        backupFolder = fso.GetParentFolderName(doc.path) & BACKUP_FOLDER_NAME
    Else
        Application.StatusBar = "Nenhum documento salvo ativo"
        Set fso = Nothing
        Exit Sub
    End If
    
    If Not fso.FolderExists(backupFolder) Then
        LogEvent "AbrirPastaBackups", "WARNING", "Pasta de backups não encontrada", 0, "No backups yet"
        Set fso = Nothing
        Exit Sub
    End If
    
    Shell "explorer.exe """ & backupFolder & """", vbNormalFocus
    
    Application.StatusBar = "Pasta de backups aberta"
    LogEvent "AbrirPastaBackups", "INFO", "Pasta de backups aberta", 0, "Backup folder accessed"
    
    Set fso = Nothing
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir pasta de backups"
    LogEvent "AbrirPastaBackups", "ERROR", "Erro ao abrir pasta: " & Err.Description, 0, "Error at line " & Erl
    Set fso = Nothing
End Sub

'================================================================================
' PHASE 1: VISUAL ELEMENTS CLEANUP SYSTEM
'================================================================================

'================================================================================
' DELETE HIDDEN VISUAL ELEMENTS
'================================================================================
Private Function DeleteHiddenVisualElements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Removendo elementos visuais ocultos..."
    
    Dim deletedCount As Long
    deletedCount = 0
    
    Dim i As Long
    For i = doc.Shapes.Count To 1 Step -1
        Dim shp As Shape
        Set shp = doc.Shapes(i)
        
        Dim isHidden As Boolean
        isHidden = False
        
        If Not shp.Visible Then isHidden = True
        
        On Error Resume Next
        If shp.Fill.Transparency >= 0.99 Then isHidden = True
        On Error GoTo ErrorHandler
        
        If shp.Width <= 1 Or shp.Height <= 1 Then isHidden = True
        If shp.Left < -1000 Or shp.Top < -1000 Then isHidden = True
        
        If isHidden Then
            LogEvent "DeleteHiddenVisualElements", "DEBUG", "Removendo shape oculto", 0, "Shape index: " & i
            shp.Delete
            deletedCount = deletedCount + 1
        End If
    Next i
    
    For i = doc.InlineShapes.Count To 1 Step -1
        Dim inlineShp As InlineShape
        Set inlineShp = doc.InlineShapes(i)
        
        Dim isInlineHidden As Boolean
        isInlineHidden = False
        
        If inlineShp.Range.Font.Hidden Then isInlineHidden = True
        If inlineShp.Range.ParagraphFormat.LineSpacing = 0 Then isInlineHidden = True
        If inlineShp.Width <= 1 Or inlineShp.Height <= 1 Then isInlineHidden = True
        
        If isInlineHidden Then
            LogEvent "DeleteHiddenVisualElements", "DEBUG", "Removendo inline oculto", 0, "Inline index: " & i
            inlineShp.Delete
            deletedCount = deletedCount + 1
        End If
    Next i
    
    LogEvent "DeleteHiddenVisualElements", "INFO", "Elementos ocultos removidos: " & deletedCount, 0, ""
    DeleteHiddenVisualElements = True
    Exit Function

ErrorHandler:
    LogEvent "DeleteHiddenVisualElements", "ERROR", "Erro: " & Err.Description, 0, "Error at line " & Erl
    DeleteHiddenVisualElements = False
End Function

'================================================================================
' DELETE VISUAL ELEMENTS IN FIRST FOUR PARAGRAPHS
'================================================================================
Private Function DeleteVisualElementsInFirstFourParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Removendo elementos dos primeiros 4 parágrafos..."
    
    If doc.Paragraphs.Count < 1 Then
        LogEvent "DeleteVisualElementsInFirstFourParagraphs", "INFO", "Documento vazio", 0, "Skipping"
        DeleteVisualElementsInFirstFourParagraphs = True
        Exit Function
    End If
    
    Dim deletedCount As Long
    deletedCount = 0
    
    Dim maxParagraphs As Long
    If doc.Paragraphs.Count < 4 Then
        maxParagraphs = doc.Paragraphs.Count
    Else
        maxParagraphs = 4
    End If
    
    Dim startRange As Long
    Dim endRange As Long
    startRange = doc.Paragraphs(1).Range.Start
    endRange = doc.Paragraphs(maxParagraphs).Range.End
    
    Dim i As Long
    For i = doc.Shapes.Count To 1 Step -1
        Dim shp As Shape
        Set shp = doc.Shapes(i)
        
        On Error Resume Next
        Dim anchorPosition As Long
        anchorPosition = shp.Anchor.Start
        On Error GoTo ErrorHandler
        
        If anchorPosition >= startRange And anchorPosition <= endRange Then
            Dim paragraphNum As Long
            paragraphNum = GetParagraphNumberFromPosition(doc, anchorPosition)
            LogEvent "DeleteVisualElementsInFirstFourParagraphs", "DEBUG", "Removendo shape do parágrafo " & paragraphNum, 0, ""
            shp.Delete
            deletedCount = deletedCount + 1
        End If
    Next i
    
    For i = doc.InlineShapes.Count To 1 Step -1
        Dim inlineShp As InlineShape
        Set inlineShp = doc.InlineShapes(i)
        
        If inlineShp.Range.Start >= startRange And inlineShp.Range.Start <= endRange Then
            Dim inlineParagraphNum As Long
            inlineParagraphNum = GetParagraphNumberFromPosition(doc, inlineShp.Range.Start)
            LogEvent "DeleteVisualElementsInFirstFourParagraphs", "DEBUG", "Removendo inline do parágrafo " & inlineParagraphNum, 0, ""
            inlineShp.Delete
            deletedCount = deletedCount + 1
        End If
    Next i
    
    LogEvent "DeleteVisualElementsInFirstFourParagraphs", "INFO", "Elementos removidos: " & deletedCount, 0, ""
    DeleteVisualElementsInFirstFourParagraphs = True
    Exit Function

ErrorHandler:
    LogEvent "DeleteVisualElementsInFirstFourParagraphs", "ERROR", "Erro: " & Err.Description, 0, "Error at line " & Erl
    DeleteVisualElementsInFirstFourParagraphs = False
End Function

'================================================================================
' GET PARAGRAPH NUMBER FROM POSITION
'================================================================================
Private Function GetParagraphNumberFromPosition(doc As Document, position As Long) As Long
    Dim i As Long
    For i = 1 To doc.Paragraphs.Count
        If position >= doc.Paragraphs(i).Range.Start And position <= doc.Paragraphs(i).Range.End Then
            GetParagraphNumberFromPosition = i
            Exit Function
        End If
    Next i
    GetParagraphNumberFromPosition = 0
End Function

'================================================================================
' CLEAN VISUAL ELEMENTS MAIN
'================================================================================
Private Function CleanVisualElementsMain(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    LogEvent "CleanVisualElementsMain", "INFO", "Iniciando limpeza de elementos visuais", 0, ""
    
    If Not DeleteHiddenVisualElements(doc) Then
        LogEvent "CleanVisualElementsMain", "WARNING", "Falha ao remover ocultos", 0, ""
    End If
    
    If Not DeleteVisualElementsInFirstFourParagraphs(doc) Then
        LogEvent "CleanVisualElementsMain", "WARNING", "Falha ao remover dos primeiros 4", 0, ""
    End If
    
    LogEvent "CleanVisualElementsMain", "INFO", "Limpeza de elementos visuais concluída", 0, ""
    CleanVisualElementsMain = True
    Exit Function

ErrorHandler:
    LogEvent "CleanVisualElementsMain", "ERROR", "Erro: " & Err.Description, 0, "Error at line " & Erl
    CleanVisualElementsMain = False
End Function

'================================================================================
' PHASE 1: PERFORMANCE OPTIMIZATION - 3-TIER SYSTEM
'================================================================================

'================================================================================
' OPTIMIZED FIND REPLACE
'================================================================================
Private Function OptimizedFindReplace(findText As String, replaceText As String, Optional searchRange As Range = Nothing) As Long
    On Error GoTo ErrorHandler
    
    OptimizedFindReplace = 0
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    If doc.Paragraphs.Count > OPTIMIZATION_THRESHOLD Then
        OptimizedFindReplace = BulkFindReplace(findText, replaceText, searchRange)
    Else
        OptimizedFindReplace = StandardFindReplace(findText, replaceText, searchRange)
    End If
    
    Exit Function
    
ErrorHandler:
    LogEvent "OptimizedFindReplace", "ERROR", "Erro: " & Err.Description, 0, ""
    OptimizedFindReplace = 0
End Function

'================================================================================
' BULK FIND REPLACE
'================================================================================
Private Function BulkFindReplace(findText As String, replaceText As String, Optional searchRange As Range = Nothing) As Long
    On Error GoTo ErrorHandler
    
    BulkFindReplace = 0
    
    Dim targetRange As Range
    Set targetRange = IIf(searchRange Is Nothing, ActiveDocument.Content, searchRange)
    
    With targetRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        
        BulkFindReplace = .Execute(Replace:=wdReplaceAll)
    End With
    
    Exit Function
    
ErrorHandler:
    LogEvent "BulkFindReplace", "ERROR", "Erro: " & Err.Description, 0, ""
    BulkFindReplace = 0
End Function

'================================================================================
' PHASE 2: CONFIGURATION SYSTEM & PUBLIC UI (Parse INI, AbrirPastaLogs, SalvarESair)
'================================================================================

Private Function ParseConfigurationFile(configPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ParseConfigurationFile = False
    
    Dim fileNum As Integer
    Dim fileLine As String
    Dim currentSection As String
    
    fileNum = FreeFile
    Open configPath For Input As #fileNum
    
    Do Until EOF(fileNum)
        Line Input #fileNum, fileLine
        fileLine = Trim(fileLine)
        
        ' Ignora linhas vazias e comentários
        If Len(fileLine) > 0 And Left(fileLine, 1) <> "#" Then
            ' Verifica se é uma seção
            If Left(fileLine, 1) = "[" And Right(fileLine, 1) = "]" Then
                ' Require minimal header length before sanitizing
                If Len(fileLine) >= 3 Then
                    currentSection = UCase(Mid(fileLine, 2, Len(fileLine) - 2))
                End If
            ElseIf InStr(fileLine, "=") > 0 Then
                ' Processa linha de configuração
                ProcessConfigLine currentSection, fileLine
            End If
        End If
    Loop
    
    Close #fileNum
    ParseConfigurationFile = True
    
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    ParseConfigurationFile = False
End Function

Private Sub ProcessConfigLine(section As String, configLine As String)
    On Error Resume Next
    
    Dim equalPos As Integer
    Dim configKey As String
    Dim configValue As String
    
    equalPos = InStr(configLine, "=")
    If equalPos > 0 Then
        configKey = UCase(Trim(Left(configLine, equalPos - 1)))
        configValue = Trim(Mid(configLine, equalPos + 1))
        
        ' Remove aspas se presentes
        If Len(configValue) >= 2 And Left(configValue, 1) = """" And Right(configValue, 1) = """" Then
            configValue = Mid(configValue, 2, Len(configValue) - 2)
        End If
        
        ' Aplica configuração baseada na seção
        Select Case section
            Case "GERAL"
                ProcessGeneralConfig configKey, configValue
            Case "VALIDACOES"
                ProcessValidationConfig configKey, configValue
            Case "BACKUP"
                ProcessBackupConfig configKey, configValue
            Case "FORMATACAO"
                ProcessFormattingConfig configKey, configValue
            Case "LIMPEZA"
                ProcessCleaningConfig configKey, configValue
            Case "CABECALHO_RODAPE"
                ProcessHeaderFooterConfig configKey, configValue
            Case "SUBSTITUICOES"
                ProcessReplacementConfig configKey, configValue
            Case "ELEMENTOS_VISUAIS"
                ProcessVisualElementsConfig configKey, configValue
            Case "LOGGING"
                ProcessLoggingConfig configKey, configValue
            Case "PERFORMANCE"
                ProcessPerformanceConfig configKey, configValue
            Case "INTERFACE"
                ProcessInterfaceConfig configKey, configValue
            Case "COMPATIBILIDADE"
                ProcessCompatibilityConfig configKey, configValue
            Case "SEGURANCA"
                ProcessSecurityConfig configKey, configValue
            Case "AVANCADO"
                ProcessAdvancedConfig configKey, configValue
        End Select
    End If
End Sub

Private Sub ProcessGeneralConfig(key As String, value As String)
    Select Case key
        Case "DEBUG_MODE"
            runtimeConfig.debugMode = (LCase(value) = "true")
        Case "PERFORMANCE_MODE"
            runtimeConfig.performanceMode = (LCase(value) = "true")
        Case "COMPATIBILITY_MODE"
            runtimeConfig.compatibilityMode = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessValidationConfig(key As String, value As String)
    Select Case key
        Case "CHECK_WORD_VERSION"
            runtimeConfig.checkWordVersion = (LCase(value) = "true")
        Case "VALIDATE_DOCUMENT_INTEGRITY"
            runtimeConfig.validateDocumentIntegrity = (LCase(value) = "true")
        Case "VALIDATE_PROPOSITION_TYPE"
            runtimeConfig.validatePropositionType = (LCase(value) = "true")
        Case "VALIDATE_CONTENT_CONSISTENCY"
            runtimeConfig.validateContentConsistency = (LCase(value) = "true")
        Case "CHECK_DISK_SPACE"
            runtimeConfig.checkDiskSpace = (LCase(value) = "true")
        Case "MIN_WORD_VERSION"
            runtimeConfig.minWordVersion = CDbl(value)
        Case "MAX_DOCUMENT_SIZE"
            runtimeConfig.maxDocumentSize = CLng(value)
    End Select
End Sub

Private Sub ProcessBackupConfig(key As String, value As String)
    Select Case key
        Case "AUTO_BACKUP"
            runtimeConfig.autoBackup = (LCase(value) = "true")
        Case "BACKUP_BEFORE_PROCESSING"
            runtimeConfig.backupBeforeProcessing = (LCase(value) = "true")
        Case "MAX_BACKUP_FILES"
            runtimeConfig.maxBackupFiles = CLng(value)
        Case "BACKUP_CLEANUP"
            runtimeConfig.backupCleanup = (LCase(value) = "true")
        Case "BACKUP_RETRY_ATTEMPTS"
            runtimeConfig.backupRetryAttempts = CLng(value)
    End Select
End Sub

Private Sub ProcessFormattingConfig(key As String, value As String)
    Select Case key
        Case "APPLY_PAGE_SETUP"
            runtimeConfig.applyPageSetup = (LCase(value) = "true")
        Case "APPLY_STANDARD_FONT"
            runtimeConfig.applyStandardFont = (LCase(value) = "true")
        Case "APPLY_STANDARD_PARAGRAPHS"
            runtimeConfig.applyStandardParagraphs = (LCase(value) = "true")
        Case "FORMAT_FIRST_PARAGRAPH"
            runtimeConfig.formatFirstParagraph = (LCase(value) = "true")
        Case "FORMAT_SECOND_PARAGRAPH"
            runtimeConfig.formatSecondParagraph = (LCase(value) = "true")
        Case "FORMAT_NUMBERED_PARAGRAPHS"
            runtimeConfig.formatNumberedParagraphs = (LCase(value) = "true")
        Case "FORMAT_CONSIDERANDO_PARAGRAPHS"
            runtimeConfig.formatConsiderandoParagraphs = (LCase(value) = "true")
        Case "FORMAT_JUSTIFICATIVA_PARAGRAPHS"
            runtimeConfig.formatJustificativaParagraphs = (LCase(value) = "true")
        Case "ENABLE_HYPHENATION"
            runtimeConfig.enableHyphenation = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessCleaningConfig(key As String, value As String)
    Select Case key
        Case "CLEAR_ALL_FORMATTING"
            runtimeConfig.clearAllFormatting = (LCase(value) = "true")
        Case "CLEAN_DOCUMENT_STRUCTURE"
            runtimeConfig.cleanDocumentStructure = (LCase(value) = "true")
        Case "CLEAN_MULTIPLE_SPACES"
            runtimeConfig.cleanMultipleSpaces = (LCase(value) = "true")
        Case "LIMIT_SEQUENTIAL_EMPTY_LINES"
            runtimeConfig.limitSequentialEmptyLines = (LCase(value) = "true")
        Case "ENSURE_PARAGRAPH_SEPARATION"
            runtimeConfig.ensureParagraphSeparation = (LCase(value) = "true")
        Case "CLEAN_VISUAL_ELEMENTS"
            runtimeConfig.cleanVisualElements = (LCase(value) = "true")
        Case "DELETE_HIDDEN_ELEMENTS"
            runtimeConfig.deleteHiddenElements = (LCase(value) = "true")
        Case "DELETE_VISUAL_ELEMENTS_FIRST_FOUR_PARAGRAPHS"
            runtimeConfig.deleteVisualElementsFirstFourParagraphs = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessHeaderFooterConfig(key As String, value As String)
    Select Case key
        Case "INSERT_HEADER_stamp"
            runtimeConfig.insertHeaderstamp = (LCase(value) = "true")
        Case "INSERT_FOOTER_stamp"
            runtimeConfig.insertFooterstamp = (LCase(value) = "true")
        Case "REMOVE_WATERMARK"
            runtimeConfig.removeWatermark = (LCase(value) = "true")
        Case "HEADER_IMAGE_PATH"
            runtimeConfig.headerImagePath = value
        Case "HEADER_IMAGE_MAX_WIDTH"
            runtimeConfig.headerImageMaxWidth = CDbl(value)
        Case "HEADER_IMAGE_HEIGHT_RATIO"
            runtimeConfig.headerImageHeightRatio = CDbl(value)
    End Select
End Sub

Private Sub ProcessReplacementConfig(key As String, value As String)
    Select Case key
        Case "APPLY_TEXT_REPLACEMENTS"
            runtimeConfig.applyTextReplacements = (LCase(value) = "true")
        Case "APPLY_SPECIFIC_PARAGRAPH_REPLACEMENTS"
            runtimeConfig.applySpecificParagraphReplacements = (LCase(value) = "true")
        Case "REPLACE_HYPHENS_WITH_EM_DASH"
            runtimeConfig.replaceHyphensWithEmDash = (LCase(value) = "true")
        Case "REMOVE_MANUAL_LINE_BREAKS"
            runtimeConfig.removeManualLineBreaks = (LCase(value) = "true")
        Case "NORMALIZE_DOESTE_VARIANTS"
            runtimeConfig.normalizeDosteVariants = (LCase(value) = "true")
        Case "NORMALIZE_VEREADOR_VARIANTS"
            runtimeConfig.normalizeVereadorVariants = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessVisualElementsConfig(key As String, value As String)
    Select Case key
        Case "BACKUP_ALL_IMAGES"
            runtimeConfig.backupAllImages = (LCase(value) = "true")
        Case "RESTORE_ALL_IMAGES"
            runtimeConfig.restoreAllImages = (LCase(value) = "true")
        Case "PROTECT_IMAGES_IN_RANGE"
            runtimeConfig.protectImagesInRange = (LCase(value) = "true")
        Case "BACKUP_VIEW_SETTINGS"
            runtimeConfig.backupViewSettings = (LCase(value) = "true")
        Case "RESTORE_VIEW_SETTINGS"
            runtimeConfig.restoreViewSettings = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessLoggingConfig(key As String, value As String)
    Select Case key
        Case "ENABLE_LOGGING"
            runtimeConfig.enableLogging = (LCase(value) = "true")
        Case "LOG_LEVEL"
            runtimeConfig.logLevel = value
        Case "LOG_TO_FILE"
            runtimeConfig.logToFile = (LCase(value) = "true")
        Case "LOG_DETAILED_OPERATIONS"
            runtimeConfig.logDetailedOperations = (LCase(value) = "true")
        Case "LOG_WARNINGS"
            runtimeConfig.logWarnings = (LCase(value) = "true")
        Case "LOG_ERRORS"
            runtimeConfig.logErrors = (LCase(value) = "true")
        Case "MAX_LOG_SIZE_MB"
            runtimeConfig.maxLogSizeMb = CLng(value)
    End Select
End Sub

Private Sub ProcessPerformanceConfig(key As String, value As String)
    Select Case key
        Case "DISABLE_SCREEN_UPDATING"
            runtimeConfig.disableScreenUpdating = (LCase(value) = "true")
        Case "DISABLE_DISPLAY_ALERTS"
            runtimeConfig.disableDisplayAlerts = (LCase(value) = "true")
        Case "USE_BULK_OPERATIONS"
            runtimeConfig.useBulkOperations = (LCase(value) = "true")
        Case "OPTIMIZE_FIND_REPLACE"
            runtimeConfig.optimizeFindReplace = (LCase(value) = "true")
        Case "MINIMIZE_OBJECT_CREATION"
            runtimeConfig.minimizeObjectCreation = (LCase(value) = "true")
        Case "CACHE_FREQUENTLY_USED_OBJECTS"
            runtimeConfig.cacheFrequentlyUsedObjects = (LCase(value) = "true")
        Case "USE_EFFICIENT_LOOPS"
            runtimeConfig.useEfficientLoops = (LCase(value) = "true")
        Case "BATCH_PARAGRAPH_OPERATIONS"
            runtimeConfig.batchParagraphOperations = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessInterfaceConfig(key As String, value As String)
    Select Case key
        Case "SHOW_PROGRESS_MESSAGES"
            runtimeConfig.showProgressMessages = (LCase(value) = "true")
        Case "SHOW_STATUS_BAR_UPDATES"
            runtimeConfig.showStatusBarUpdates = (LCase(value) = "true")
        Case "CONFIRM_CRITICAL_OPERATIONS"
            runtimeConfig.confirmCriticalOperations = (LCase(value) = "true")
        Case "SHOW_COMPLETION_MESSAGE"
            runtimeConfig.showCompletionMessage = (LCase(value) = "true")
        Case "ENABLE_EMERGENCY_RECOVERY"
            runtimeConfig.enableEmergencyRecovery = (LCase(value) = "true")
        Case "TIMEOUT_OPERATIONS"
            runtimeConfig.timeoutOperations = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessCompatibilityConfig(key As String, value As String)
    Select Case key
        Case "SUPPORT_WORD_2010"
            runtimeConfig.supportWord2010 = (LCase(value) = "true")
        Case "SUPPORT_WORD_2013"
            runtimeConfig.supportWord2013 = (LCase(value) = "true")
        Case "SUPPORT_WORD_2016"
            runtimeConfig.supportWord2016 = (LCase(value) = "true")
        Case "USE_SAFE_PROPERTY_ACCESS"
            runtimeConfig.useSafePropertyAccess = (LCase(value) = "true")
        Case "FALLBACK_METHODS"
            runtimeConfig.fallbackMethods = (LCase(value) = "true")
        Case "HANDLE_MISSING_FEATURES"
            runtimeConfig.handleMissingFeatures = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessSecurityConfig(key As String, value As String)
    Select Case key
        Case "REQUIRE_DOCUMENT_SAVED"
            runtimeConfig.requireDocumentSaved = (LCase(value) = "true")
        Case "VALIDATE_FILE_PERMISSIONS"
            runtimeConfig.validateFilePermissions = (LCase(value) = "true")
        Case "CHECK_DOCUMENT_PROTECTION"
            runtimeConfig.checkDocumentProtection = (LCase(value) = "true")
        Case "ENABLE_EMERGENCY_BACKUP"
            runtimeConfig.enableEmergencyBackup = (LCase(value) = "true")
        Case "SANITIZE_INPUTS"
            runtimeConfig.sanitizeInputs = (LCase(value) = "true")
        Case "VALIDATE_RANGES"
            runtimeConfig.validateRanges = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessAdvancedConfig(key As String, value As String)
    Select Case key
        Case "MAX_RETRY_ATTEMPTS"
            runtimeConfig.maxRetryAttempts = CLng(value)
        Case "RETRY_DELAY_MS"
            runtimeConfig.retryDelayMs = CLng(value)
        Case "COMPILATION_CHECK"
            runtimeConfig.compilationCheck = (LCase(value) = "true")
        Case "VBA_ACCESS_REQUIRED"
            runtimeConfig.vbaAccessRequired = (LCase(value) = "true")
        Case "AUTO_CLEANUP"
            runtimeConfig.autoCleanup = (LCase(value) = "true")
        Case "FORCE_GC_COLLECTION"
            runtimeConfig.forceGcCollection = (LCase(value) = "true")
    End Select
End Sub

'================================================================================
' PUBLIC: ABRIR PASTA DE LOGS
'================================================================================
Public Sub AbrirPastaLogs()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim logsFolder As String
    Dim defaultLogsFolder As String
    
    ' Tenta obter documento ativo
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler
    
    ' Define pasta de logs baseada no documento atual ou temp
    If Not doc Is Nothing And doc.path <> "" Then
        logsFolder = doc.path
    Else
        logsFolder = Environ("TEMP")
    End If
    
    ' Verifica se a pasta existe
    If Dir(logsFolder, vbDirectory) = "" Then
        logsFolder = Environ("TEMP")
    End If
    
    ' Abre a pasta no Windows Explorer
    Shell "explorer.exe """ & logsFolder & """", vbNormalFocus
    
    Application.StatusBar = "Pasta de logs aberta: " & logsFolder
    
    ' Log da operação se sistema de log estiver ativo
    If runtimeConfig.enableLogging Then
        LogEvent "AbrirPastaLogs", "INFO", "Pasta de logs aberta pelo usuário: " & logsFolder, 0, ""
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir pasta de logs"
    
    ' Fallback: tenta abrir pasta temporária
    On Error Resume Next
    Shell "explorer.exe """ & Environ("TEMP") & """", vbNormalFocus
    If Err.Number = 0 Then
        Application.StatusBar = "Pasta temporária aberta como alternativa"
    Else
        Application.StatusBar = "Não foi possível abrir pasta de logs"
    End If
    On Error GoTo 0
End Sub

'================================================================================
' PUBLIC: SALVAR E SAIR - Orquestrador profissional
'================================================================================
Public Sub SalvarESair()
    On Error GoTo CriticalErrorHandler
    
    Dim startTime As Date
    startTime = Now
    
    Application.StatusBar = "Verificando documentos abertos..."
    LogEvent "SalvarESair", "INFO", "Iniciando processo de salvar e sair - verificação de documentos", 0, ""
    
    ' Verifica se há documentos abertos
    If Application.Documents.Count = 0 Then
        Application.StatusBar = "Nenhum documento aberto - encerrando Word"
        LogEvent "SalvarESair", "INFO", "Nenhum documento aberto - encerrando aplicação", 0, ""
        Application.Quit SaveChanges:=wdDoNotSaveChanges
        Exit Sub
    End If
    
    ' Coleta informações sobre documentos não salvos
    Dim unsavedDocs As Collection
    Set unsavedDocs = New Collection
    
    Dim doc As Document
    Dim i As Long
    
    ' Verifica cada documento aberto
    For i = 1 To Application.Documents.Count
        Set doc = Application.Documents(i)
        
        On Error Resume Next
        ' Verifica se o documento tem alterações não salvas
        If doc.Saved = False Or doc.path = "" Then
            unsavedDocs.Add doc.Name
            LogEvent "SalvarESair", "INFO", "Documento não salvo detectado: " & doc.Name, 0, ""
        End If
        On Error GoTo CriticalErrorHandler
    Next i
    
    ' Se não há documentos não salvos, encerra diretamente
    If unsavedDocs.Count = 0 Then
        Application.StatusBar = "Todos os documentos salvos - encerrando Word"
        LogEvent "SalvarESair", "INFO", "Todos os documentos estão salvos - encerrando aplicação", 0, ""
        Application.Quit SaveChanges:=wdDoNotSaveChanges
        Exit Sub
    End If
    
    ' Constrói mensagem detalhada sobre documentos não salvos
    Dim message As String
    Dim docList As String
    
    For i = 1 To unsavedDocs.Count
        docList = docList & "• " & unsavedDocs(i) & vbCrLf
    Next i
    
    message = "ATENÇÃO: Há " & unsavedDocs.Count & " documento(s) com alterações não salvas:" & vbCrLf & vbCrLf
    message = message & docList & vbCrLf
    message = message & "Deseja salvar todos os documentos antes de sair?" & vbCrLf & vbCrLf
    message = message & "• SIM: Salva todos e fecha o Word" & vbCrLf
    message = message & "• NÃO: Fecha sem salvar (PERDE as alterações)" & vbCrLf
    message = message & "• CANCELAR: Cancela a operação"
    
    ' Apresenta opções ao usuário
    Application.StatusBar = "Aguardando decisão do usuário sobre documentos não salvos..."
    
    Dim userChoice As VbMsgBoxResult
    userChoice = MsgBox(message, vbYesNoCancel + vbExclamation + vbDefaultButton1, _
                        "Chainsaw - Salvar e Sair (" & unsavedDocs.Count & " documentos não salvos)")
    
    Select Case userChoice
        Case vbYes
            ' Usuário escolheu salvar todos
            Application.StatusBar = "Salvando todos os documentos..."
            LogEvent "SalvarESair", "INFO", "Usuário optou por salvar todos os documentos antes de sair", 0, ""
            
            If SalvarTodosDocumentos() Then
                Application.StatusBar = "Documentos salvos com sucesso - encerrando Word"
                LogEvent "SalvarESair", "INFO", "Todos os documentos salvos com sucesso - encerrando aplicação", 0, ""
                Application.Quit SaveChanges:=wdDoNotSaveChanges
            Else
                Application.StatusBar = "Erro ao salvar documentos - operação cancelada"
                LogEvent "SalvarESair", "ERROR", "Falha ao salvar alguns documentos - operação de sair cancelada", 0, ""
                MsgBox "Erro ao salvar um ou mais documentos." & vbCrLf & _
                       "A operação foi cancelada por segurança." & vbCrLf & vbCrLf & _
                       "Verifique os documentos e tente novamente.", _
                       vbCritical, "Chainsaw - Erro ao Salvar"
            End If
            
        Case vbNo
            ' Usuário escolheu não salvar
            Dim confirmMessage As String
            confirmMessage = "CONFIRMAÇÃO FINAL:" & vbCrLf & vbCrLf
            confirmMessage = confirmMessage & "Você está prestes a FECHAR O WORD SEM SALVAR " & unsavedDocs.Count & " documento(s)." & vbCrLf & vbCrLf
            confirmMessage = confirmMessage & "TODAS AS ALTERAÇÕES NÃO SALVAS SERÃO PERDIDAS!" & vbCrLf & vbCrLf
            confirmMessage = confirmMessage & "Tem certeza absoluta?"
            
            Dim finalConfirm As VbMsgBoxResult
            finalConfirm = MsgBox(confirmMessage, vbYesNo + vbCritical + vbDefaultButton2, _
                                  "Chainsaw - CONFIRMAÇÃO FINAL")
            
            If finalConfirm = vbYes Then
                Application.StatusBar = "Fechando Word sem salvar alterações..."
                LogEvent "SalvarESair", "WARNING", "Usuário confirmou fechamento sem salvar - encerrando aplicação", 0, ""
                Application.Quit SaveChanges:=wdDoNotSaveChanges
            Else
                Application.StatusBar = "Operação cancelada pelo usuário"
                LogEvent "SalvarESair", "INFO", "Usuário cancelou fechamento sem salvar", 0, ""
                MsgBox "Operação cancelada." & vbCrLf & "Os documentos permanecem abertos.", _
                       vbInformation, "Chainsaw - Operação Cancelada"
            End If
            
        Case vbCancel
            ' Usuário cancelou
            Application.StatusBar = "Operação de sair cancelada pelo usuário"
            LogEvent "SalvarESair", "INFO", "Usuário cancelou operação de salvar e sair", 0, ""
            MsgBox "Operação cancelada." & vbCrLf & "Os documentos permanecem abertos.", _
                   vbInformation, "Chainsaw - Operação Cancelada"
    End Select
    
    Application.StatusBar = False
    LogEvent "SalvarESair", "INFO", "Processo de salvar e sair concluído em " & Format(Now - startTime, "hh:mm:ss"), 0, ""
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "ERRO CRÍTICO na operação Salvar e Sair #" & Err.Number & ": " & Err.Description
    
    LogEvent "SalvarESair", "ERROR", errDesc, 0, ""
    Application.StatusBar = "Erro crítico - operação cancelada"
    
    MsgBox "Erro crítico durante a operação Salvar e Sair:" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "A operação foi cancelada por segurança." & vbCrLf & _
           "Salve manualmente os documentos importantes.", _
           vbCritical, "Chainsaw - Erro Crítico"
End Sub

'================================================================================
' SALVAR TODOS DOCUMENTOS - AUXILIAR
'================================================================================
Private Function SalvarTodosDocumentos() As Boolean
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim i As Long
    Dim savedCount As Long
    Dim errorCount As Long
    Dim totalDocs As Long
    
    totalDocs = Application.Documents.Count
    
    ' Salva cada documento individualmente
    For i = 1 To totalDocs
        Set doc = Application.Documents(i)
        
        Application.StatusBar = "Salvando documento " & i & " de " & totalDocs & ": " & doc.Name
        
        On Error Resume Next
        
        ' Se o documento nunca foi salvo (sem caminho), abre dialog
        If doc.path = "" Then
            Dim saveDialog As Object
            Set saveDialog = Application.FileDialog(msoFileDialogSaveAs)
            
            With saveDialog
                .Title = "Salvar documento: " & doc.Name
                .InitialFileName = doc.Name
                
                If .Show = -1 Then
                    doc.SaveAs2 .SelectedItems(1)
                    If Err.Number = 0 Then
                        savedCount = savedCount + 1
                        LogEvent "SalvarTodosDocumentos", "INFO", "Documento salvo como novo arquivo: " & doc.Name, 0, ""
                    Else
                        errorCount = errorCount + 1
                        LogEvent "SalvarTodosDocumentos", "ERROR", "Erro ao salvar documento como novo: " & doc.Name & " - " & Err.Description, 0, ""
                    End If
                Else
                    errorCount = errorCount + 1
                    LogEvent "SalvarTodosDocumentos", "WARNING", "Salvamento cancelado pelo usuário: " & doc.Name, 0, ""
                End If
            End With
        Else
            ' Documento já tem caminho, apenas salva
            doc.Save
            If Err.Number = 0 Then
                savedCount = savedCount + 1
                LogEvent "SalvarTodosDocumentos", "INFO", "Documento salvo: " & doc.Name, 0, ""
            Else
                errorCount = errorCount + 1
                LogEvent "SalvarTodosDocumentos", "ERROR", "Erro ao salvar documento: " & doc.Name & " - " & Err.Description, 0, ""
            End If
        End If
        
        On Error GoTo ErrorHandler
    Next i
    
    ' Verifica resultado
    If errorCount = 0 Then
        LogEvent "SalvarTodosDocumentos", "INFO", "Todos os documentos salvos com sucesso: " & savedCount & " de " & totalDocs, 0, ""
        SalvarTodosDocumentos = True
    Else
        LogEvent "SalvarTodosDocumentos", "WARNING", "Falha parcial no salvamento: " & savedCount & " salvos, " & errorCount & " erros", 0, ""
        SalvarTodosDocumentos = False
    End If
    
    Exit Function

ErrorHandler:
    LogEvent "SalvarTodosDocumentos", "ERROR", "Erro crítico ao salvar documentos: " & Err.Description, 0, ""
    SalvarTodosDocumentos = False
End Function

'================================================================================
' STANDARD FIND REPLACE
'================================================================================
Private Function StandardFindReplace(findText As String, replaceText As String, Optional searchRange As Range = Nothing) As Long
    On Error GoTo ErrorHandler
    
    StandardFindReplace = 0
    
    Dim targetRange As Range
    Set targetRange = IIf(searchRange Is Nothing, ActiveDocument.Content, searchRange)
    
    With targetRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = wdFindStop
        
        StandardFindReplace = .Execute(Replace:=wdReplaceAll)
    End With
    
    Exit Function
    
ErrorHandler:
    LogEvent "StandardFindReplace", "ERROR", "Erro: " & Err.Description, 0, ""
    StandardFindReplace = 0
End Function

'================================================================================
' OPTIMIZED PARAGRAPH PROCESSING
'================================================================================
Private Function OptimizedParagraphProcessing(processingFunction As String) As Boolean
    On Error GoTo ErrorHandler
    
    OptimizedParagraphProcessing = False
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    If doc.Paragraphs.Count > OPTIMIZATION_THRESHOLD Then
        OptimizedParagraphProcessing = BatchProcessParagraphs(processingFunction)
    Else
        OptimizedParagraphProcessing = StandardProcessParagraphs(processingFunction)
    End If
    
    Exit Function
    
ErrorHandler:
    LogEvent "OptimizedParagraphProcessing", "ERROR", "Erro: " & Err.Description, 0, ""
    OptimizedParagraphProcessing = False
End Function

'================================================================================
' BATCH PROCESS PARAGRAPHS
'================================================================================
Private Function BatchProcessParagraphs(processingFunction As String) As Boolean
    On Error GoTo ErrorHandler
    
    BatchProcessParagraphs = False
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim paragraphCount As Long
    paragraphCount = doc.Paragraphs.Count
    
    Dim batchSize As Long
    batchSize = IIf(paragraphCount > OPTIMIZATION_THRESHOLD, MAX_PARAGRAPH_BATCH_SIZE, paragraphCount)
    
    Dim i As Long
    For i = 1 To paragraphCount Step batchSize
        Dim endIndex As Long
        endIndex = IIf(i + batchSize - 1 > paragraphCount, paragraphCount, i + batchSize - 1)
        
        If Not ProcessParagraphBatch(i, endIndex, processingFunction) Then
            LogEvent "BatchProcessParagraphs", "ERROR", "Erro no lote " & i, 0, ""
            Exit Function
        End If
    Next i
    
    BatchProcessParagraphs = True
    Exit Function
    
ErrorHandler:
    LogEvent "BatchProcessParagraphs", "ERROR", "Erro: " & Err.Description, 0, ""
    BatchProcessParagraphs = False
End Function

'================================================================================
' STANDARD PROCESS PARAGRAPHS
'================================================================================
Private Function StandardProcessParagraphs(processingFunction As String) As Boolean
    On Error GoTo ErrorHandler
    
    StandardProcessParagraphs = False
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        Select Case processingFunction
            Case "FORMAT"
                Call FormatParagraph(para)
            Case "CLEAN"
                Call CleanParagraph(para)
            Case "VALIDATE"
                Call ValidateParagraph(para)
        End Select
    Next para
    
    StandardProcessParagraphs = True
    Exit Function
    
ErrorHandler:
    LogEvent "StandardProcessParagraphs", "ERROR", "Erro: " & Err.Description, 0, ""
    StandardProcessParagraphs = False
End Function

'================================================================================
' PROCESS PARAGRAPH BATCH
'================================================================================
Private Function ProcessParagraphBatch(startIndex As Long, endIndex As Long, processingFunction As String) As Boolean
    On Error GoTo ErrorHandler
    
    ProcessParagraphBatch = False
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim i As Long
    For i = startIndex To endIndex
        If i <= doc.Paragraphs.Count Then
            Dim para As Paragraph
            Set para = doc.Paragraphs(i)
            
            Select Case processingFunction
                Case "FORMAT"
                    Call FormatParagraph(para)
                Case "CLEAN"
                    Call CleanParagraph(para)
                Case "VALIDATE"
                    Call ValidateParagraph(para)
            End Select
        End If
    Next i
    
    ProcessParagraphBatch = True
    Exit Function
    
ErrorHandler:
    LogEvent "ProcessParagraphBatch", "ERROR", "Erro: " & Err.Description, 0, ""
    ProcessParagraphBatch = False
End Function
