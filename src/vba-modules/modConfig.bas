' SPDX-License-Identifier: GPL-3.0-or-later
' =============================================================================
' MODULE: modConfig
' PURPOSE: Holds configuration types, globals, and parsing logic extracted
'          from the original monolithic module for cleaner modular architecture.
' =============================================================================
Option Explicit

' Public flag enabling ASCII-safe normalization of dialog strings
Public dialogAsciiNormalizationEnabled As Boolean

' Configuration structure (was Private in monolith)
Public Type ConfigSettings
    ' General
    debugMode As Boolean
    performanceMode As Boolean
    compatibilityMode As Boolean
    
    ' Validations
    CheckWordVersion As Boolean
    ValidateDocumentIntegrity As Boolean
    ValidatePropositionType As Boolean
    ValidateContentConsistency As Boolean
    CheckDiskSpace As Boolean
    minWordVersion As Double
    maxDocumentSize As Long
    
    ' Formatting
    ApplyPageSetup As Boolean
    applyStandardFont As Boolean
    applyStandardParagraphs As Boolean
    FormatFirstParagraph As Boolean
    FormatSecondParagraph As Boolean
    FormatNumberedParagraphs As Boolean
    FormatConsiderandoParagraphs As Boolean
    formatJustificativaParagraphs As Boolean
    EnableHyphenation As Boolean
    
    ' Cleaning
    CleanDocumentStructure As Boolean
    CleanMultipleSpaces As Boolean
    LimitSequentialEmptyLines As Boolean
    EnsureParagraphSeparation As Boolean
    cleanVisualElements As Boolean
    deleteHiddenElements As Boolean
    deleteVisualElementsFirstFourParagraphs As Boolean
    
    ' Header/Footer
    InsertHeaderstamp As Boolean
    InsertFooterstamp As Boolean
    RemoveWatermark As Boolean
    headerImagePath As String
    
    ' Text Replacements
    ApplyTextReplacements As Boolean
    ApplySpecificParagraphReplacements As Boolean
    replaceHyphensWithEmDash As Boolean
    removeManualLineBreaks As Boolean
    normalizeDosteVariants As Boolean
    
    ' Performance
    disableScreenUpdating As Boolean
    disableDisplayAlerts As Boolean
    useBulkOperations As Boolean
    optimizeFindReplace As Boolean
    
    ' Interface
    showProgressMessages As Boolean
    showStatusBarUpdates As Boolean
    confirmCriticalOperations As Boolean
    showCompletionMessage As Boolean
    enableEmergencyRecovery As Boolean
    timeoutOperations As Boolean
    
    ' Compatibility
    supportWord2010 As Boolean
    supportWord2013 As Boolean
    supportWord2016 As Boolean
    useSafePropertyAccess As Boolean
    fallbackMethods As Boolean
    handleMissingFeatures As Boolean
    
    ' Security
    requireDocumentSaved As Boolean
    validateFilePermissions As Boolean
    checkDocumentProtection As Boolean
    sanitizeInputs As Boolean
    validateRanges As Boolean
    
    ' Advanced
    maxRetryAttempts As Long
    retryDelayMs As Long
    autoRunSmokeTest As Boolean
End Type

' Active configuration instance (now public)
Public Config As ConfigSettings

' Configuration file constants (replicated locally to avoid cross-module hidden deps)
Private Const CONFIG_FILE_NAME As String = "chainsaw-config.ini"
Private Const CONFIG_FILE_PATH As String = "\chainsaw\"

' Loads configuration (defaults + file overrides). Public for orchestrator pipeline.
Public Function LoadConfiguration() As Boolean
    On Error GoTo ErrorHandler
    LoadConfiguration = False

    SetDefaultConfiguration
    dialogAsciiNormalizationEnabled = True ' default on for safe ASCII dialogs

    Dim configPath As String
    configPath = GetConfigurationFilePath()

    If Len(configPath) = 0 Or Dir(configPath) = "" Then
        LoadConfiguration = True ' Use defaults silently
        Exit Function
    End If

    If ParseConfigurationFile(configPath) Then
        LoadConfiguration = True
    Else
        SetDefaultConfiguration
        LoadConfiguration = True ' Fallback to defaults on error
    End If
    Exit Function
ErrorHandler:
    LoadConfiguration = True ' Fail-soft: defaults already applied
End Function

Public Function GetConfigurationFilePath() As String
    On Error GoTo ErrorHandler
    Dim doc As Document
    Dim basePath As String
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    If Not doc Is Nothing And doc.Path <> "" Then
        basePath = doc.Path
    Else
        basePath = Environ("USERPROFILE") & "\Documents"
    End If
    On Error GoTo ErrorHandler
    GetConfigurationFilePath = basePath & CONFIG_FILE_PATH & CONFIG_FILE_NAME
    Exit Function
ErrorHandler:
    GetConfigurationFilePath = ""
End Function

Public Sub SetDefaultConfiguration()
    With Config
        ' General
        .debugMode = False
        .performanceMode = True
        .compatibilityMode = True
        
        ' Validations
        .CheckWordVersion = True
        .ValidateDocumentIntegrity = True
        .ValidatePropositionType = True
        .ValidateContentConsistency = True
        .CheckDiskSpace = True
        .minWordVersion = 14#
        .maxDocumentSize = 500000
        
        ' Formatting
        .ApplyPageSetup = True
        .applyStandardFont = True
        .applyStandardParagraphs = True
        .FormatFirstParagraph = True
        .FormatSecondParagraph = True
        .FormatNumberedParagraphs = True
        .FormatConsiderandoParagraphs = True
        .formatJustificativaParagraphs = True
        .EnableHyphenation = True
        
        ' Cleaning
        .CleanDocumentStructure = True
        .CleanMultipleSpaces = True
        .LimitSequentialEmptyLines = True
        .EnsureParagraphSeparation = True
        .cleanVisualElements = True
        .deleteHiddenElements = True
        .deleteVisualElementsFirstFourParagraphs = True
        
        ' Header/Footer
        .InsertHeaderstamp = True
        .InsertFooterstamp = True
        .RemoveWatermark = True
        .headerImagePath = "assets\stamp.png"
        
        ' Text Replacements
        .ApplyTextReplacements = True
        .ApplySpecificParagraphReplacements = True
        .replaceHyphensWithEmDash = True
        .removeManualLineBreaks = True
        .normalizeDosteVariants = True
        
        ' Performance
        .disableScreenUpdating = True
        .disableDisplayAlerts = True
        .useBulkOperations = True
        .optimizeFindReplace = True
        .showCompletionMessage = True
        .enableEmergencyRecovery = True
        .timeoutOperations = True
        
        ' Compatibility
        .supportWord2010 = True
        .supportWord2013 = True
        .supportWord2016 = True
        .useSafePropertyAccess = True
        .fallbackMethods = True
        .handleMissingFeatures = True
        
        ' Security
        .requireDocumentSaved = True
        .validateFilePermissions = True
        .checkDocumentProtection = True
        .sanitizeInputs = True
        .validateRanges = True
        
        ' Advanced
        .maxRetryAttempts = 3
        .retryDelayMs = 1000
        .autoRunSmokeTest = False
    End With
End Sub

Public Function ParseConfigurationFile(configPath As String) As Boolean
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
        If Len(fileLine) > 0 And Left(fileLine, 1) <> "#" Then
            If Left(fileLine, 1) = "[" And Right(fileLine, 1) = "]" Then
                currentSection = UCase(Mid(fileLine, 2, Len(fileLine) - 2))
            ElseIf InStr(fileLine, "=") > 0 Then
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

Public Sub ProcessConfigLine(section As String, configLine As String)
    On Error Resume Next
    Dim equalPos As Integer
    Dim configKey As String
    Dim configValue As String
    equalPos = InStr(configLine, "=")
    If equalPos > 0 Then
        configKey = UCase(Trim(Left(configLine, equalPos - 1)))
        configValue = Trim(Mid(configLine, equalPos + 1))
        If Left(configValue, 1) = """" And Right(configValue, 1) = """" Then
            configValue = Mid(configValue, 2, Len(configValue) - 2)
        End If
        Select Case section
            Case "GERAL", "GENERAL":        ProcessGeneralConfig configKey, configValue
            Case "VALIDACOES", "VALIDATIONS": ProcessValidationConfig configKey, configValue
            Case "FORMATACAO", "FORMATTING": ProcessFormattingConfig configKey, configValue
            Case "LIMPEZA", "CLEANUP":       ProcessCleaningConfig configKey, configValue
            Case "CABECALHO_RODAPE", "HEADER_FOOTER": ProcessHeaderFooterConfig configKey, configValue
            Case "SUBSTITUICOES", "REPLACEMENTS": ProcessReplacementConfig configKey, configValue
            Case "PERFORMANCE":             ProcessPerformanceConfig configKey, configValue
            Case "INTERFACE":               ProcessInterfaceConfig configKey, configValue
            Case "COMPATIBILIDADE", "COMPATIBILITY": ProcessCompatibilityConfig configKey, configValue
            Case "SEGURANCA", "SECURITY":   ProcessSecurityConfig configKey, configValue
            Case "AVANCADO", "ADVANCED":    ProcessAdvancedConfig configKey, configValue
            Case Else
                ' Deprecated sections ignored silently
        End Select
    End If
End Sub

Public Sub ProcessGeneralConfig(key As String, value As String)
    Select Case key
        Case "DEBUG_MODE":          Config.debugMode = (LCase(value) = "true")
        Case "PERFORMANCE_MODE":    Config.performanceMode = (LCase(value) = "true")
        Case "COMPATIBILITY_MODE":  Config.compatibilityMode = (LCase(value) = "true")
        Case "AUTO_RUN_SMOKE_TEST": Config.autoRunSmokeTest = (LCase(value) = "true")
    End Select
End Sub

Public Sub ProcessValidationConfig(key As String, value As String)
    Select Case key
        Case "CHECK_WORD_VERSION":            Config.CheckWordVersion = (LCase(value) = "true")
        Case "VALIDATE_DOCUMENT_INTEGRITY":   Config.ValidateDocumentIntegrity = (LCase(value) = "true")
        Case "VALIDATE_PROPOSITION_TYPE":     Config.ValidatePropositionType = (LCase(value) = "true")
        Case "VALIDATE_CONTENT_CONSISTENCY":  Config.ValidateContentConsistency = (LCase(value) = "true")
        Case "CHECK_DISK_SPACE":              Config.CheckDiskSpace = (LCase(value) = "true")
        Case "MIN_WORD_VERSION":              Config.minWordVersion = CDbl(value)
        Case "MAX_DOCUMENT_SIZE":             Config.maxDocumentSize = CLng(value)
    End Select
End Sub

Public Sub ProcessFormattingConfig(key As String, value As String)
    Select Case key
        Case "APPLY_PAGE_SETUP":               Config.ApplyPageSetup = (LCase(value) = "true")
        Case "APPLY_STANDARD_FONT":            Config.applyStandardFont = (LCase(value) = "true")
        Case "APPLY_STANDARD_PARAGRAPHS":      Config.applyStandardParagraphs = (LCase(value) = "true")
        Case "FORMAT_FIRST_PARAGRAPH":         Config.FormatFirstParagraph = (LCase(value) = "true")
        Case "FORMAT_SECOND_PARAGRAPH":        Config.FormatSecondParagraph = (LCase(value) = "true")
        Case "FORMAT_NUMBERED_PARAGRAPHS":     Config.FormatNumberedParagraphs = (LCase(value) = "true")
        Case "FORMAT_CONSIDERANDO_PARAGRAPHS": Config.FormatConsiderandoParagraphs = (LCase(value) = "true")
        Case "FORMAT_JUSTIFICATIVA_PARAGRAPHS":Config.formatJustificativaParagraphs = (LCase(value) = "true")
        Case "ENABLE_HYPHENATION":             Config.EnableHyphenation = (LCase(value) = "true")
    End Select
End Sub

Public Sub ProcessCleaningConfig(key As String, value As String)
    Select Case key
        Case "CLEAN_DOCUMENT_STRUCTURE":                 Config.CleanDocumentStructure = (LCase(value) = "true")
        Case "CLEAN_MULTIPLE_SPACES":                    Config.CleanMultipleSpaces = (LCase(value) = "true")
        Case "LIMIT_SEQUENTIAL_EMPTY_LINES":             Config.LimitSequentialEmptyLines = (LCase(value) = "true")
        Case "ENSURE_PARAGRAPH_SEPARATION":              Config.EnsureParagraphSeparation = (LCase(value) = "true")
        Case "CLEAN_VISUAL_ELEMENTS":                    Config.cleanVisualElements = (LCase(value) = "true")
        Case "DELETE_HIDDEN_ELEMENTS":                   Config.deleteHiddenElements = (LCase(value) = "true")
        Case "DELETE_VISUAL_ELEMENTS_FIRST_FOUR_PARAGRAPHS": Config.deleteVisualElementsFirstFourParagraphs = (LCase(value) = "true")
    End Select
End Sub

Public Sub ProcessHeaderFooterConfig(key As String, value As String)
    Select Case key
        Case "INSERT_HEADER_STAMP": Config.InsertHeaderstamp = (LCase(value) = "true")
        Case "INSERT_FOOTER_STAMP": Config.InsertFooterstamp = (LCase(value) = "true")
        Case "REMOVE_WATERMARK":    Config.RemoveWatermark = (LCase(value) = "true")
        Case "HEADER_IMAGE_PATH":   Config.headerImagePath = value
    End Select
End Sub

Public Sub ProcessReplacementConfig(key As String, value As String)
    Select Case key
        Case "APPLY_TEXT_REPLACEMENTS":             Config.ApplyTextReplacements = (LCase(value) = "true")
        Case "APPLY_SPECIFIC_PARAGRAPH_REPLACEMENTS": Config.ApplySpecificParagraphReplacements = (LCase(value) = "true")
        Case "REPLACE_HYPHENS_WITH_EM_DASH":        Config.replaceHyphensWithEmDash = (LCase(value) = "true")
        Case "REMOVE_MANUAL_LINE_BREAKS":           Config.removeManualLineBreaks = (LCase(value) = "true")
        Case "NORMALIZE_DOESTE_VARIANTS":           Config.normalizeDosteVariants = (LCase(value) = "true")
    End Select
End Sub

Public Sub ProcessPerformanceConfig(key As String, value As String)
    Select Case key
        Case "DISABLE_SCREEN_UPDATING": Config.disableScreenUpdating = (LCase(value) = "true")
        Case "DISABLE_DISPLAY_ALERTS":  Config.disableDisplayAlerts = (LCase(value) = "true")
        Case "USE_BULK_OPERATIONS":     Config.useBulkOperations = (LCase(value) = "true")
        Case "OPTIMIZE_FIND_REPLACE":   Config.optimizeFindReplace = (LCase(value) = "true")
    End Select
End Sub

Public Sub ProcessInterfaceConfig(key As String, value As String)
    Select Case key
        Case "SHOW_PROGRESS_MESSAGES":      Config.showProgressMessages = (LCase(value) = "true")
        Case "SHOW_STATUS_BAR_UPDATES":     Config.showStatusBarUpdates = (LCase(value) = "true")
        Case "CONFIRM_CRITICAL_OPERATIONS": Config.confirmCriticalOperations = (LCase(value) = "true")
        Case "SHOW_COMPLETION_MESSAGE":     Config.showCompletionMessage = (LCase(value) = "true")
        Case "ENABLE_EMERGENCY_RECOVERY":   Config.enableEmergencyRecovery = (LCase(value) = "true")
        Case "TIMEOUT_OPERATIONS":          Config.timeoutOperations = (LCase(value) = "true")
        Case "DIALOG_ASCII_NORMALIZATION", "DIALOG_ASCII_NORMALIZE", "ASCII_DIALOGS": dialogAsciiNormalizationEnabled = (LCase(value) <> "false" And LCase(value) <> "0")
    End Select
End Sub

Public Sub ProcessCompatibilityConfig(key As String, value As String)
    Select Case key
        Case "SUPPORT_WORD_2010":        Config.supportWord2010 = (LCase(value) = "true")
        Case "SUPPORT_WORD_2013":        Config.supportWord2013 = (LCase(value) = "true")
        Case "SUPPORT_WORD_2016":        Config.supportWord2016 = (LCase(value) = "true")
        Case "USE_SAFE_PROPERTY_ACCESS": Config.useSafePropertyAccess = (LCase(value) = "true")
        Case "FALLBACK_METHODS":         Config.fallbackMethods = (LCase(value) = "true")
        Case "HANDLE_MISSING_FEATURES":  Config.handleMissingFeatures = (LCase(value) = "true")
    End Select
End Sub

Public Sub ProcessSecurityConfig(key As String, value As String)
    Select Case key
        Case "REQUIRE_DOCUMENT_SAVED":   Config.requireDocumentSaved = (LCase(value) = "true")
        Case "VALIDATE_FILE_PERMISSIONS": Config.validateFilePermissions = (LCase(value) = "true")
        Case "CHECK_DOCUMENT_PROTECTION": Config.checkDocumentProtection = (LCase(value) = "true")
        Case "SANITIZE_INPUTS":          Config.sanitizeInputs = (LCase(value) = "true")
        Case "VALIDATE_RANGES":          Config.validateRanges = (LCase(value) = "true")
    End Select
End Sub

Public Sub ProcessAdvancedConfig(key As String, value As String)
    Select Case key
        Case "MAX_RETRY_ATTEMPTS": Config.maxRetryAttempts = CLng(value)
        Case "RETRY_DELAY_MS":     Config.retryDelayMs = CLng(value)
    End Select
End Sub
