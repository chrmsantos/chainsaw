'================================================================================
' MODULE: modConfig
' PURPOSE: Configuration loading and parsing for Chainsaw Proposituras.
' NOTE: Extracted from monolithic chainsaw.bas on 2025-10-06.
'================================================================================
Option Explicit

Public Type ConfigSettings
    ' --- CORE EXECUTION MODES ---
    debugMode As Boolean
    performanceMode As Boolean
    compatibilityMode As Boolean
    CheckWordVersion As Boolean
    ValidateDocumentIntegrity As Boolean
    ValidatePropositionType As Boolean
    ValidateContentConsistency As Boolean
    CheckDiskSpace As Boolean
    minWordVersion As Double
    maxDocumentSize As Long
    ' --- BACKUP SYSTEM (DEPRECATED / DISABLED IN CURRENT BETA) ---
    ' Retained for forward compatibility; currently ignored by processing pipeline.
    autoBackup As Boolean              ' Deprecated: no backups are performed
    backupBeforeProcessing As Boolean  ' Deprecated
    maxBackupFiles As Long             ' Deprecated
    backupCleanup As Boolean           ' Deprecated
    backupRetryAttempts As Long        ' Deprecated
    ApplyPageSetup As Boolean
    applyStandardFont As Boolean
    applyStandardParagraphs As Boolean
    FormatFirstParagraph As Boolean
    FormatSecondParagraph As Boolean
    FormatNumberedParagraphs As Boolean
    FormatConsiderandoParagraphs As Boolean
    formatJustificativaParagraphs As Boolean
    EnableHyphenation As Boolean
    CleanDocumentStructure As Boolean
    CleanMultipleSpaces As Boolean
    LimitSequentialEmptyLines As Boolean
    EnsureParagraphSeparation As Boolean
    cleanVisualElements As Boolean
    deleteHiddenElements As Boolean
    deleteVisualElementsFirstFourParagraphs As Boolean
    InsertHeaderstamp As Boolean
    InsertFooterstamp As Boolean
    RemoveWatermark As Boolean
    headerImagePath As String
    ApplyTextReplacements As Boolean
    ApplySpecificParagraphReplacements As Boolean
    replaceHyphensWithEmDash As Boolean
    removeManualLineBreaks As Boolean
    normalizeDosteVariants As Boolean
    BackupAllImages As Boolean
    RestoreAllImages As Boolean
    ProtectImagesInRange As Boolean
    BackupViewSettings As Boolean
    RestoreViewSettings As Boolean
    ' --- LOGGING (STUB ONLY IN CURRENT BETA) ---
    enableLogging As Boolean           ' Deprecated: logging routines are no-op
    logLevel As String                 ' Reserved for future reinstatement
    logToFile As Boolean               ' Deprecated
    maxLogSizeMb As Long               ' Deprecated
    disableScreenUpdating As Boolean
    disableDisplayAlerts As Boolean
    useBulkOperations As Boolean
    optimizeFindReplace As Boolean
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
End Type

Public Config As ConfigSettings

Private Const CONFIG_FILE_NAME As String = "chainsaw-config.ini"
Private Const CONFIG_FILE_PATH As String = "\chainsaw\"

Public Function modConfig_LoadConfiguration() As Boolean
    On Error GoTo FailSafe
    modConfig_LoadConfiguration = False
    SetDefaultConfiguration
    Dim configPath As String
    configPath = GetConfigurationFilePath()
    If Len(configPath) = 0 Or Dir(configPath) = "" Then
        modConfig_LoadConfiguration = True
        Exit Function
    End If
    If ParseConfigurationFile(configPath) Then
        modConfig_LoadConfiguration = True
    Else
        SetDefaultConfiguration
        modConfig_LoadConfiguration = True
    End If
    Exit Function
FailSafe:
    SetDefaultConfiguration
    modConfig_LoadConfiguration = True
End Function

Public Function modConfig_LoadConfigIfNeeded(ByRef isLoaded As Boolean) As Boolean
    If Not isLoaded Then
        isLoaded = modConfig_LoadConfiguration()
    End If
    modConfig_LoadConfigIfNeeded = isLoaded
End Function

Private Function GetConfigurationFilePath() As String
    On Error GoTo ErrHandler
    Dim doc As Document, basePath As String
    Set doc = Nothing
    On Error Resume Next: Set doc = ActiveDocument: On Error GoTo ErrHandler
    If Not doc Is Nothing And doc.Path <> "" Then
        basePath = doc.Path
    Else
        basePath = Environ("USERPROFILE") & "\Documents"
    End If
    If Right(basePath, 1) <> "\" Then basePath = basePath & "\"
    GetConfigurationFilePath = basePath & CONFIG_FILE_PATH & CONFIG_FILE_NAME
    Exit Function
ErrHandler:
    GetConfigurationFilePath = ""
End Function

Private Sub SetDefaultConfiguration()
    With Config
        .debugMode = False: .performanceMode = True: .compatibilityMode = True
        .CheckWordVersion = True: .ValidateDocumentIntegrity = True: .ValidatePropositionType = True: .ValidateContentConsistency = True: .CheckDiskSpace = True
        .minWordVersion = 14#: .maxDocumentSize = 500000
    ' Backup defaults forced to disabled (feature deprecated this beta)
    .autoBackup = False: .backupBeforeProcessing = False: .maxBackupFiles = 0: .backupCleanup = False: .backupRetryAttempts = 0
        .ApplyPageSetup = True: .applyStandardFont = True: .applyStandardParagraphs = True: .FormatFirstParagraph = True: .FormatSecondParagraph = True
        .FormatNumberedParagraphs = True: .FormatConsiderandoParagraphs = True: .formatJustificativaParagraphs = True: .EnableHyphenation = True
        .CleanDocumentStructure = True: .CleanMultipleSpaces = True: .LimitSequentialEmptyLines = True: .EnsureParagraphSeparation = True
        .cleanVisualElements = True: .deleteHiddenElements = True: .deleteVisualElementsFirstFourParagraphs = True
        .InsertHeaderstamp = True: .InsertFooterstamp = True: .RemoveWatermark = True: .headerImagePath = ""
        .ApplyTextReplacements = True: .ApplySpecificParagraphReplacements = True: .replaceHyphensWithEmDash = True: .removeManualLineBreaks = True
        .normalizeDosteVariants = True
    .BackupAllImages = True: .RestoreAllImages = True: .ProtectImagesInRange = True: .BackupViewSettings = True: .RestoreViewSettings = True
    ' Logging disabled; keep defaults minimal
    .enableLogging = False: .logLevel = "INFO": .logToFile = False: .maxLogSizeMb = 0
        .disableScreenUpdating = True: .disableDisplayAlerts = True: .useBulkOperations = True: .optimizeFindReplace = True
        .showProgressMessages = True: .showStatusBarUpdates = True: .confirmCriticalOperations = True: .showCompletionMessage = True
        .enableEmergencyRecovery = True: .timeoutOperations = False
        .supportWord2010 = True: .supportWord2013 = True: .supportWord2016 = True: .useSafePropertyAccess = True: .fallbackMethods = True: .handleMissingFeatures = True
        .requireDocumentSaved = True: .validateFilePermissions = True: .checkDocumentProtection = True: .enableEmergencyBackup = True
        .sanitizeInputs = True: .validateRanges = True
        .maxRetryAttempts = 3: .retryDelayMs = 250
    End With
End Sub

Private Function ParseConfigurationFile(configPath As String) As Boolean
    On Error GoTo ErrHandler
    Dim f As Integer, line As String, key As String, value As String, pos As Long
    ParseConfigurationFile = False
    f = FreeFile
    Open configPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, line
        line = Trim(line)
        If Len(line) > 0 And Left(line, 1) <> "#" And Left(line, 2) <> "//" Then
            pos = InStr(line, "=")
            If pos > 0 Then
                key = Trim(Left(line, pos - 1))
                value = Trim(Mid(line, pos + 1))
                ApplyConfigurationKey key, value
            End If
        End If
    Loop
    Close #f
    ParseConfigurationFile = True
    Exit Function
ErrHandler:
    On Error Resume Next: Close #f
    ParseConfigurationFile = False
End Function

Private Sub ApplyConfigurationKey(key As String, value As String)
    On Error Resume Next
    Select Case LCase(key)
        Case "debugmode": Config.debugMode = CBool(value)
        Case "performancemode": Config.performanceMode = CBool(value)
        Case "compatibilitymode": Config.compatibilityMode = CBool(value)
        Case "checkwordversion": Config.CheckWordVersion = CBool(value)
        Case "validatedocumentintegrity": Config.ValidateDocumentIntegrity = CBool(value)
        Case "validatepropositiontype": Config.ValidatePropositionType = CBool(value)
        Case "validatecontentconsistency": Config.ValidateContentConsistency = CBool(value)
        Case "checkdiskspace": Config.CheckDiskSpace = CBool(value)
        Case "minwordversion": Config.minWordVersion = CDbl(value)
        Case "maxdocumentsize": Config.maxDocumentSize = CLng(value)
        Case "autobackup": Config.autoBackup = CBool(value)
        Case "backupbeforeprocessing": Config.backupBeforeProcessing = CBool(value)
        Case "maxbackupfiles": Config.maxBackupFiles = CLng(value)
        Case "backupcleanup": Config.backupCleanup = CBool(value)
        Case "backupretryattempts": Config.backupRetryAttempts = CLng(value)
        Case "applypagesetup": Config.ApplyPageSetup = CBool(value)
        Case "applystandardfont": Config.applyStandardFont = CBool(value)
        Case "applystandardparagraphs": Config.applyStandardParagraphs = CBool(value)
        Case "formatfirstparagraph": Config.FormatFirstParagraph = CBool(value)
        Case "formatsecondparagraph": Config.FormatSecondParagraph = CBool(value)
        Case "formatnumberedparagraphs": Config.FormatNumberedParagraphs = CBool(value)
        Case "formatconsiderandoparagraphs": Config.FormatConsiderandoParagraphs = CBool(value)
        Case "formatjustificativaparagraphs": Config.formatJustificativaParagraphs = CBool(value)
        Case "enablehyphenation": Config.EnableHyphenation = CBool(value)
        Case "cleandocumentstructure": Config.CleanDocumentStructure = CBool(value)
        Case "cleanmultiplespaces": Config.CleanMultipleSpaces = CBool(value)
        Case "limitsequentialemptylines": Config.LimitSequentialEmptyLines = CBool(value)
        Case "ensureparagraphseparation": Config.EnsureParagraphSeparation = CBool(value)
        Case "cleanvisualelements": Config.cleanVisualElements = CBool(value)
        Case "deletehiddenelements": Config.deleteHiddenElements = CBool(value)
        Case "deletevisualelementsfirstfourparagraphs": Config.deleteVisualElementsFirstFourParagraphs = CBool(value)
        Case "insertheaderstamp": Config.InsertHeaderstamp = CBool(value)
        Case "insertfooterstamp": Config.InsertFooterstamp = CBool(value)
        Case "removewatermark": Config.RemoveWatermark = CBool(value)
        Case "headerimagepath": Config.headerImagePath = value
        Case "applytextreplacements": Config.ApplyTextReplacements = CBool(value)
        Case "applyspecificparagraphreplacements": Config.ApplySpecificParagraphReplacements = CBool(value)
        Case "replacehyphenswithemdash": Config.replaceHyphensWithEmDash = CBool(value)
        Case "removemanuallinebreaks": Config.removeManualLineBreaks = CBool(value)
        Case "normalizedostevariants": Config.normalizeDosteVariants = CBool(value)
        Case "backupallimages": Config.BackupAllImages = CBool(value)
        Case "restoreallimages": Config.RestoreAllImages = CBool(value)
        Case "protectimagesinrange": Config.ProtectImagesInRange = CBool(value)
        Case "backupviewsettings": Config.BackupViewSettings = CBool(value)
        Case "restoreviewsettings": Config.RestoreViewSettings = CBool(value)
        Case "enablelogging": Config.enableLogging = CBool(value)
        Case "loglevel": Config.logLevel = value
        Case "logtofile": Config.logToFile = CBool(value)
        Case "maxlogsizemb": Config.maxLogSizeMb = CLng(value)
        Case "disablescreenupdating": Config.disableScreenUpdating = CBool(value)
        Case "disabledisplayalerts": Config.disableDisplayAlerts = CBool(value)
        Case "usebulkoperations": Config.useBulkOperations = CBool(value)
        Case "optimizefindreplace": Config.optimizeFindReplace = CBool(value)
        Case "showprogressmessages": Config.showProgressMessages = CBool(value)
        Case "showstatusbarupdates": Config.showStatusBarUpdates = CBool(value)
        Case "confirmcriticaloperations": Config.confirmCriticalOperations = CBool(value)
        Case "showcompletionmessage": Config.showCompletionMessage = CBool(value)
        Case "enableemergencyrecovery": Config.enableEmergencyRecovery = CBool(value)
        Case "timeoutoperations": Config.timeoutOperations = CBool(value)
        Case "supportword2010": Config.supportWord2010 = CBool(value)
        Case "supportword2013": Config.supportWord2013 = CBool(value)
        Case "supportword2016": Config.supportWord2016 = CBool(value)
        Case "usesafepropertyaccess": Config.useSafePropertyAccess = CBool(value)
        Case "fallbackmethods": Config.fallbackMethods = CBool(value)
        Case "handlemissingfeatures": Config.handleMissingFeatures = CBool(value)
        Case "requiredocumentsaved": Config.requireDocumentSaved = CBool(value)
        Case "validatefilepermissions": Config.validateFilePermissions = CBool(value)
        Case "checkdocumentprotection": Config.checkDocumentProtection = CBool(value)
        Case "enableemergencybackup": Config.enableEmergencyBackup = CBool(value)
        Case "sanitizeinputs": Config.sanitizeInputs = CBool(value)
        Case "validateranges": Config.validateRanges = CBool(value)
        Case "maxretryattempts": Config.maxRetryAttempts = CLng(value)
        Case "retrydelayms": Config.retryDelayMs = CLng(value)
    End Select
End Sub
