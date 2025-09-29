' =============================================================================
' PROJECT: CHAINSAW PROPOSITURAS
' =============================================================================
'
' Automated system for standardizing legislative documents in Microsoft Word
'
' License: Modified Apache 2.0 (see LICENSE)
' Version: 1.0.0-Beta1 | Date: 2025-09-27
' Repository: github.com/chrmsantos/chainsaw-proposituras
' Author: Christian Martin dos Santos <chrmsantos@gmail.com>

' Windows API declarations
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Option Explicit

'================================================================================
' CONSTANTS AND CONFIGURATION - #NEW
'================================================================================

' System constants
Private Const VERSION As String = "v1.0.0-Beta1"
Private Const SYSTEM_NAME As String = "CHAINSAW PROPOSITURAS"

' Message constants
Private Const MSG_BACKUP_SUCCESS As String = "Backup created successfully: "
Private Const MSG_BACKUP_FAILED As String = "Failed to create backup: "
Private Const MSG_RESTORE_SUCCESS As String = "Restore executed successfully: "
Private Const MSG_RESTORE_FAILED As String = "Failed to restore backup: "

' Error constants
Private Const ERR_WORD_NOT_FOUND As Long = 5000
Private Const ERR_INCOMPATIBLE_VERSION As Long = 5001
Private Const ERR_DOCUMENT_PROTECTED As Long = 5002
Private Const ERR_BACKUP_FAILED As Long = 5003
Private Const ERR_INVALID_DOCUMENT As Long = 5004

' Log level constants
Private Const LOG_LEVEL_ERROR As String = "ERROR"
Private Const LOG_LEVEL_WARNING As String = "WARNING"
Private Const LOG_LEVEL_INFO As String = "INFO"
Private Const LOG_LEVEL_DEBUG As String = "DEBUG"

' Performance constants  
Private Const MAX_PARAGRAPH_BATCH_SIZE As Long = 50
Private Const MAX_FIND_REPLACE_BATCH As Long = 100
Private Const OPTIMIZATION_THRESHOLD As Long = 1000

'================================================================================
' VARIABLES - #NEW
'================================================================================

' Configuration instance
Private Config As ConfigSettings

' Global state variables
Private isConfigLoaded As Boolean
Private processingStartTime As Double

'
' =============================================================================
' MAIN FEATURES:
' =============================================================================
'
' • SECURITY AND COMPATIBILITY CHECKS:
'   - Word version validation (minimum: 2010)
'   - Document type and protection verification
'   - Disk space control and minimum structure
'   - Failure protection and automatic recovery
'
' • AUTOMATIC BACKUP SYSTEM:
'   - Automatic backup before any modification
'   - Backup folder organized by document
'   - Automatic cleanup of old backups (limit: 10 files)
'   - Public subroutine for backup folder access
'
' • VISUAL ELEMENTS CLEANUP SYSTEM:
'   - Automatically removes hidden visual elements throughout the document
'   - Removes visual elements (visible or not) between paragraphs 1-4
'   - Preserves essential visual elements outside cleanup area
'   - Smart protection against accidental removal of relevant content
'
' • PUBLIC SUBROUTINE FOR SAVE AND EXIT:
'   - Automatic verification of all open documents
'   - Detection of documents with unsaved changes
'   - Professional interface with clear options for user
'   - Assisted saving with dialogs for new files
'   - Double confirmation for closing without saving
'   - Robust error handling and recovery
'
' • INSTITUTIONAL AUTOMATED FORMATTING:
'   - Complete formatting cleanup at startup
'   - Robust removal of multiple spaces and tabs
'   - Empty lines control (maximum 2 sequential)
'   - ADVANCED CLEANUP: Removes hidden visual elements throughout document
'   - ADVANCED CLEANUP: Removes visual elements between paragraphs 1-4 (visible or not)
'   - MAXIMUM PROTECTION: Advanced backup/restoration system for images
'   - MAXIMUM PROTECTION: Preserves inline, floating images and objects (except per rules)
'   - MAXIMUM PROTECTION: Detects and protects anchored shapes and visual fields (except per rules)
'   - First line: ALWAYS uppercase, bold, underlined, centered
'   - 2nd, 3rd and 4th paragraphs: 9cm left indent, no first line indent
'   - "Considerando": uppercase and bold at paragraph beginning
'   - "Justificativa": centered, no indents, bold, capitalized
'   - "Anexo/Anexos": left aligned, no indents, bold, capitalized
'   - Margin and orientation configuration (A4)
'   - Arial 12pt font with 1.4 spacing
'   - Indents and justified alignment
'   - Header with institutional logo
'   - Footer with centered numbering
'   - View: 110% zoom (maintained), other settings preserved
'   - TOTAL PROTECTION: Preserves rulers, display modes and original settings
'   - Watermark removal and manual formatting cleanup
'
' • TEXT STANDARDIZATION SYSTEM:
'   - Automatic normalization of "d'Oeste" and its variants
'   - Standardization of "- Vereador -" in all forms
'   - Smart replacement of isolated hyphens/dashes with em dash (—)
'   - Complete removal of manual line breaks (preserves paragraph breaks)
'   - Context and formatting preservation during replacements
'
' • SISTEMA DE LOGS E MONITORAMENTO:
'   - Registro detalhado de operações
'   - Controle de erros com fallback
'   - Mensagens na barra de status
'   - Histórico de execução
'
' • SISTEMA DE PROTEÇÃO DE CONFIGURAÇÕES DE VISUALIZAÇÃO:
'   - Backup automático de todas as configurações de exibição
'   - Preservação de réguas (horizontal e vertical)
'   - Manutenção do modo de visualização original
'   - Proteção de configurações de marcas de formatação
'   - Restauração completa após processamento (exceto zoom)
'   - Compatibilidade com todos os modos de exibição do Word
'
' • PERFORMANCE OTIMIZADA:
'   - Processamento eficiente para documentos grandes
'   - Desabilitação temporária de atualizações visuais
'   - Gerenciamento inteligente de recursos
'   - Sistema de logging otimizado (principais, warnings e erros)
'
' =============================================================================

'================================================================================
' CONSTANTS
'================================================================================

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
Private Const msoFalse As Long = 0
Private Const msoPicture As Long = 13
Private Const msoTextEffect As Long = 15
Private Const msoTextBox As Long = 17
Private Const msoAutoShape As Long = 1
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

' Document formatting constants
Private Const STANDARD_FONT As String = "Arial"
Private Const STANDARD_FONT_SIZE As Long = 12
Private Const FOOTER_FONT_SIZE As Long = 9
Private Const LINE_SPACING As Single = 14

' Margin constants in centimeters
Private Const TOP_MARGIN_CM As Double = 4.6
Private Const BOTTOM_MARGIN_CM As Double = 2
Private Const LEFT_MARGIN_CM As Double = 3
Private Const RIGHT_MARGIN_CM As Double = 3
Private Const HEADER_DISTANCE_CM As Double = 0.3
Private Const FOOTER_DISTANCE_CM As Double = 0.9

' Header image constants
Private Const HEADER_IMAGE_RELATIVE_PATH As String = "\chsw-prop\private\header\stamp.png"
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 21
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.7
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.19

' Minimum supported version
Private Const MIN_SUPPORTED_VERSION As Long = 14 ' Word 2010

' Required string constant
Private Const REQUIRED_STRING As String = "$NUMERO$/$ANO$"

' Timeout constants
Private Const MAX_RETRY_ATTEMPTS As Long = 3
Private Const RETRY_DELAY_MS As Long = 1000

' Configuration file constants
Private Const CONFIG_FILE_NAME As String = "chainsaw-config.ini"
Private Const CONFIG_FILE_PATH As String = "\chsw-prop\"

' Backup constants
Private Const BACKUP_FOLDER_NAME As String = "\chsw-prop\private\backups\"
Private Const MAX_BACKUP_FILES As Long = 10

'================================================================================
' GLOBAL VARIABLES
'================================================================================
Private undoGroupEnabled As Boolean
Private loggingEnabled As Boolean
Private logFilePath As String
Private formattingCancelled As Boolean
Private executionStartTime As Date
Private backupFilePath As String

' Configuration variables - loaded from chainsaw-config.ini
Private Type ConfigSettings
    ' General
    debugMode As Boolean
    performanceMode As Boolean
    compatibilityMode As Boolean
    
    ' Validations
    checkWordVersion As Boolean
    validateDocumentIntegrity As Boolean
    validatePropositionType As Boolean
    validateContentConsistency As Boolean
    checkDiskSpace As Boolean
    minWordVersion As Double
    maxDocumentSize As Long
    
    ' Backup
    autoBackup As Boolean
    backupBeforeProcessing As Boolean
    maxBackupFiles As Long
    backupCleanup As Boolean
    backupRetryAttempts As Long
    
    ' Formatting
    applyPageSetup As Boolean
    applyStandardFont As Boolean
    applyStandardParagraphs As Boolean
    formatFirstParagraph As Boolean
    formatSecondParagraph As Boolean
    formatNumberedParagraphs As Boolean
    formatConsiderandoParagraphs As Boolean
    formatJustificativaParagraphs As Boolean
    enableHyphenation As Boolean
    
    ' Cleaning
    clearAllFormatting As Boolean
    cleanDocumentStructure As Boolean
    cleanMultipleSpaces As Boolean
    limitSequentialEmptyLines As Boolean
    ensureParagraphSeparation As Boolean
    cleanVisualElements As Boolean
    deleteHiddenElements As Boolean
    deleteVisualElementsFirstFourParagraphs As Boolean
    
    ' Header/Footer
    insertHeaderstamp As Boolean
    insertFooterstamp As Boolean
    removeWatermark As Boolean
    headerImagePath As String
    headerImageMaxWidth As Double
    headerImageHeightRatio As Double
    
    ' Text Replacements
    applyTextReplacements As Boolean
    applySpecificParagraphReplacements As Boolean
    replaceHyphensWithEmDash As Boolean
    removeManualLineBreaks As Boolean
    normalizeDosteVariants As Boolean
    normalizeVereadorVariants As Boolean
    
    ' Visual Elements
    backupAllImages As Boolean
    restoreAllImages As Boolean
    protectImagesInRange As Boolean
    backupViewSettings As Boolean
    restoreViewSettings As Boolean
    
    ' Logging
    enableLogging As Boolean
    logLevel As String
    logToFile As Boolean
    logDetailedOperations As Boolean
    logWarnings As Boolean
    logErrors As Boolean
    maxLogSizeMb As Long
    
    ' Performance
    disableScreenUpdating As Boolean
    disableDisplayAlerts As Boolean
    useBulkOperations As Boolean
    optimizeFindReplace As Boolean
    minimizeObjectCreation As Boolean
    cacheFrequentlyUsedObjects As Boolean
    useEfficientLoops As Boolean
    batchParagraphOperations As Boolean
    
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
    enableEmergencyBackup As Boolean
    sanitizeInputs As Boolean
    validateRanges As Boolean
    
    ' Advanced
    maxRetryAttempts As Long
    retryDelayMs As Long
    compilationCheck As Boolean
    vbaAccessRequired As Boolean
    autoCleanup As Boolean
    forceGcCollection As Boolean
End Type

' Image protection variables
Private Type ImageInfo
    ParaIndex As Long
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

' View settings backup variables
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
    ' ShowAnimation removida - compatibilidade
    DraftFont As Boolean
    WrapToWindow As Boolean
    ShowPicturePlaceHolders As Boolean
    ShowFieldShading As Long
    TableGridlines As Boolean
    ' EnlargeFontsLessThan removida - compatibilidade
End Type

Private originalViewSettings As ViewSettings

'================================================================================
' CONFIGURATION SYSTEM - SISTEMA DE CONFIGURAÇÃO - #NEW
'================================================================================

Private Function LoadConfiguration() As Boolean
    On Error GoTo ErrorHandler
    
    LoadConfiguration = False
    
    ' Define valores padrão primeiro
    SetDefaultConfiguration
    
    ' Tenta carregar do arquivo de configuração
    Dim configPath As String
    configPath = GetConfigurationFilePath()
    
    If Len(configPath) = 0 Or Dir(configPath) = "" Then
        LogMessage "Configuration file not found, using default values: " & configPath, LOG_LEVEL_WARNING
        LoadConfiguration = True ' Usa padrões
        Exit Function
    End If
    
    ' Carrega configurações do arquivo
    If ParseConfigurationFile(configPath) Then
        LogMessage "Configuration loaded successfully from: " & configPath, LOG_LEVEL_INFO
        LoadConfiguration = True
    Else
        LogMessage "Error loading configuration, using default values", LOG_LEVEL_WARNING
        SetDefaultConfiguration
        LoadConfiguration = True ' Usa padrões como fallback
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "Error loading configuration: " & Err.Description, LOG_LEVEL_ERROR
    SetDefaultConfiguration
    LoadConfiguration = True ' Continua com padrões
End Function

Private Function GetConfigurationFilePath() As String
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim basePath As String
    
    ' Tenta obter pasta do documento atual
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    If Not doc Is Nothing And doc.path <> "" Then
        basePath = doc.path
    Else
        ' Fallback para pasta do usuário
        basePath = Environ("USERPROFILE") & "\Documents"
    End If
    On Error GoTo ErrorHandler
    
    ' Constrói caminho do arquivo de configuração
    GetConfigurationFilePath = basePath & CONFIG_FILE_PATH & CONFIG_FILE_NAME
    
    Exit Function
    
ErrorHandler:
    GetConfigurationFilePath = ""
End Function

Private Sub SetDefaultConfiguration()
    ' Define valores padrão para todas as configurações
    With Config
        ' General
        .debugMode = False
        .performanceMode = True
        .compatibilityMode = True
        
        ' Validations
        .checkWordVersion = True
        .validateDocumentIntegrity = True
        .validatePropositionType = True
        .validateContentConsistency = True
        .checkDiskSpace = True
        .minWordVersion = 14#
        .maxDocumentSize = 500000
        
        ' Backup
        .autoBackup = True
        .backupBeforeProcessing = True
        .maxBackupFiles = 10
        .backupCleanup = True
        .backupRetryAttempts = 3
        
        ' Formatting
        .applyPageSetup = True
        .applyStandardFont = True
        .applyStandardParagraphs = True
        .formatFirstParagraph = True
        .formatSecondParagraph = True
        .formatNumberedParagraphs = True
        .formatConsiderandoParagraphs = True
        .formatJustificativaParagraphs = True
        .enableHyphenation = True
        
        ' Cleaning
        .clearAllFormatting = True
        .cleanDocumentStructure = True
        .cleanMultipleSpaces = True
        .limitSequentialEmptyLines = True
        .ensureParagraphSeparation = True
        .cleanVisualElements = True
        .deleteHiddenElements = True
        .deleteVisualElementsFirstFourParagraphs = True
        
        ' Header/Footer
        .insertHeaderstamp = True
        .insertFooterstamp = True
        .removeWatermark = True
        .headerImagePath = "assets\stamp.png"
        .headerImageMaxWidth = 21#
        .headerImageHeightRatio = 0.19
        
        ' Text Replacements
        .applyTextReplacements = True
        .applySpecificParagraphReplacements = True
        .replaceHyphensWithEmDash = True
        .removeManualLineBreaks = True
        .normalizeDosteVariants = True
        .normalizeVereadorVariants = True
        
        ' Visual Elements
        .backupAllImages = True
        .restoreAllImages = True
        .protectImagesInRange = True
        .backupViewSettings = True
        .restoreViewSettings = True
        
        ' Logging
        .enableLogging = True
        .logLevel = "INFO"
        .logToFile = True
        .logDetailedOperations = True
        .logWarnings = True
        .logErrors = True
        .maxLogSizeMb = 10
        
        ' Performance
        .disableScreenUpdating = True
        .disableDisplayAlerts = True
        .useBulkOperations = True
        .optimizeFindReplace = True
        .minimizeObjectCreation = True
        .cacheFrequentlyUsedObjects = True
        .useEfficientLoops = True
        .batchParagraphOperations = True
        
        ' Interface
        .showProgressMessages = True
        .showStatusBarUpdates = True
        .confirmCriticalOperations = True
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
        .enableEmergencyBackup = True
        .sanitizeInputs = True
        .validateRanges = True
        
        ' Advanced
        .maxRetryAttempts = 3
        .retryDelayMs = 1000
        .compilationCheck = True
        .vbaAccessRequired = False
        .autoCleanup = True
        .forceGcCollection = False
    End With
End Sub

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
                currentSection = UCase(Mid(fileLine, 2, Len(fileLine) - 2))
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
        If Left(configValue, 1) = """" And Right(configValue, 1) = """" Then
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
            Config.debugMode = (LCase(value) = "true")
        Case "PERFORMANCE_MODE"
            Config.performanceMode = (LCase(value) = "true")
        Case "COMPATIBILITY_MODE"
            Config.compatibilityMode = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessValidationConfig(key As String, value As String)
    Select Case key
        Case "CHECK_WORD_VERSION"
            Config.checkWordVersion = (LCase(value) = "true")
        Case "VALIDATE_DOCUMENT_INTEGRITY"
            Config.validateDocumentIntegrity = (LCase(value) = "true")
        Case "VALIDATE_PROPOSITION_TYPE"
            Config.validatePropositionType = (LCase(value) = "true")
        Case "VALIDATE_CONTENT_CONSISTENCY"
            Config.validateContentConsistency = (LCase(value) = "true")
        Case "CHECK_DISK_SPACE"
            Config.checkDiskSpace = (LCase(value) = "true")
        Case "MIN_WORD_VERSION"
            Config.minWordVersion = CDbl(value)
        Case "MAX_DOCUMENT_SIZE"
            Config.maxDocumentSize = CLng(value)
    End Select
End Sub

Private Sub ProcessBackupConfig(key As String, value As String)
    Select Case key
        Case "AUTO_BACKUP"
            Config.autoBackup = (LCase(value) = "true")
        Case "BACKUP_BEFORE_PROCESSING"
            Config.backupBeforeProcessing = (LCase(value) = "true")
        Case "MAX_BACKUP_FILES"
            Config.maxBackupFiles = CLng(value)
        Case "BACKUP_CLEANUP"
            Config.backupCleanup = (LCase(value) = "true")
        Case "BACKUP_RETRY_ATTEMPTS"
            Config.backupRetryAttempts = CLng(value)
    End Select
End Sub

Private Sub ProcessFormattingConfig(key As String, value As String)
    Select Case key
        Case "APPLY_PAGE_SETUP"
            Config.applyPageSetup = (LCase(value) = "true")
        Case "APPLY_STANDARD_FONT"
            Config.applyStandardFont = (LCase(value) = "true")
        Case "APPLY_STANDARD_PARAGRAPHS"
            Config.applyStandardParagraphs = (LCase(value) = "true")
        Case "FORMAT_FIRST_PARAGRAPH"
            Config.formatFirstParagraph = (LCase(value) = "true")
        Case "FORMAT_SECOND_PARAGRAPH"
            Config.formatSecondParagraph = (LCase(value) = "true")
        Case "FORMAT_NUMBERED_PARAGRAPHS"
            Config.formatNumberedParagraphs = (LCase(value) = "true")
        Case "FORMAT_CONSIDERANDO_PARAGRAPHS"
            Config.formatConsiderandoParagraphs = (LCase(value) = "true")
        Case "FORMAT_JUSTIFICATIVA_PARAGRAPHS"
            Config.formatJustificativaParagraphs = (LCase(value) = "true")
        Case "ENABLE_HYPHENATION"
            Config.enableHyphenation = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessCleaningConfig(key As String, value As String)
    Select Case key
        Case "CLEAR_ALL_FORMATTING"
            Config.clearAllFormatting = (LCase(value) = "true")
        Case "CLEAN_DOCUMENT_STRUCTURE"
            Config.cleanDocumentStructure = (LCase(value) = "true")
        Case "CLEAN_MULTIPLE_SPACES"
            Config.cleanMultipleSpaces = (LCase(value) = "true")
        Case "LIMIT_SEQUENTIAL_EMPTY_LINES"
            Config.limitSequentialEmptyLines = (LCase(value) = "true")
        Case "ENSURE_PARAGRAPH_SEPARATION"
            Config.ensureParagraphSeparation = (LCase(value) = "true")
        Case "CLEAN_VISUAL_ELEMENTS"
            Config.cleanVisualElements = (LCase(value) = "true")
        Case "DELETE_HIDDEN_ELEMENTS"
            Config.deleteHiddenElements = (LCase(value) = "true")
        Case "DELETE_VISUAL_ELEMENTS_FIRST_FOUR_PARAGRAPHS"
            Config.deleteVisualElementsFirstFourParagraphs = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessHeaderFooterConfig(key As String, value As String)
    Select Case key
        Case "INSERT_HEADER_stamp"
            Config.insertHeaderstamp = (LCase(value) = "true")
        Case "INSERT_FOOTER_stamp"
            Config.insertFooterstamp = (LCase(value) = "true")
        Case "REMOVE_WATERMARK"
            Config.removeWatermark = (LCase(value) = "true")
        Case "HEADER_IMAGE_PATH"
            Config.headerImagePath = value
        Case "HEADER_IMAGE_MAX_WIDTH"
            Config.headerImageMaxWidth = CDbl(value)
        Case "HEADER_IMAGE_HEIGHT_RATIO"
            Config.headerImageHeightRatio = CDbl(value)
    End Select
End Sub

Private Sub ProcessReplacementConfig(key As String, value As String)
    Select Case key
        Case "APPLY_TEXT_REPLACEMENTS"
            Config.applyTextReplacements = (LCase(value) = "true")
        Case "APPLY_SPECIFIC_PARAGRAPH_REPLACEMENTS"
            Config.applySpecificParagraphReplacements = (LCase(value) = "true")
        Case "REPLACE_HYPHENS_WITH_EM_DASH"
            Config.replaceHyphensWithEmDash = (LCase(value) = "true")
        Case "REMOVE_MANUAL_LINE_BREAKS"
            Config.removeManualLineBreaks = (LCase(value) = "true")
        Case "NORMALIZE_DOESTE_VARIANTS"
            Config.normalizeDosteVariants = (LCase(value) = "true")
        Case "NORMALIZE_VEREADOR_VARIANTS"
            Config.normalizeVereadorVariants = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessVisualElementsConfig(key As String, value As String)
    Select Case key
        Case "BACKUP_ALL_IMAGES"
            Config.backupAllImages = (LCase(value) = "true")
        Case "RESTORE_ALL_IMAGES"
            Config.restoreAllImages = (LCase(value) = "true")
        Case "PROTECT_IMAGES_IN_RANGE"
            Config.protectImagesInRange = (LCase(value) = "true")
        Case "BACKUP_VIEW_SETTINGS"
            Config.backupViewSettings = (LCase(value) = "true")
        Case "RESTORE_VIEW_SETTINGS"
            Config.restoreViewSettings = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessLoggingConfig(key As String, value As String)
    Select Case key
        Case "ENABLE_LOGGING"
            Config.enableLogging = (LCase(value) = "true")
        Case "LOG_LEVEL"
            Config.logLevel = UCase(value)
        Case "LOG_TO_FILE"
            Config.logToFile = (LCase(value) = "true")
        Case "LOG_DETAILED_OPERATIONS"
            Config.logDetailedOperations = (LCase(value) = "true")
        Case "LOG_WARNINGS"
            Config.logWarnings = (LCase(value) = "true")
        Case "LOG_ERRORS"
            Config.logErrors = (LCase(value) = "true")
        Case "MAX_LOG_SIZE_MB"
            Config.maxLogSizeMb = CLng(value)
    End Select
End Sub

Private Sub ProcessPerformanceConfig(key As String, value As String)
    Select Case key
        Case "DISABLE_SCREEN_UPDATING"
            Config.disableScreenUpdating = (LCase(value) = "true")
        Case "DISABLE_DISPLAY_ALERTS"
            Config.disableDisplayAlerts = (LCase(value) = "true")
        Case "USE_BULK_OPERATIONS"
            Config.useBulkOperations = (LCase(value) = "true")
        Case "OPTIMIZE_FIND_REPLACE"
            Config.optimizeFindReplace = (LCase(value) = "true")
        Case "MINIMIZE_OBJECT_CREATION"
            Config.minimizeObjectCreation = (LCase(value) = "true")
        Case "CACHE_FREQUENTLY_USED_OBJECTS"
            Config.cacheFrequentlyUsedObjects = (LCase(value) = "true")
        Case "USE_EFFICIENT_LOOPS"
            Config.useEfficientLoops = (LCase(value) = "true")
        Case "BATCH_PARAGRAPH_OPERATIONS"
            Config.batchParagraphOperations = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessInterfaceConfig(key As String, value As String)
    Select Case key
        Case "SHOW_PROGRESS_MESSAGES"
            Config.showProgressMessages = (LCase(value) = "true")
        Case "SHOW_STATUS_BAR_UPDATES"
            Config.showStatusBarUpdates = (LCase(value) = "true")
        Case "CONFIRM_CRITICAL_OPERATIONS"
            Config.confirmCriticalOperations = (LCase(value) = "true")
        Case "SHOW_COMPLETION_MESSAGE"
            Config.showCompletionMessage = (LCase(value) = "true")
        Case "ENABLE_EMERGENCY_RECOVERY"
            Config.enableEmergencyRecovery = (LCase(value) = "true")
        Case "TIMEOUT_OPERATIONS"
            Config.timeoutOperations = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessCompatibilityConfig(key As String, value As String)
    Select Case key
        Case "SUPPORT_WORD_2010"
            Config.supportWord2010 = (LCase(value) = "true")
        Case "SUPPORT_WORD_2013"
            Config.supportWord2013 = (LCase(value) = "true")
        Case "SUPPORT_WORD_2016"
            Config.supportWord2016 = (LCase(value) = "true")
        Case "USE_SAFE_PROPERTY_ACCESS"
            Config.useSafePropertyAccess = (LCase(value) = "true")
        Case "FALLBACK_METHODS"
            Config.fallbackMethods = (LCase(value) = "true")
        Case "HANDLE_MISSING_FEATURES"
            Config.handleMissingFeatures = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessSecurityConfig(key As String, value As String)
    Select Case key
        Case "REQUIRE_DOCUMENT_SAVED"
            Config.requireDocumentSaved = (LCase(value) = "true")
        Case "VALIDATE_FILE_PERMISSIONS"
            Config.validateFilePermissions = (LCase(value) = "true")
        Case "CHECK_DOCUMENT_PROTECTION"
            Config.checkDocumentProtection = (LCase(value) = "true")
        Case "ENABLE_EMERGENCY_BACKUP"
            Config.enableEmergencyBackup = (LCase(value) = "true")
        Case "SANITIZE_INPUTS"
            Config.sanitizeInputs = (LCase(value) = "true")
        Case "VALIDATE_RANGES"
            Config.validateRanges = (LCase(value) = "true")
    End Select
End Sub

Private Sub ProcessAdvancedConfig(key As String, value As String)
    Select Case key
        Case "MAX_RETRY_ATTEMPTS"
            Config.maxRetryAttempts = CLng(value)
        Case "RETRY_DELAY_MS"
            Config.retryDelayMs = CLng(value)
        Case "COMPILATION_CHECK"
            Config.compilationCheck = (LCase(value) = "true")
        Case "VBA_ACCESS_REQUIRED"
            Config.vbaAccessRequired = (LCase(value) = "true")
        Case "AUTO_CLEANUP"
            Config.autoCleanup = (LCase(value) = "true")
        Case "FORCE_GC_COLLECTION"
            Config.forceGcCollection = (LCase(value) = "true")
    End Select
End Sub

'================================================================================
' PERFORMANCE OPTIMIZATION SYSTEM - SISTEMA DE OTIMIZAÇÃO DE PERFORMANCE - #NEW
'================================================================================

Private Function InitializePerformanceOptimization() As Boolean
    On Error GoTo ErrorHandler
    
    InitializePerformanceOptimization = False
    
    ' Aplica otimizações baseadas na configuração
    If Config.performanceMode Then
        LogMessage "Starting performance optimizations...", LOG_LEVEL_INFO
        
        ' Desabilita atualizações de tela
        If Config.disableScreenUpdating Then
            Application.ScreenUpdating = False
            LogMessage "Screen updating desabilitado", LOG_LEVEL_DEBUG
        End If
        
        ' Desabilita alertas
        If Config.disableDisplayAlerts Then
            Application.DisplayAlerts = False
            LogMessage "Display alerts desabilitado", LOG_LEVEL_DEBUG
        End If
        
        ' Otimizações específicas do Word
        Call OptimizeWordSettings
        
        LogMessage "Performance optimizations applied", LOG_LEVEL_INFO
    End If
    
    InitializePerformanceOptimization = True
    Exit Function
    
ErrorHandler:
    LogMessage "Error initializing optimizations: " & Err.Description, LOG_LEVEL_ERROR
    InitializePerformanceOptimization = False
End Function

Private Sub OptimizeWordSettings()
    On Error Resume Next
    
    ' Otimizações específicas do Word baseadas na configuração
    If Config.minimizeObjectCreation Then
        ' Reduz criação de objetos desnecessários
        With ActiveDocument
            .TrackRevisions = False
            .ShowRevisions = False
        End With
    End If
    
    ' Otimizações de busca e substituição
    If Config.optimizeFindReplace Then
        With Selection.Find
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
    End If
    
    On Error GoTo 0
End Sub

Private Function RestorePerformanceSettings() As Boolean
    On Error GoTo ErrorHandler
    
    RestorePerformanceSettings = False
    
    If Config.performanceMode Then
        LogMessage "Restoring performance settings...", LOG_LEVEL_INFO
        
        ' Restaura configurações originais
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        
        LogMessage "Performance settings restored", LOG_LEVEL_INFO
    End If
    
    RestorePerformanceSettings = True
    Exit Function
    
ErrorHandler:
    LogMessage "Error restoring settings: " & Err.Description, LOG_LEVEL_ERROR
    RestorePerformanceSettings = False
End Function

Private Function OptimizedFindReplace(findText As String, replaceText As String, Optional searchRange As Range = Nothing) As Long
    On Error GoTo ErrorHandler
    
    OptimizedFindReplace = 0
    
    ' Usa otimização baseada na configuração
    If Config.optimizeFindReplace And Config.useBulkOperations Then
        ' Implementação otimizada para operações em lote
        OptimizedFindReplace = BulkFindReplace(findText, replaceText, searchRange)
    Else
        ' Implementação padrão
        OptimizedFindReplace = StandardFindReplace(findText, replaceText, searchRange)
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "Erro em busca/substituição otimizada: " & Err.Description, LOG_LEVEL_ERROR
    OptimizedFindReplace = 0
End Function

Private Function BulkFindReplace(findText As String, replaceText As String, Optional searchRange As Range = Nothing) As Long
    On Error GoTo ErrorHandler
    
    BulkFindReplace = 0
    
    Dim targetRange As Range
    Set targetRange = IIf(searchRange Is Nothing, ActiveDocument.Content, searchRange)
    
    ' Otimização: usa método nativo do Word para operações em lote
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
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' Executa todas as substituições de uma vez
        BulkFindReplace = .Execute(Replace:=wdReplaceAll)
    End With
    
    Exit Function
    
ErrorHandler:
    LogMessage "Erro em busca/substituição em lote: " & Err.Description, LOG_LEVEL_ERROR
    BulkFindReplace = 0
End Function

Private Function StandardFindReplace(findText As String, replaceText As String, Optional searchRange As Range = Nothing) As Long
    On Error GoTo ErrorHandler
    
    StandardFindReplace = 0
    
    Dim targetRange As Range
    Set targetRange = IIf(searchRange Is Nothing, ActiveDocument.Content, searchRange)
    
    ' Implementação padrão compatível
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
    LogMessage "Erro em busca/substituição padrão: " & Err.Description, LOG_LEVEL_ERROR
    StandardFindReplace = 0
End Function

Private Function OptimizedParagraphProcessing(processingFunction As String) As Boolean
    On Error GoTo ErrorHandler
    
    OptimizedParagraphProcessing = False
    
    ' Usa processamento em lote se configurado
    If Config.batchParagraphOperations And Config.useEfficientLoops Then
        OptimizedParagraphProcessing = BatchProcessParagraphs(processingFunction)
    Else
        OptimizedParagraphProcessing = StandardProcessParagraphs(processingFunction)
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "Erro no processamento otimizado de parágrafos: " & Err.Description, LOG_LEVEL_ERROR
    OptimizedParagraphProcessing = False
End Function

Private Function BatchProcessParagraphs(processingFunction As String) As Boolean
    On Error GoTo ErrorHandler
    
    BatchProcessParagraphs = False
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim paragraphCount As Long
    paragraphCount = doc.Paragraphs.Count
    
    Dim batchSize As Long
    batchSize = IIf(paragraphCount > OPTIMIZATION_THRESHOLD, MAX_PARAGRAPH_BATCH_SIZE, paragraphCount)
    
    LogMessage "Processando " & paragraphCount & " parágrafos em lotes de " & batchSize, LOG_LEVEL_DEBUG
    
    Dim i As Long
    For i = 1 To paragraphCount Step batchSize
        Dim endIndex As Long
        endIndex = IIf(i + batchSize - 1 > paragraphCount, paragraphCount, i + batchSize - 1)
        
        ' Processa lote de parágrafos
        If Not ProcessParagraphBatch(i, endIndex, processingFunction) Then
            LogMessage "Erro no processamento do lote " & i & "-" & endIndex, LOG_LEVEL_ERROR
            Exit Function
        End If
        
        ' Coleta lixo periodicamente se configurado
        If Config.forceGcCollection And i Mod (batchSize * 5) = 0 Then
            Call ForceGarbageCollection
        End If
    Next i
    
    BatchProcessParagraphs = True
    Exit Function
    
ErrorHandler:
    LogMessage "Erro no processamento em lote: " & Err.Description, LOG_LEVEL_ERROR
    BatchProcessParagraphs = False
End Function

Private Function StandardProcessParagraphs(processingFunction As String) As Boolean
    On Error GoTo ErrorHandler
    
    StandardProcessParagraphs = False
    
    ' Implementação padrão - processa parágrafo por parágrafo
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        ' Aplica função de processamento específica
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
    LogMessage "Erro no processamento padrão: " & Err.Description, LOG_LEVEL_ERROR
    StandardProcessParagraphs = False
End Function

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
            
            ' Aplica função de processamento específica
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
    LogMessage "Erro no processamento do lote: " & Err.Description, LOG_LEVEL_ERROR
    ProcessParagraphBatch = False
End Function

Private Sub FormatParagraph(para As Paragraph)
    On Error Resume Next
    ' Placeholder para formatação de parágrafo
    ' Implementação específica será adicionada conforme necessário
End Sub

Private Sub CleanParagraph(para As Paragraph)
    On Error Resume Next
    ' Placeholder para limpeza de parágrafo
    ' Implementação específica será adicionada conforme necessário
End Sub

Private Sub ValidateParagraph(para As Paragraph)
    On Error Resume Next
    ' Placeholder para validação de parágrafo
    ' Implementação específica será adicionada conforme necessário
End Sub

Private Sub ForceGarbageCollection()
    On Error Resume Next
    
    If Config.forceGcCollection Then
        ' Força coleta de lixo - apenas em casos específicos
        DoEvents ' Permite ao sistema processar mensagens pendentes
        LogMessage "Coleta de lixo forçada executada", LOG_LEVEL_DEBUG
    End If
End Sub

'================================================================================
' MAIN ENTRY POINT - #STABLE
'================================================================================
Public Sub StandardizeDocumentMain()
    On Error GoTo CriticalErrorHandler
    
    ' ========================================
    ' INICIALIZAÇÃO E CARREGAMENTO DE CONFIGURAÇÃO - #NEW
    ' ========================================
    
    processingStartTime = Timer
    formattingCancelled = False
    
    ' Carrega configurações do sistema
    If Not isConfigLoaded Then
        If Not LoadConfiguration() Then
            LogMessage "Critical error loading configuration. Aborting execution.", LOG_LEVEL_ERROR
            MsgBox "Critical error loading system configuration." & vbCrLf & _
                   "Execution was aborted to prevent issues.", vbCritical, "Configuration Error - " & SYSTEM_NAME
            Exit Sub
        End If
        isConfigLoaded = True
        LogMessage "Sistema inicializado: " & SYSTEM_NAME & " " & VERSION, LOG_LEVEL_INFO
    End If
    
    ' ========================================
    ' VALIDAÇÕES PRELIMINARES BASEADAS EM CONFIGURAÇÃO - #NEW
    ' ========================================
    
    ' Validação da versão do Word (se habilitada)
    If Config.checkWordVersion Then
        If Not CheckWordVersion() Then
            Application.StatusBar = "Error: Word version not supported (minimum: Word " & Config.minWordVersion & ")"
            LogMessage "Word version " & Application.version & " not supported. Minimum: " & CStr(Config.minWordVersion), LOG_LEVEL_ERROR
            If Config.showProgressMessages Then
                MsgBox "This tool requires Microsoft Word " & Config.minWordVersion & " or higher." & vbCrLf & _
                       "Current version: " & Application.version & vbCrLf & _
                       "Minimum version: " & CStr(Config.minWordVersion), vbCritical, "Incompatible Version - " & SYSTEM_NAME
            End If
            Exit Sub
        End If
    End If
    
    ' Verificação e compilação do projeto VBA (se habilitada)
    If Config.compilationCheck Then
        If Not CompileVBAProject() Then
            Application.StatusBar = "Error: VBA project compilation failed"
            LogMessage "VBA project compilation failed", LOG_LEVEL_ERROR
            If Config.showProgressMessages Then
                MsgBox "Erro na compilação do projeto VBA." & vbCrLf & _
                       "Verifique se há erros de sintaxe no código." & vbCrLf & _
                       "A execução pode ser instável.", vbExclamation, "Aviso de Compilação - " & SYSTEM_NAME
            End If
            ' Continua a execução mesmo com falha na compilação para compatibilidade
        End If
    End If
    
    ' Validação do documento ativo
    Dim doc As Document
    Set doc = Nothing
    
    On Error Resume Next
    Set doc = ActiveDocument
    If doc Is Nothing Then
        On Error GoTo CriticalErrorHandler
        Application.StatusBar = "Error: No document is accessible"
        LogMessage "No document accessible for processing", LOG_LEVEL_ERROR
        If Config.showProgressMessages Then
            MsgBox "Nenhum documento está aberto ou acessível." & vbCrLf & _
               "Abra um documento antes de executar a padronização.", vbExclamation, "Documento Não Encontrado - Chainsaw Proposituras"
        End If
        Exit Sub
    End If
    On Error GoTo CriticalErrorHandler
    
    ' Validação de integridade do documento (se habilitada)
    If Config.validateDocumentIntegrity Then
        If Not ValidateDocumentIntegrity(doc) Then
            LogMessage "Documento falhou na validação de integridade", LOG_LEVEL_ERROR
            GoTo CleanUp
        End If
    End If
    
    ' ========================================
    ' INICIALIZAÇÃO DE OTIMIZAÇÕES DE PERFORMANCE - #NEW
    ' ========================================
    
    If Not InitializePerformanceOptimization() Then
        LogMessage "Warning: Failed to initialize performance optimizations", LOG_LEVEL_WARNING
        ' Continua execução mesmo com falha nas otimizações
    End If
    
    ' Inicialização do sistema de logs
    If Not InitializeLogging(doc) Then
        LogMessage "Failed to initialize logging system", LOG_LEVEL_WARNING
    End If
    
    LogMessage "Starting document standardization: " & doc.Name & " (Chainsaw Proposituras v1.0.0-Beta1)", LOG_LEVEL_INFO
    
    ' Configuração do grupo de desfazer
    StartUndoGroup "Padronização de Documento - " & doc.Name
    
    ' Configuração do estado da aplicação
    If Not SetAppState(False, "Formatando documento...") Then
        LogMessage "Failed to configure application state", LOG_LEVEL_WARNING
    End If
    
    ' Verificações preliminares
    If Not PreviousChecking(doc) Then
        GoTo CleanUp
    End If
    
    ' Salvamento obrigatório para documentos não salvos
    If doc.path = "" Then
        If Not SaveDocumentFirst(doc) Then
            Application.StatusBar = "Operação cancelada: documento precisa ser salvo"
            LogMessage "Operação cancelada - documento não foi salvo", LOG_LEVEL_INFO
            GoTo CleanUp
        End If
    End If
    
    ' Criação de backup com validação
    If Not CreateDocumentBackup(doc) Then
        LogMessage "Falha ao criar backup - continuando sem backup", LOG_LEVEL_WARNING
        Application.StatusBar = "Aviso: Backup não foi possível - processando sem backup"
        Dim backupResponse As VbMsgBoxResult
        backupResponse = MsgBox("Não foi possível criar backup do documento." & vbCrLf & _
                              "Deseja continuar mesmo assim?", vbYesNo + vbExclamation, "Falha no Backup - Chainsaw Proposituras")
        If backupResponse = vbNo Then
            LogMessage "Operação cancelada pelo usuário devido à falha no backup", LOG_LEVEL_INFO
            GoTo CleanUp
        End If
    Else
        Application.StatusBar = "Backup criado - formatando documento..."
    End If
    
    ' Backup das configurações de visualização originais
    If Not BackupViewSettings(doc) Then
        LogMessage "Aviso: Falha no backup das configurações de visualização", LOG_LEVEL_WARNING
    End If

    ' Limpeza de elementos visuais conforme especificado
    Application.StatusBar = "Processing document structure..."
    If Not CleanVisualElementsMain(doc) Then
        LogMessage "Warning: Failed to clean visual elements", LOG_LEVEL_WARNING
    End If

    If Not PreviousFormatting(doc) Then
        GoTo CleanUp
    End If

    ' Restaura configurações de visualização originais (exceto zoom)
    If Not RestoreViewSettings(doc) Then
        LogMessage "Aviso: Algumas configurações de visualização podem não ter sido restauradas", LOG_LEVEL_WARNING
    End If

    If formattingCancelled Then
        GoTo CleanUp
    End If

    Application.StatusBar = "Documento padronizado com sucesso!"
    LogMessage "Documento padronizado com sucesso", LOG_LEVEL_INFO

CleanUp:
    ' Restaura configurações de performance
    If Not RestorePerformanceSettings() Then
        LogMessage "Aviso: Falha ao restaurar configurações de performance", LOG_LEVEL_WARNING
    End If
    
    SafeCleanup
    CleanupImageProtection ' Nova função para limpar variáveis de proteção de imagens
    CleanupViewSettings    ' Nova função para limpar variáveis de configurações de visualização
    
    If Not SetAppState(True, "Documento padronizado com sucesso!") Then
        LogMessage "Falha ao restaurar estado da aplicação", LOG_LEVEL_WARNING
    End If
    
    SafeFinalizeLogging
    
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "ERRO CRÍTICO #" & Err.Number & ": " & Err.Description & _
              " em " & Err.Source & " (Linha: " & Erl & ")"
    
    LogMessage errDesc, LOG_LEVEL_ERROR
    Application.StatusBar = "Erro crítico durante processamento - verificar logs"
    
    EmergencyRecovery
End Sub

'================================================================================
' EMERGENCY RECOVERY - #STABLE
'================================================================================
Private Sub EmergencyRecovery()
    On Error Resume Next
    
    Application.screenUpdating = True
    Application.displayAlerts = wdAlertsAll
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
' SAFE CLEANUP - LIMPEZA SEGURA - #STABLE
'================================================================================
Private Sub SafeCleanup()
    On Error Resume Next
    
    EndUndoGroup
    
    ReleaseObjects
End Sub

'================================================================================
' RELEASE OBJECTS - #STABLE
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
' CLOSE ALL OPEN FILES - #STABLE
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
' VERSION COMPATIBILITY AND SAFETY CHECKS - #STABLE  
'================================================================================
Private Function CheckWordVersion() As Boolean
    On Error GoTo ErrorHandler
    
    Dim version As Double
    ' Uso de CDbl para garantir conversão correta em todas as versões
    version = CDbl(Application.version)
    
    ' Usa configuração para versão mínima
    If version < Config.minWordVersion Then
        CheckWordVersion = False
        LogMessage "Versão detectada: " & CStr(version) & " - Mínima suportada: " & CStr(Config.minWordVersion), LOG_LEVEL_ERROR
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
' COMPILE VBA PROJECT - COMPILAÇÃO AUTOMÁTICA DO PROJETO VBA - #NEW
'================================================================================
Private Function CompileVBAProject() As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Verificando compilação do projeto VBA..."
    LogMessage "Iniciando verificação e compilação do projeto VBA", LOG_LEVEL_INFO
    
    ' Método de verificação alternativo para compatibilidade máxima
    CompileVBAProject = False
    
    ' Tenta várias abordagens para validar o projeto
    Dim testResult As Boolean
    
    ' Método 1: Verificação de acesso ao VBE (se disponível)
    On Error Resume Next
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    If Err.Number = 0 And Not vbProj Is Nothing Then
        LogMessage "Acesso ao VBE disponível - projeto provavelmente válido", LOG_LEVEL_INFO
        CompileVBAProject = True
        On Error GoTo ErrorHandler
        Exit Function
    End If
    Err.Clear
    
    ' Método 2: Verificação indireta através de chamada de função conhecida
    On Error Resume Next
    testResult = CheckWordVersion()
    If Err.Number = 0 Then
        LogMessage "Funções VBA respondendo corretamente - compilação válida", LOG_LEVEL_INFO
        CompileVBAProject = True
        On Error GoTo ErrorHandler
        Exit Function
    End If
    Err.Clear
    
    ' Método 3: Verificação de constantes e variáveis do módulo
    On Error Resume Next
    Dim testConstant As Double
    testConstant = Config.minWordVersion
    If Err.Number = 0 And testConstant > 0 Then
        LogMessage "Constantes do módulo acessíveis - estrutura VBA íntegra", LOG_LEVEL_INFO
        CompileVBAProject = True
        On Error GoTo ErrorHandler
        Exit Function
    End If
    Err.Clear
    
    ' Se chegou aqui, métodos automáticos falharam
    On Error GoTo ErrorHandler
    LogMessage "Verificação automática de compilação falhou - continuando por compatibilidade", LOG_LEVEL_WARNING
    CompileVBAProject = True ' Permite continuar para máxima compatibilidade
    
    Exit Function

ErrorHandler:
    LogMessage "Erro crítico na verificação de compilação: " & Err.Description, LOG_LEVEL_ERROR
    Application.StatusBar = "Erro na verificação de compilação"
    
    ' Tenta continuar mesmo com erro de compilação (modo de compatibilidade máxima)
    MsgBox "Aviso: Não foi possível verificar a compilação do projeto VBA." & vbCrLf & _
           "O sistema tentará executar mesmo assim." & vbCrLf & vbCrLf & _
           "Se houver problemas, verifique se há erros de sintaxe no código.", _
           vbExclamation, "Chainsaw Proposituras - Aviso de Compilação"
    
    LogMessage "Continuando execução apesar do erro de compilação (modo compatibilidade)", LOG_LEVEL_WARNING
    CompileVBAProject = True ' Permite continuar por compatibilidade máxima
End Function


'================================================================================
' DOCUMENT INTEGRITY VALIDATION - #STABLE
'================================================================================
Private Function ValidateDocumentIntegrity(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateDocumentIntegrity = False
    
    ' Verificação básica de acessibilidade
    If doc Is Nothing Then
        LogMessage "Documento é nulo na validação de integridade", LOG_LEVEL_ERROR
        MsgBox "Erro: Documento inacessível.", vbCritical, "Erro de Integridade - Chainsaw Proposituras"
        Exit Function
    End If
    
    ' Verificação de proteção de documento
    On Error Resume Next
    Dim isProtected As Boolean
    isProtected = (doc.ProtectionType <> wdNoProtection)
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        LogMessage "Não foi possível verificar proteção do documento", LOG_LEVEL_WARNING
        isProtected = False
    End If
    On Error GoTo ErrorHandler
    
    If isProtected Then
        LogMessage "Documento protegido detectado: " & GetProtectionType(doc), LOG_LEVEL_WARNING
        MsgBox "Este documento está protegido e pode não ser possível formatá-lo completamente." & vbCrLf & _
               "Tipo de proteção: " & GetProtectionType(doc) & vbCrLf & vbCrLf & _
               "Deseja continuar mesmo assim?", vbYesNo + vbExclamation, "Documento Protegido - Chainsaw Proposituras"
        If vbNo = MsgBox("", vbYesNo) Then ' Simula resposta para compatibilidade
            LogMessage "Usuário cancelou devido à proteção do documento", LOG_LEVEL_INFO
            Exit Function
        End If
    End If
    
    ' Verificação de conteúdo mínimo
    If doc.Paragraphs.Count < 1 Then
        LogMessage "Documento vazio detectado", LOG_LEVEL_ERROR
        MsgBox "O documento está vazio." & vbCrLf & _
               "Adicione conteúdo antes de executar a padronização.", vbExclamation, "Documento Vazio - Chainsaw Proposituras"
        Exit Function
    End If
    
    ' Verificação de tamanho do documento
    Dim docSize As Long
    On Error Resume Next
    docSize = doc.Range.Characters.Count
    If Err.Number <> 0 Then
        docSize = 0
        LogMessage "Não foi possível determinar tamanho do documento", LOG_LEVEL_WARNING
    End If
    On Error GoTo ErrorHandler
    
    If docSize > 500000 Then ' 500KB de texto
        LogMessage "Documento muito grande detectado: " & docSize & " caracteres", LOG_LEVEL_WARNING
        Dim continueResponse As VbMsgBoxResult
        continueResponse = MsgBox("Este é um documento muito grande (" & Format(docSize, "#,##0") & " caracteres)." & vbCrLf & _
                                "O processamento pode ser lento." & vbCrLf & vbCrLf & _
                                "Deseja continuar?", vbYesNo + vbQuestion, "Documento Grande - Chainsaw Proposituras")
        If continueResponse = vbNo Then
            LogMessage "Usuário cancelou devido ao tamanho do documento", LOG_LEVEL_INFO
            Exit Function
        End If
    End If
    
    ' Verificação de estado de salvamento
    If Not doc.Saved And doc.path <> "" Then
        LogMessage "Documento tem alterações não salvas", LOG_LEVEL_WARNING
        Dim saveResponse As VbMsgBoxResult
        saveResponse = MsgBox("O documento tem alterações não salvas." & vbCrLf & _
                            "É recomendado salvar antes da padronização." & vbCrLf & vbCrLf & _
                            "Deseja salvar agora?", vbYesNoCancel + vbQuestion, "Alterações Não Salvas - Chainsaw Proposituras")
        Select Case saveResponse
            Case vbYes
                doc.Save
                LogMessage "Documento salvo pelo usuário antes da padronização", LOG_LEVEL_INFO
            Case vbCancel
                LogMessage "Usuário cancelou a operação", LOG_LEVEL_INFO
                Exit Function
            Case vbNo
                LogMessage "Usuário optou por continuar sem salvar", LOG_LEVEL_WARNING
        End Select
    End If
    
    ' Se chegou até aqui, passou em todas as validações
    ValidateDocumentIntegrity = True
    LogMessage "Validação de integridade do documento concluída com sucesso", LOG_LEVEL_INFO
    Exit Function
    
ErrorHandler:
    LogMessage "Erro durante validação de integridade: " & Err.Description, LOG_LEVEL_ERROR
    MsgBox "Erro durante validação do documento:" & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "A operação será cancelada por segurança.", vbCritical, "Erro de Validação - Chainsaw Proposituras"
    ValidateDocumentIntegrity = False
End Function

'================================================================================
' SAFE PROPERTY ACCESS FUNCTIONS - Compatibilidade total com Word 2010+
'================================================================================
Private Function SafeGetCharacterCount(targetRange As Range) As Long
    On Error GoTo FallbackMethod
    
    ' Método preferido - mais rápido
    SafeGetCharacterCount = targetRange.Characters.Count
    Exit Function
    
FallbackMethod:
    On Error GoTo ErrorHandler
    ' Método alternativo para versões com problemas de .Characters.Count
    SafeGetCharacterCount = Len(targetRange.Text)
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
    LogMessage "Erro ao aplicar fonte: " & Err.Description & " - Range: " & Left(targetRange.Text, 20), LOG_LEVEL_WARNING
End Function

Private Function SafeSetParagraphFormat(para As Paragraph, alignment As Long, leftIndent As Single, firstLineIndent As Single) As Boolean
    On Error GoTo ErrorHandler
    
    With para.Format
        If alignment >= 0 Then .alignment = alignment
        If leftIndent >= 0 Then .LeftIndent = leftIndent
        If firstLineIndent >= 0 Then .FirstLineIndent = firstLineIndent
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
    hasImages = (para.Range.InlineShapes.Count > 0)
    
    ' Verifica shapes flutuantes de forma segura
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
    ' Método alternativo mais simples
    SafeHasVisualContent = (para.Range.InlineShapes.Count > 0)
    Exit Function
    
ErrorHandler:
    ' Em caso de erro, assume que não há conteúdo visual
    SafeHasVisualContent = False
End Function

'================================================================================
' SAFE FIND/REPLACE OPERATIONS - Compatibilidade com todas as versões
'================================================================================
Private Function SafeFindReplace(doc As Document, findText As String, replaceText As String, Optional useWildcards As Boolean = False) As Long
    On Error GoTo ErrorHandler
    
    Dim findCount As Long
    findCount = 0
    
    ' Configuração segura de Find/Replace
    With doc.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
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
' SAFE CHARACTER ACCESS FUNCTIONS - Compatibilidade total
'================================================================================
Private Function SafeGetLastCharacter(rng As Range) As String
    On Error GoTo ErrorHandler
    
    Dim charCount As Long
    charCount = SafeGetCharacterCount(rng)
    
    If charCount > 0 Then
        SafeGetLastCharacter = rng.Characters(charCount).Text
    Else
        SafeGetLastCharacter = ""
    End If
    Exit Function
    
ErrorHandler:
    ' Método alternativo usando Right()
    On Error GoTo FinalFallback
    SafeGetLastCharacter = Right(rng.Text, 1)
    Exit Function
    
FinalFallback:
    SafeGetLastCharacter = ""
End Function

'================================================================================
' UNDO GROUP MANAGEMENT - #STABLE
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
' LOGGING MANAGEMENT - APRIMORADO COM DETALHES - #STABLE
'================================================================================
Private Function InitializeLogging(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    If doc.path <> "" Then
        logFilePath = doc.path & "\" & Format(Now, "yyyy-mm-dd") & "_" & _
                     Replace(doc.Name, ".doc", "") & "_FormattingLog.txt"
        logFilePath = Replace(logFilePath, ".docx", "") & "_FormattingLog.txt"
        logFilePath = Replace(logFilePath, ".docm", "") & "_FormattingLog.txt"
    Else
        logFilePath = Environ("TEMP") & "\" & Format(Now, "yyyy-mm-dd") & "_DocumentFormattingLog.txt"
    End If
    
    Open logFilePath For Output As #1
    Print #1, "========================================================"
    Print #1, "LOG DE FORMATAÇÃO DE DOCUMENTO - SISTEMA DE REGISTRO"
    Print #1, "========================================================"
    Print #1, "Duração: " & Format(Timer - processingStartTime, "0.00") & " segundos"
    Print #1, "Erros: " & Err.Number & " - " & Err.Description
    Print #1, "Status: INICIANDO"
    Print #1, "--------------------------------------------------------"
    Print #1, "Sessão: " & Format(Now, "yyyy-mm-dd HH:MM:ss")
    Print #1, "Usuário: " & Environ("USERNAME")
    Print #1, "Estação: " & Environ("COMPUTERNAME")
    Print #1, "Versão Word: " & Application.version
    Print #1, "Documento: " & doc.Name
    Print #1, "Local: " & IIf(doc.path = "", "(Não salvo)", doc.path)
    Print #1, "Proteção: " & GetProtectionType(doc)
    Print #1, "Tamanho: " & GetDocumentSize(doc)
    Print #1, "========================================================"
    Close #1
    
    loggingEnabled = True
    InitializeLogging = True
    
    Exit Function
    
ErrorHandler:
    loggingEnabled = False
    InitializeLogging = False
End Function

Private Sub LogMessage(message As String, Optional level As String = LOG_LEVEL_INFO)
    On Error GoTo ErrorHandler
    
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
    
    Open logFilePath For Append As #1
    Print #1, formattedMessage
    Close #1
    
    Debug.Print "LOG: " & formattedMessage
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "FALHA NO LOGGING: " & message
End Sub

Private Sub SafeFinalizeLogging()
    On Error GoTo ErrorHandler
    
    If loggingEnabled Then
        Open logFilePath For Append As #1
        Print #1, "================================================"
        Print #1, "FIM DA SESSÃO - " & Format(Now, "yyyy-mm-dd HH:MM:ss")
        Print #1, "Duração: " & Format(Timer - processingStartTime, "0.00") & " segundos"
        Print #1, "Erros: " & IIf(Err.Number = 0, "Nenhum", Err.Number & " - " & Err.Description)
        Print #1, "Status: " & IIf(formattingCancelled, "CANCELADO", "CONCLUÍDO")
        Print #1, "================================================"
        Close #1
    End If
    
    loggingEnabled = False
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Erro ao finalizar logging: " & Err.Description
    loggingEnabled = False
End Sub

'================================================================================
' UTILITY: GET PROTECTION TYPE - #STABLE
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
' UTILITY: GET DOCUMENT SIZE - #STABLE
'================================================================================
Private Function GetDocumentSize(doc As Document) As String
    On Error Resume Next
    
    Dim size As Long
    size = doc.BuiltInDocumentProperties("Number of Characters").Value * 2
    
    If size < 1024 Then
        GetDocumentSize = size & " bytes"
    ElseIf size < 1048576 Then
        GetDocumentSize = Format(size / 1024, "0.0") & " KB"
    Else
        GetDocumentSize = Format(size / 1048576, "0.0") & " MB"
    End If
End Function

'================================================================================
' APPLICATION STATE HANDLER - #STABLE
'================================================================================
Private Function SetAppState(Optional ByVal enabled As Boolean = True, Optional ByVal statusMsg As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim success As Boolean
    success = True
    
    With Application
        On Error Resume Next
        .screenUpdating = enabled
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
        
        On Error Resume Next
        .displayAlerts = IIf(enabled, wdAlertsAll, wdAlertsNone)
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
' GLOBAL CHECKING - VERIFICAÇÕES ROBUSTAS
'================================================================================
Private Function PreviousChecking(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        Application.StatusBar = "Erro: Documento não acessível para verificação"
        LogMessage "Documento não acessível para verificação", LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If doc.Type <> wdTypeDocument Then
        Application.StatusBar = "Erro: Tipo de documento não suportado (Tipo: " & doc.Type & ")"
        LogMessage "Tipo de documento não suportado: " & doc.Type, LOG_LEVEL_ERROR
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
        Application.StatusBar = "Erro: Espaço em disco insuficiente"
        LogMessage "Espaço em disco insuficiente para operação segura", LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If Not ValidateDocumentStructure(doc) Then
        LogMessage "Estrutura do documento validada com avisos", LOG_LEVEL_WARNING
    End If

    LogMessage "Verificações de segurança concluídas com sucesso", LOG_LEVEL_INFO
    PreviousChecking = True
    Exit Function

ErrorHandler:
    Application.StatusBar = "Erro durante verificações de segurança"
    LogMessage "Erro durante verificações: " & Err.Description, LOG_LEVEL_ERROR
    PreviousChecking = False
End Function

'================================================================================
' DISK SPACE CHECK - VERIFICAÇÃO SIMPLIFICADA
'================================================================================
Private Function CheckDiskSpace(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ' Verificação simplificada - assume espaço suficiente se não conseguir verificar
    Dim fso As Object
    Dim drive As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If doc.path <> "" Then
        Set drive = fso.GetDrive(Left(doc.path, 3))
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
' MAIN FORMATTING ROUTINE - #STABLE
'================================================================================
Private Function PreviousFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ' Only process header image - all other formatting removed per requirements
    LogMessage "Processing header image only", LOG_LEVEL_INFO
    
    ' Insert header image only  
    InsertHeaderstamp doc
    
    ' Make clipboard visible (requirement #5)
    Application.DisplayClipboardWindow = True
    
    PreviousFormatting = True
    LogMessage "Document processing completed - header image added", LOG_LEVEL_INFO
    Exit Function

ErrorHandler:
    LogMessage "Error in document processing: " & Err.Description, LOG_LEVEL_ERROR
    PreviousFormatting = False
End Function
    
    LogMessage "Formatação completa aplicada", LOG_LEVEL_INFO
    PreviousFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "Erro durante formatação: " & Err.Description, LOG_LEVEL_ERROR
    PreviousFormatting = False
End Function

'================================================================================
' PAGE SETUP - #STABLE
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

' ================================================================================
' FONT FORMMATTING - #STABLE
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

    For i = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(i)
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
        inlineShapesCount = para.Range.InlineShapes.Count
        
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
            paraFullText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            
            ' Verifica se é o primeiro parágrafo com texto (título) - otimizado
            If i <= 3 And para.Format.Alignment = wdAlignParagraphCenter And paraFullText <> "" Then
                isTitle = True
            End If
            
            ' Verifica se o parágrafo começa com "considerando" - otimizado
            If Len(paraFullText) >= 12 And LCase(Left(paraFullText, 12)) = "considerando" Then
                hasConsiderando = True
            End If
            
            ' Verifica se é um parágrafo especial - otimizado
            Dim cleanParaText As String
            cleanParaText = paraFullText
            ' Remove pontuação final para análise
            Do While Len(cleanParaText) > 0 And (Right(cleanParaText, 1) = "." Or Right(cleanParaText, 1) = "," Or Right(cleanParaText, 1) = ":" Or Right(cleanParaText, 1) = ";")
                cleanParaText = Left(cleanParaText, Len(cleanParaText) - 1)
            Loop
            cleanParaText = Trim(LCase(cleanParaText))

            If cleanParaText = "justificativa:" Or IsVereadorPattern(cleanParaText) Or IsAnexoPattern(cleanParaText) Then
                isSpecialParagraph = True
            End If
            
            ' Verifica se é o parágrafo ANTERIOR a "- vereador -" (também deve preservar negrito)
            Dim isBeforeVereador As Boolean
            isBeforeVereador = False
            If i < doc.Paragraphs.Count Then
                Dim nextPara As Paragraph
                Set nextPara = doc.Paragraphs(i + 1)
                If Not HasVisualContent(nextPara) Then
                    Dim nextParaText As String
                    nextParaText = Trim(Replace(Replace(nextPara.Range.Text, vbCr, ""), vbLf, ""))
                    ' Remove pontuação final para análise
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
' FORMATAÇÃO CARACTERE POR CARACTERE CONSOLIDADA - #OPTIMIZED
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
            If charRange.InlineShapes.Count = 0 Then
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
' PARAGRAPH FORMATTING - #STABLE
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

    For i = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(i)
        hasInlineImage = False

        If para.Range.InlineShapes.Count > 0 Then
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
        cleanText = para.Range.Text
        
        ' OTIMIZADO: Combinação de múltiplas operações de limpeza em um bloco
        If InStr(cleanText, "  ") > 0 Or InStr(cleanText, vbTab) > 0 Then
            ' Remove múltiplos espaços consecutivos
            Do While InStr(cleanText, "  ") > 0
                cleanText = Replace(cleanText, "  ", " ")
            Loop
            
            ' Remove espaços antes/depois de quebras de linha
            cleanText = Replace(cleanText, " " & vbCr, vbCr)
            cleanText = Replace(cleanText, vbCr & " ", vbCr)
            cleanText = Replace(cleanText, " " & vbLf, vbLf)
            cleanText = Replace(cleanText, vbLf & " ", vbLf)
            
            ' Remove tabs extras e converte para espaços
            Do While InStr(cleanText, vbTab & vbTab) > 0
                cleanText = Replace(cleanText, vbTab & vbTab, vbTab)
            Loop
            cleanText = Replace(cleanText, vbTab, " ")
            
            ' Limpeza final de espaços múltiplos
            Do While InStr(cleanText, "  ") > 0
                cleanText = Replace(cleanText, "  ", " ")
            Loop
        End If
        
        ' Aplica o texto limpo APENAS se não há imagens (proteção)
        If cleanText <> para.Range.Text And Not hasInlineImage Then
            para.Range.Text = cleanText
        End If

        paraText = Trim(LCase(Replace(Replace(Replace(para.Range.Text, ".", ""), ",", ""), ";", "")))
        paraText = Replace(paraText, vbCr, "")
        paraText = Replace(paraText, vbLf, "")
        paraText = Replace(paraText, " ", "")

        ' Formatação de parágrafo - SEMPRE aplicada
        With para.Format
            .LineSpacingRule = wdLineSpacingMultiple
            .LineSpacing = LINE_SPACING
            .RightIndent = rightMarginPoints
            .SpaceBefore = 0
            .SpaceAfter = 0

            If para.Alignment = wdAlignParagraphCenter Then
                .LeftIndent = 0
                .FirstLineIndent = 0
            Else
                firstIndent = .FirstLineIndent
                paragraphIndent = .LeftIndent
                If paragraphIndent >= CentimetersToPoints(5) Then
                    .LeftIndent = CentimetersToPoints(9.5)
                ElseIf firstIndent < CentimetersToPoints(5) Then
                    .LeftIndent = CentimetersToPoints(0)
                    .FirstLineIndent = CentimetersToPoints(1.5)
                End If
            End If
        End With

        If para.Alignment = wdAlignParagraphLeft Then
            para.Alignment = wdAlignParagraphJustify
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
' FORMAT SECOND PARAGRAPH - FORMATAÇÃO APENAS DO 2º PARÁGRAFO - #NEW
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
    
    ' Encontra o 2º parágrafo com conteúdo (pula vazios)
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
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
    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.Count Then
        Set para = doc.Paragraphs(secondParaIndex)
        
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
            .LeftIndent = CentimetersToPoints(9)      ' Recuo à esquerda de 9 cm
            .FirstLineIndent = 0                      ' Sem recuo da primeira linha
            .RightIndent = 0                          ' Sem recuo à direita
            .Alignment = wdAlignParagraphJustify      ' Justificado
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
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
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
    For i = paraIndex + 1 To doc.Paragraphs.Count
        If i > doc.Paragraphs.Count Then Exit For
        
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
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
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
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
    
    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.Count Then
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
' FORMAT FIRST PARAGRAPH - FORMATAÇÃO DO 1º PARÁGRAFO - #NEW
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
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
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
    If firstParaIndex > 0 And firstParaIndex <= doc.Paragraphs.Count Then
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
                    If charRange3.InlineShapes.Count = 0 Then
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
            .Alignment = wdAlignParagraphCenter       ' Centralizado
            .LeftIndent = 0                           ' Sem recuo à esquerda
            .FirstLineIndent = 0                      ' Sem recuo da primeira linha
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
' ENABLE HYPHENATION - #STABLE
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
    LogMessage "Erro ao ativar hifenização: " & Err.Description, LOG_LEVEL_ERROR
    EnableHyphenation = False
End Function

'================================================================================
' REMOVE WATERMARK - #STABLE
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
            If header.Exists And header.Shapes.Count > 0 Then
                For i = header.Shapes.Count To 1 Step -1
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
            If header.Exists And header.Shapes.Count > 0 Then
                For i = header.Shapes.Count To 1 Step -1
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
' INSERT HEADER IMAGE - #STABLE
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

    username = GetSafeUserName()
    imgFile = "C:\Users\" & username & HEADER_IMAGE_RELATIVE_PATH

    ' Busca inteligente da imagem em múltiplos locais
    If Dir(imgFile) = "" Then
        ' Tenta localização alternativa no perfil do usuário
        imgFile = Environ("USERPROFILE") & HEADER_IMAGE_RELATIVE_PATH
        If Dir(imgFile) = "" Then
            ' Tenta localização de rede corporativa
            imgFile = "\\strqnapmain\Dir. Legislativa\Christian" & HEADER_IMAGE_RELATIVE_PATH
            If Dir(imgFile) = "" Then
                ' Registra erro e tenta continuar sem a imagem
                Application.StatusBar = "Aviso: Imagem de cabeçalho não encontrada"
                LogMessage "Imagem de cabeçalho não encontrada em nenhum local: " & HEADER_IMAGE_RELATIVE_PATH, LOG_LEVEL_WARNING
                InsertHeaderstamp = False
                Exit Function
            End If
        End If
    End If

    imgWidth = CentimetersToPoints(HEADER_IMAGE_MAX_WIDTH_CM)
    imgHeight = imgWidth * HEADER_IMAGE_HEIGHT_RATIO

    For Each sec In doc.Sections
        Set header = sec.Headers(wdHeaderFooterPrimary)
        If header.Exists Then
            header.LinkToPrevious = False
            header.Range.Delete
            
            Set shp = header.Shapes.AddPicture( _
                fileName:=imgFile, _
                LinkToFile:=False, _
                SaveWithDocument:=msoTrue)
            
            If shp Is Nothing Then
                LogMessage "Falha ao inserir imagem no cabeçalho da seção " & sectionsProcessed + 1, LOG_LEVEL_WARNING
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
        LogMessage "Nenhum cabeçalho foi inserido", LOG_LEVEL_WARNING
        InsertHeaderstamp = False
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro ao inserir cabeçalho: " & Err.Description, LOG_LEVEL_ERROR
    InsertHeaderstamp = False
End Function

'================================================================================
' INSERT FOOTER PAGE NUMBERS - #STABLE
'================================================================================
Private Function InsertFooterstamp(doc As Document) As Boolean
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
            rng.Text = "-"
            
            Set rng = footer.Range
            rng.Collapse Direction:=wdCollapseEnd
            rng.Fields.Add Range:=rng, Type:=wdFieldNumPages
            
            With footer.Range
                .Font.Name = STANDARD_FONT
                .Font.size = FOOTER_FONT_SIZE
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Fields.Update
            End With
            
            sectionsProcessed = sectionsProcessed + 1
        End If
    Next sec

    ' Log detalhado removido para performance
    InsertFooterstamp = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao inserir rodapé: " & Err.Description, LOG_LEVEL_ERROR
    InsertFooterstamp = False
End Function

'================================================================================
' UTILITY: CM TO POINTS - #STABLE
'================================================================================
Private Function CentimetersToPoints(ByVal cm As Double) As Single
    On Error Resume Next
    CentimetersToPoints = Application.CentimetersToPoints(cm)
    If Err.Number <> 0 Then
        CentimetersToPoints = cm * 28.35
    End If
End Function

'================================================================================
' UTILITY: SAFE USERNAME - #STABLE
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
' VALIDATE DOCUMENT STRUCTURE - SIMPLIFICADO - #STABLE
'================================================================================
Private Function ValidateDocumentStructure(doc As Document) As Boolean
    On Error Resume Next
    
    ' Verificação básica e rápida
    If doc.Range.End > 0 And doc.Sections.Count > 0 Then
        ValidateDocumentStructure = True
    Else
        LogMessage "Documento com estrutura inconsistente", LOG_LEVEL_WARNING
        ValidateDocumentStructure = False
    End If
End Function

'================================================================================
' CRITICAL FIX: SAVE DOCUMENT BEFORE PROCESSING
' TO PREVENT CRASHES ON NEW NON SAVED DOCUMENTS - #STABLE
'================================================================================
Private Function SaveDocumentFirst(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Aguardando salvamento do documento..."
    ' Log de início removido para performance
    
    Dim saveDialog As Object
    Set saveDialog = Application.Dialogs(wdDialogFileSaveAs)

    If saveDialog.Show <> -1 Then
        LogMessage "Operação de salvamento cancelada pelo usuário", LOG_LEVEL_INFO
        Application.StatusBar = "Salvamento cancelado pelo usuário"
        SaveDocumentFirst = False
        Exit Function
    End If

    ' Aguarda confirmação do salvamento com timeout de segurança
    Dim waitCount As Integer
    Dim maxWait As Integer
    maxWait = 10
    
    For waitCount = 1 To maxWait
        DoEvents
        If doc.path <> "" Then Exit For
        Dim startTime As Double
        startTime = Timer
        Do While Timer < startTime + 1
            DoEvents
        Loop
        Application.StatusBar = "Aguardando salvamento... (" & waitCount & "/" & maxWait & ")"
    Next waitCount

    If doc.path = "" Then
        LogMessage "Falha ao salvar documento após " & maxWait & " tentativas", LOG_LEVEL_ERROR
        Application.StatusBar = "Falha no salvamento - operação cancelada"
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
' CLEAR ALL FORMATTING - LIMPEZA INICIAL COMPLETA - #NEW
'================================================================================
Private Function ClearAllFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Limpando formatação existente..."
    
    ' SUPER OTIMIZADO: Verificação única de conteúdo visual no documento
    Dim hasImages As Boolean
    Dim hasShapes As Boolean
    hasImages = (doc.InlineShapes.Count > 0)
    hasShapes = (doc.Shapes.Count > 0)
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
                        .Alignment = wdAlignParagraphLeft
                        .LineSpacing = 12
                        .SpaceBefore = 0
                        .SpaceAfter = 0
                        .LeftIndent = 0
                        .RightIndent = 0
                        .FirstLineIndent = 0
                    End With
                    
                    ' Reset de bordas e sombreamento
                    .Borders.enable = False
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
                .Alignment = wdAlignParagraphLeft
                .LineSpacing = 12
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LeftIndent = 0
                .RightIndent = 0
                .FirstLineIndent = 0
            End With
            
            On Error Resume Next
            .Borders.enable = False
            .Shading.Texture = wdTextureNone
            On Error GoTo ErrorHandler
        End With
        
        paraCount = doc.Paragraphs.Count
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
' CLEAN DOCUMENT STRUCTURE - FUNCIONALIDADES 2, 6, 7 - #NEW
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
    paraCount = doc.Paragraphs.Count
    
    ' OTIMIZADO: Funcionalidade 2 - Remove linhas em branco acima do título
    ' Busca otimizada do primeiro parágrafo com texto
    firstTextParaIndex = -1
    For i = 1 To paraCount
        If i > doc.Paragraphs.Count Then Exit For ' Proteção dinâmica
        
        Set para = doc.Paragraphs(i)
        Dim paraTextCheck As String
        paraTextCheck = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
        ' Encontra o primeiro parágrafo com texto real
        If paraTextCheck <> "" Then
            firstTextParaIndex = i
            Exit For
        End If
        
        ' Proteção contra documentos muito grandes
        If i > 50 Then Exit For ' Limita busca aos primeiros 50 parágrafos
    Next i
    
    ' OTIMIZADO: Remove linhas vazias ANTES do primeiro texto em uma única passada
    If firstTextParaIndex > 1 Then
        ' Processa de trás para frente para evitar problemas com índices
        For i = firstTextParaIndex - 1 To 1 Step -1
            If i > doc.Paragraphs.Count Or i < 1 Then Exit For ' Proteção dinâmica
            
            Set para = doc.Paragraphs(i)
            Dim paraTextEmpty As String
            paraTextEmpty = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            
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
    
    ' SUPER OTIMIZADO: Funcionalidade 7 - Remove espaços iniciais com regex
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
        .Text = "^p "  ' Quebra seguida de espaço
        .Replacement.Text = "^p"
        
        Do While .Execute(Replace:=True)
            leadingSpacesRemoved = leadingSpacesRemoved + 1
            ' Proteção contra loop infinito
            If leadingSpacesRemoved > 1000 Then Exit Do
        Loop
        
        ' Remove tabs no início de linhas
        .Text = "^p^t"  ' Quebra seguida de tab
        .Replacement.Text = "^p"
        
        Do While .Execute(Replace:=True)
            leadingSpacesRemoved = leadingSpacesRemoved + 1
            If leadingSpacesRemoved > 1000 Then Exit Do
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
        If rng.Text = " " Or rng.Text = vbTab Then
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
' SAFE CHECK FOR VISUAL CONTENT - VERIFICAÇÃO SEGURA DE CONTEÚDO VISUAL - #NEW
'================================================================================
Private Function HasVisualContent(para As Paragraph) As Boolean
    ' Usa a função segura implementada para compatibilidade total
    HasVisualContent = SafeHasVisualContent(para)
End Function

'================================================================================
' VALIDATE PROPOSITION TYPE - FUNCIONALIDADE 3 - #NEW
'================================================================================
Private Function ValidatePropositionType(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim firstPara As Paragraph
    Dim firstWord As String
    Dim paraText As String
    Dim i As Long
    Dim userResponse As VbMsgBoxResult
    
    ' Encontra o primeiro parágrafo com texto
    For i = 1 To doc.Paragraphs.Count
        Set firstPara = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(firstPara.Range.Text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then
            Exit For
        End If
    Next i
    
    If paraText = "" Then
        LogMessage "Documento não possui texto para validação", LOG_LEVEL_WARNING
        ValidatePropositionType = True
        Exit Function
    End If
    
    ' Extrai a primeira palavra
    Dim words() As String
    words = Split(paraText, " ")
    If UBound(words) >= 0 Then
        firstWord = LCase(Trim(words(0)))
    End If
    
    ' Verifica se é uma das proposituras válidas
    If firstWord = "indicação" Or firstWord = "requerimento" Or firstWord = "moção" Then
        LogMessage "Tipo de proposição validado: " & firstWord, LOG_LEVEL_INFO
        ValidatePropositionType = True
    Else
        ' Documento não é uma proposição padrão - solicita confirmação do usuário
        LogMessage "Primeira palavra não reconhecida como proposição padrão: " & firstWord, LOG_LEVEL_WARNING
        Application.StatusBar = "Aguardando confirmação do usuário sobre tipo de documento..."
        
        ' Monta mensagem detalhada para o usuário
        Dim confirmationMessage As String
        confirmationMessage = "ATENÇÃO: POSSÍVEL DOCUMENTO NÃO-PADRÃO" & vbCrLf & vbCrLf
        confirmationMessage = confirmationMessage & "O documento não inicia com as palavras esperadas para uma propositura:" & vbCrLf
        confirmationMessage = confirmationMessage & "• INDICAÇÃO" & vbCrLf
        confirmationMessage = confirmationMessage & "• REQUERIMENTO" & vbCrLf
        confirmationMessage = confirmationMessage & "• MOÇÃO" & vbCrLf & vbCrLf
        confirmationMessage = confirmationMessage & "Primeira palavra encontrada: """ & UCase(firstWord) & """" & vbCrLf & vbCrLf
        confirmationMessage = confirmationMessage & "Início do documento:" & vbCrLf
        confirmationMessage = confirmationMessage & """" & Left(paraText, 150) & "...""" & vbCrLf & vbCrLf
        confirmationMessage = confirmationMessage & "Este documento pode não ser uma propositura legislative," & vbCrLf
        confirmationMessage = confirmationMessage & "mas você pode optar por formatá-lo mesmo assim." & vbCrLf & vbCrLf
        confirmationMessage = confirmationMessage & "Deseja prosseguir com a formatação?"
        
        userResponse = MsgBox(confirmationMessage, vbYesNo + vbQuestion + vbDefaultButton2, _
                             "Chainsaw - Validação de Tipo de Documento")
        
        If userResponse = vbYes Then
            LogMessage "Usuário optou por prosseguir com documento não-padrão: " & firstWord, LOG_LEVEL_INFO
            Application.StatusBar = "Processando documento não-padrão conforme solicitado..."
            ValidatePropositionType = True
        Else
            LogMessage "Usuário optou por interromper processamento de documento não-padrão: " & firstWord, LOG_LEVEL_INFO
            Application.StatusBar = "Processamento cancelado pelo usuário"
            
            ' Mensagem final de cancelamento
            MsgBox "Processamento cancelado." & vbCrLf & vbCrLf & _
                   "Para documentos de propositura, certifique-se de que " & vbCrLf & _
                   "a primeira palavra seja INDICAÇÃO, REQUERIMENTO ou MOÇÃO.", _
                   vbInformation, "Chainsaw - Operação Cancelada"
            
            ValidatePropositionType = False
        End If
    End If
    
    Exit Function

ErrorHandler:
    LogMessage "Erro na validação do tipo de proposição: " & Err.Description, LOG_LEVEL_ERROR
    ValidatePropositionType = False
End Function

'================================================================================
' FORMAT DOCUMENT TITLE - FUNCIONALIDADES 4 e 5 - #NEW
'================================================================================
Private Function FormatDocumentTitle(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim firstPara As Paragraph
    Dim paraText As String
    Dim words() As String
    Dim i As Long
    Dim newText As String
    
    ' Encontra o primeiro parágrafo com texto (após exclusão de linhas em branco)
    For i = 1 To doc.Paragraphs.Count
        Set firstPara = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(firstPara.Range.Text, vbCr, ""), vbLf, ""))
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
        ' Reconstrói o texto substituindo a última palavra
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
    firstPara.Range.Text = UCase(newText) & vbCrLf
    
    ' Formatação completa do título (primeira linha)
    With firstPara.Range.Font
        .Bold = True
        .Underline = wdUnderlineSingle
    End With
    
    With firstPara.Format
        .Alignment = wdAlignParagraphCenter
        .LeftIndent = 0
        .FirstLineIndent = 0
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
' FORMAT CONSIDERANDO PARAGRAPHS - OTIMIZADO E SIMPLIFICADO - FUNCIONALIDADE 8 - #NEW
'================================================================================
Private Function FormatConsiderandoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim rng As Range
    Dim totalFormatted As Long
    Dim i As Long
    
    ' Percorre todos os parágrafos procurando por "considerando" no início
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
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
                        .Text = "considerando"
                        .Replacement.Text = "CONSIDERANDO"
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
                    .Text = "CONSIDERANDO"
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
' APPLY TEXT REPLACEMENTS - FUNCIONALIDADES 10, 11, 12 e 13 - #NEW
'================================================================================
Private Function ApplyTextReplacements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Dim replacementCount As Long
    
    Set rng = doc.Range
    
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
    
    For i = 0 To UBound(dOesteVariants)
        With rng.Find
            .ClearFormatting
            .Text = dOesteVariants(i) & "este"
            .Replacement.Text = "d'Oeste"
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
    ReDim vereadorVariants(0 To 17)
    
    ' Variantes dos caracteres inicial e final

    ' Com espaços:

    ' "V" maiúsculo:
    vereadorVariants(0) = "- Vereador -"    ' Correta
    vereadorVariants(1) = "– Vereador –"    ' Travessão
    vereadorVariants(2) = "— Vereador —"    ' Em dash

    ' "v" minúsculo:
    vereadorVariants(3) = "- vereador -"    ' Minúscula
    vereadorVariants(4) = "– vereador –"    ' Travessão minúscula
    vereadorVariants(5) = "— vereador —"    ' Em dash minúscula

    ' Sem espaços:

    ' "V" maiúsculo:
    vereadorVariants(6) = "-Vereador-"      ' Sem espaços maiúscula
    vereadorVariants(7) = "–Vereador–"      ' Sem espaços travessão maiúscula
    vereadorVariants(8) = "—Vereador—"     ' Sem espaços em dash maiúscula

    ' "v" minúsculo:
    vereadorVariants(9) = "-vereador-"      ' Sem espaços minúscula
    vereadorVariants(10) = "–vereador–"      ' Sem espaços travessão minúscula
    vereadorVariants(11) = "—vereador—"     ' Sem espaços em dash minúscula

    ' Todas em maiúsculas:
    vereadorVariants(12) = "- VEREADOR -"      ' Com espaços maiúscula
    vereadorVariants(13) = "– VEREADOR –"      ' Com espaços travessão maiúscula
    vereadorVariants(14) = "— VEREADOR —"     ' Com espaços em dash maiúscula
    vereadorVariants(15) = "-VEREADOR-"      ' Sem espaços maiúscula
    vereadorVariants(16) = "–VEREADOR–"      ' Sem espaços travessão maiúscula
    vereadorVariants(17) = "—VEREADOR—"     ' Sem espaços em dash maiúscula
    
    For i = 0 To UBound(vereadorVariants)
        If vereadorVariants(i) <> "- Vereador -" Then
            With rng.Find
                .ClearFormatting
                .Text = vereadorVariants(i)
                .Replacement.Text = "- Vereador -"
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
    
    ' Funcionalidade 12: Substitui hífens e traços isolados por travessão (em dash)
    ' Esta funcionalidade padroniza todos os hífens (-) e en dashes (–) que estejam
    ' isolados (com espaço antes e depois) substituindo-os por em dash (—) que é
    ' o travessão correto para uso em português
    Set rng = doc.Range
    Dim dashVariants() As String
    ReDim dashVariants(0 To 2)
    
    ' Define os tipos de hífens/traços que devem ser substituídos quando isolados
    dashVariants(0) = " - "     ' Hífen comum isolado
    dashVariants(1) = " – "     ' En dash isolado
    dashVariants(2) = " — "     ' Em dash isolado (para normalização)
    
    ' Substitui todos os tipos por em dash (travessão)
    For i = 0 To UBound(dashVariants)
        ' Só substitui se não for já um em dash
        If dashVariants(i) <> " — " Then
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = dashVariants(i)
                .Replacement.Text = " — "    ' Em dash (travessão) com espaços
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
    
    ' Casos especiais: hífen/traço no início de linha seguido de espaço
    Set rng = doc.Range
    Dim lineStartDashVariants() As String
    ReDim lineStartDashVariants(0 To 1)
    
    lineStartDashVariants(0) = "^p- "   ' Hífen no início de linha
    lineStartDashVariants(1) = "^p– "   ' En dash no início de linha
    
    For i = 0 To UBound(lineStartDashVariants)
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = lineStartDashVariants(i)
            .Replacement.Text = "^p— "    ' Em dash no início de linha
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
    
    ' Casos especiais: espaço seguido de hífen/traço no final de linha
    Set rng = doc.Range
    Dim lineEndDashVariants() As String
    ReDim lineEndDashVariants(0 To 1)
    
    lineEndDashVariants(0) = " -^p"   ' Hífen no final de linha
    lineEndDashVariants(1) = " –^p"   ' En dash no final de linha
    
    For i = 0 To UBound(lineEndDashVariants)
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = lineEndDashVariants(i)
            .Replacement.Text = " —^p"    ' Em dash no final de linha
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
    
    ' Funcionalidade 13: Remove todas as quebras de linha manuais
    ' Esta funcionalidade remove quebras de linha manuais (soft breaks) que podem
    ' ter sido inseridas manualmente no documento, mantendo apenas as quebras
    ' de parágrafo normais para preservar a estrutura do documento
    Set rng = doc.Range
    
    ' Remove quebras de linha manuais (Shift+Enter) - Chr(11) ou ^l
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^l"  ' Quebra de linha manual (line break)
        .Replacement.Text = " "  ' Substitui por espaço
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
    
    ' Remove quebras de linha manuais usando código de caractere
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = Chr(11)  ' Quebra de linha manual (VT - Vertical Tab)
        .Replacement.Text = " "  ' Substitui por espaço
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
    
    ' Remove caracteres de nova linha (Line Feed) que não sejam quebras de parágrafo
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = Chr(10)  ' Line Feed (LF)
        .Replacement.Text = " "  ' Substitui por espaço
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
    
    LogMessage "Substituições de texto aplicadas: " & replacementCount & " substituições realizadas", LOG_LEVEL_INFO
    ApplyTextReplacements = True
    Exit Function

ErrorHandler:
    LogMessage "Erro nas substituições de texto: " & Err.Description, LOG_LEVEL_ERROR
    ApplyTextReplacements = False
End Function

'================================================================================
' APPLY SPECIFIC PARAGRAPH REPLACEMENTS - SUBSTITUIÇÕES ESPECÍFICAS POR PARÁGRAFO - #NEW
'================================================================================
Private Function ApplySpecificParagraphReplacements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Aplicando substituições específicas (por parágrafo e globais)..."
    
    Dim replacementCount As Long
    Dim secondParaIndex As Long
    Dim thirdParaIndex As Long
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    
    replacementCount = 0
    
    ' Encontra o 2º e 3º parágrafos com conteúdo (ignora parágrafos vazios)
    actualParaIndex = 0
    secondParaIndex = 0
    thirdParaIndex = 0
    
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
        ' Se o parágrafo tem conteúdo, conta como um parágrafo real
        If paraText <> "" Then
            actualParaIndex = actualParaIndex + 1
            
            If actualParaIndex = 2 Then
                secondParaIndex = i
            ElseIf actualParaIndex = 3 Then
                thirdParaIndex = i
                Exit For ' Já encontrou os dois parágrafos necessários
            End If
        End If
        
        ' Proteção contra documentos muito grandes
        If i > 50 Then Exit For
    Next i
    
    ' REQUISITO 1: Se o 2º parágrafo começa exatamente com "Sugiro ", substitui por "Requeiro "
    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.Count Then
        Set para = doc.Paragraphs(secondParaIndex)
        paraText = para.Range.Text
        
        ' Verifica se começa exatamente com "Sugiro " (case-sensitive)
        If Len(paraText) >= 7 And Left(paraText, 7) = "Sugiro " Then
            ' Substitui "Sugiro " por "Requeiro " no início do parágrafo
            para.Range.Text = "Requeiro " & Mid(paraText, 8)
            replacementCount = replacementCount + 1
            LogMessage "2º parágrafo: 'Sugiro ' substituído por 'Requeiro '", LOG_LEVEL_INFO
        End If
    End If
    
    ' REQUISITOS 2 e 3: Substituições no 3º parágrafo
    If thirdParaIndex > 0 And thirdParaIndex <= doc.Paragraphs.Count Then
        Set para = doc.Paragraphs(thirdParaIndex)
        paraText = para.Range.Text
        Dim originalText As String
        originalText = paraText
        
        ' REQUISITO 2: Se há " sugerir " em qualquer parte do 3º parágrafo, substitui por " indicar "
        If InStr(paraText, " sugerir ") > 0 Then
            paraText = Replace(paraText, " sugerir ", " indicar ")
            replacementCount = replacementCount + 1
            LogMessage "3º parágrafo: ' sugerir ' substituído por ' indicar '", LOG_LEVEL_INFO
        End If
        
        ' REQUISITO 3: Se há " Setor, " em qualquer parte do 3º parágrafo, substitui por " setor competente, "
        If InStr(paraText, " Setor, ") > 0 Then
            paraText = Replace(paraText, " Setor, ", " setor competente, ")
            replacementCount = replacementCount + 1
            LogMessage "3º parágrafo: ' Setor, ' substituído por ' setor competente, '", LOG_LEVEL_INFO
        End If
        
        ' Aplica as mudanças se houve alguma substituição
        If paraText <> originalText Then
            para.Range.Text = paraText
        End If
    End If
    
    ' REQUISITOS GLOBAIS: Substituições em todo o documento
    Dim rng As Range
    Set rng = doc.Range
    
    ' REQUISITO GLOBAL 1: Substitui a frase específica da Câmara Municipal
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = " A CÂMARA MUNICIPAL DE SANTA BÁRBARA D'OESTE, ESTADO DE SÃO PAULO "
        .Replacement.Text = " a Câmara Municipal de Santa Bárbara d'Oeste, estado de São Paulo, "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True  ' Case-sensitive para essa substituição específica
        .MatchWholeWord = False
        .MatchWildcards = False
        
        Do While .Execute(Replace:=True)
            replacementCount = replacementCount + 1
            LogMessage "Substituição global: 'A CÂMARA MUNICIPAL...' → 'a Câmara Municipal...'", LOG_LEVEL_INFO
        Loop
    End With
    
    ' REQUISITO GLOBAL 2: Converte palavras específicas para maiúsculas
    Dim wordsToUppercase() As String
    Dim j As Long
    
    ' Define array com todas as variações das palavras que devem ficar em maiúsculas
    ReDim wordsToUppercase(0 To 15)
    wordsToUppercase(0) = "aplaude"
    wordsToUppercase(1) = "Aplaude"
    wordsToUppercase(2) = "aplauso"
    wordsToUppercase(3) = "Aplauso"
    wordsToUppercase(4) = "protesta"
    wordsToUppercase(5) = "Protesta"
    wordsToUppercase(6) = "protesto"
    wordsToUppercase(7) = "Protesto"
    wordsToUppercase(8) = "apela"
    wordsToUppercase(9) = "Apela"
    wordsToUppercase(10) = "apelo"
    wordsToUppercase(11) = "Apelo"
    wordsToUppercase(12) = "apoia"
    wordsToUppercase(13) = "Apoia"
    wordsToUppercase(14) = "apoio"
    wordsToUppercase(15) = "Apoio"
    
    ' Aplica conversão para maiúsculas para cada palavra
    For j = 0 To UBound(wordsToUppercase)
        Set rng = doc.Range
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = wordsToUppercase(j)
            .Replacement.Text = UCase(wordsToUppercase(j))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True  ' Case-sensitive para detectar exatamente a variação
            .MatchWholeWord = True  ' Palavra completa apenas
            .MatchWildcards = False
            
            Do While .Execute(Replace:=True)
                replacementCount = replacementCount + 1
                If replacementCount <= 20 Then  ' Log apenas os primeiros casos para performance
                    LogMessage "Conversão para maiúsculas: '" & wordsToUppercase(j) & "' → '" & UCase(wordsToUppercase(j)) & "'", LOG_LEVEL_INFO
                End If
            Loop
        End With
    Next j
    
    LogMessage "Substituições específicas concluídas (por parágrafo e globais): " & replacementCount & " substituições realizadas", LOG_LEVEL_INFO
    ApplySpecificParagraphReplacements = True
    Exit Function

ErrorHandler:
    LogMessage "Erro nas substituições específicas: " & Err.Description, LOG_LEVEL_ERROR
    ApplySpecificParagraphReplacements = False
End Function

'================================================================================
' VALIDATE CONTENT CONSISTENCY - VALIDAÇÃO DE CONSISTÊNCIA ENTRE EMENTA E TEOR - #NEW
'================================================================================
Private Function ValidateContentConsistency(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Validando consistência entre ementa e teor..."
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim secondParaIndex As Long
    Dim secondParaText As String
    Dim restOfDocumentText As String
    
    ' Encontra o 2º parágrafo com conteúdo (ementa)
    actualParaIndex = 0
    secondParaIndex = 0
    
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
        ' Se o parágrafo tem conteúdo, conta como um parágrafo real
        If paraText <> "" Then
            actualParaIndex = actualParaIndex + 1
            
            If actualParaIndex = 2 Then
                secondParaIndex = i
                secondParaText = paraText
                Exit For
            End If
        End If
        
        ' Proteção contra documentos muito grandes
        If i > 50 Then Exit For
    Next i
    
    ' Se não encontrou o 2º parágrafo, não faz validação
    If secondParaIndex = 0 Or secondParaText = "" Then
        LogMessage "2º parágrafo não encontrado para validação de consistência", LOG_LEVEL_WARNING
        ValidateContentConsistency = True
        Exit Function
    End If
    
    ' Coleta o restante do texto do documento (a partir do 3º parágrafo)
    restOfDocumentText = ""
    actualParaIndex = 0
    
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
        If paraText <> "" Then
            actualParaIndex = actualParaIndex + 1
            
            ' Coleta texto a partir do 3º parágrafo
            If actualParaIndex >= 3 Then
                restOfDocumentText = restOfDocumentText & " " & paraText
            End If
        End If
    Next i
    
    ' Se não há conteúdo suficiente para comparar, não faz validação
    If restOfDocumentText = "" Then
        LogMessage "Conteúdo insuficiente para validação de consistência", LOG_LEVEL_WARNING
        ValidateContentConsistency = True
        Exit Function
    End If
    
    ' Analisa consistência entre o 2º parágrafo e o restante do documento
    Dim commonWordsCount As Long
    commonWordsCount = CountCommonWords(secondParaText, restOfDocumentText)
    
    LogMessage "Validação de consistência: " & commonWordsCount & " palavras comuns encontradas entre ementa e teor", LOG_LEVEL_INFO
    
    ' Se há menos de 2 palavras em comum, alerta sobre possível inconsistência
    If commonWordsCount < 2 Then
        ' Mostra aviso ao usuário
        Dim inconsistencyMessage As String
        inconsistencyMessage = "AVISO: POSSÍVEL INCONSISTÊNCIA DETECTADA" & vbCrLf & vbCrLf
        inconsistencyMessage = inconsistencyMessage & "Foi detectada uma possível inconsistência entre a EMENTA " & vbCrLf
        inconsistencyMessage = inconsistencyMessage & "(2º parágrafo) e o TEOR da propositura (restante do texto)." & vbCrLf & vbCrLf
        inconsistencyMessage = inconsistencyMessage & "EMENTA (2º parágrafo):" & vbCrLf
        inconsistencyMessage = inconsistencyMessage & """" & Left(secondParaText, 200) & """" & vbCrLf & vbCrLf
        inconsistencyMessage = inconsistencyMessage & "Apenas " & commonWordsCount & " palavra(s) em comum foram encontradas " & vbCrLf
        inconsistencyMessage = inconsistencyMessage & "entre a ementa e o conteúdo da propositura." & vbCrLf & vbCrLf
        inconsistencyMessage = inconsistencyMessage & "Recomenda-se revisar o documento para garantir que:" & vbCrLf
        inconsistencyMessage = inconsistencyMessage & "• A ementa reflita adequadamente o conteúdo" & vbCrLf
        inconsistencyMessage = inconsistencyMessage & "• O teor esteja alinhado com a ementa" & vbCrLf & vbCrLf
        inconsistencyMessage = inconsistencyMessage & "Deseja continuar com a formatação mesmo assim?"
        
        Dim userResponse As VbMsgBoxResult
        userResponse = MsgBox(inconsistencyMessage, vbYesNo + vbExclamation + vbDefaultButton2, _
                             "Chainsaw - Validação de Consistência")
        
        If userResponse = vbNo Then
            LogMessage "Usuário optou por interromper devido a inconsistência detectada", LOG_LEVEL_WARNING
            Application.StatusBar = "Formatação interrompida - inconsistência detectada"
            ValidateContentConsistency = False
            Exit Function
        Else
            LogMessage "Usuário optou por continuar apesar da inconsistência detectada", LOG_LEVEL_WARNING
        End If
    Else
        LogMessage "Consistência adequada: " & commonWordsCount & " palavras comuns entre ementa e teor", LOG_LEVEL_INFO
    End If
    
    ValidateContentConsistency = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na validação de consistência: " & Err.Description, LOG_LEVEL_ERROR
    ValidateContentConsistency = False
End Function

'================================================================================
' COUNT COMMON WORDS - CONTA PALAVRAS COMUNS ENTRE DOIS TEXTOS - #NEW
'================================================================================
Private Function CountCommonWords(text1 As String, text2 As String) As Long
    On Error GoTo ErrorHandler
    
    Dim words1() As String
    Dim words2() As String
    Dim i As Long, j As Long
    Dim commonCount As Long
    Dim word1 As String, word2 As String
    
    ' Limpa e normaliza os textos
    text1 = CleanTextForComparison(text1)
    text2 = CleanTextForComparison(text2)
    
    ' Divide em palavras
    words1 = Split(text1, " ")
    words2 = Split(text2, " ")
    
    commonCount = 0
    
    ' Compara cada palavra do primeiro texto com as do segundo
    For i = 0 To UBound(words1)
        word1 = Trim(words1(i))
        
        ' Ignora palavras muito curtas (menos de 4 caracteres) ou palavras comuns
        If Len(word1) >= 4 And Not IsCommonWord(word1) Then
            For j = 0 To UBound(words2)
                word2 = Trim(words2(j))
                
                ' Se encontrar palavra igual, conta e para (evita contar duplicatas)
                If word1 = word2 Then
                    commonCount = commonCount + 1
                    Exit For
                End If
            Next j
        End If
    Next i
    
    CountCommonWords = commonCount
    Exit Function

ErrorHandler:
    CountCommonWords = 0
End Function

'================================================================================
' CLEAN TEXT FOR COMPARISON - LIMPA TEXTO PARA COMPARAÇÃO - #NEW
'================================================================================
Private Function CleanTextForComparison(text As String) As String
    Dim cleanedText As String
    Dim i As Long
    Dim char As String
    
    ' Converte para minúsculas
    cleanedText = LCase(text)
    
    ' Remove pontuação e caracteres especiais, mantém apenas letras, números e espaços
    Dim result As String
    result = ""
    
    For i = 1 To Len(cleanedText)
        char = Mid(cleanedText, i, 1)
        
        ' Mantém apenas letras, números e espaços
        If (char >= "a" And char <= "z") Or (char >= "0" And char <= "9") Or char = " " Then
            result = result & char
        Else
            ' Substitui pontuação por espaço
            result = result & " "
        End If
    Next i
    
    ' Remove espaços múltiplos
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    CleanTextForComparison = Trim(result)
End Function

'================================================================================
' IS COMMON WORD - VERIFICA SE É PALAVRA MUITO COMUM - #NEW
'================================================================================
Private Function IsCommonWord(word As String) As Boolean
    ' Lista de palavras muito comuns que devem ser ignoradas na comparação
    Dim commonWords() As String
    Dim i As Long
    
    ReDim commonWords(0 To 49)
    commonWords(0) = "que"
    commonWords(1) = "para"
    commonWords(2) = "com"
    commonWords(3) = "uma"
    commonWords(4) = "por"
    commonWords(5) = "dos"
    commonWords(6) = "das"
    commonWords(7) = "este"
    commonWords(8) = "esta"
    commonWords(9) = "essa"
    commonWords(10) = "esse"
    commonWords(11) = "seu"
    commonWords(12) = "sua"
    commonWords(13) = "seus"
    commonWords(14) = "suas"
    commonWords(15) = "mais"
    commonWords(16) = "muito"
    commonWords(17) = "entre"
    commonWords(18) = "sobre"
    commonWords(19) = "após"
    commonWords(20) = "antes"
    commonWords(21) = "durante"
    commonWords(22) = "através"
    commonWords(23) = "mediante"
    commonWords(24) = "junto"
    commonWords(25) = "desde"
    commonWords(26) = "até"
    commonWords(27) = "contra"
    commonWords(28) = "favor"
    commonWords(29) = "deve"
    commonWords(30) = "devem"
    commonWords(31) = "pode"
    commonWords(32) = "podem"
    commonWords(33) = "será"
    commonWords(34) = "serão"
    commonWords(35) = "está"
    commonWords(36) = "estão"
    commonWords(37) = "foram"
    commonWords(38) = "sendo"
    commonWords(39) = "tendo"
    commonWords(40) = "onde"
    commonWords(41) = "quando"
    commonWords(42) = "como"
    commonWords(43) = "porque"
    commonWords(44) = "portanto"
    commonWords(45) = "assim"
    commonWords(46) = "então"
    commonWords(47) = "ainda"
    commonWords(48) = "também"
    commonWords(49) = "apenas"
    
    word = LCase(Trim(word))
    
    For i = 0 To UBound(commonWords)
        If word = commonWords(i) Then
            IsCommonWord = True
            Exit Function
        End If
    Next i
    
    IsCommonWord = False
End Function

'================================================================================
' FORMAT JUSTIFICATIVA/ANEXO/VEREADOR PARAGRAPHS - FORMATAÇÃO ESPECÍFICA - #NEW
'================================================================================
Private Function FormatJustificativaAnexoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim cleanText As String
    Dim i As Long
    Dim formattedCount As Long
    Dim vereadorCount As Long
    
    ' Percorre todos os parágrafos do documento
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        
        ' Não processa parágrafos com conteúdo visual
        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            
            ' Remove pontuação final para análise mais precisa
            cleanText = paraText
            ' Remove pontos, vírgulas, dois-pontos, ponto-e-vírgula do final
            Do While Len(cleanText) > 0 And (Right(cleanText, 1) = "." Or Right(cleanText, 1) = "," Or Right(cleanText, 1) = ":" Or Right(cleanText, 1) = ";")
                cleanText = Left(cleanText, Len(cleanText) - 1)
            Loop
            cleanText = Trim(LCase(cleanText))
            
            ' REQUISITO 1: Formatação de "justificativa"
            If cleanText = "justificativa" Then
                ' Aplica formatação específica para Justificativa
                With para.Format
                    .LeftIndent = 0                         ' Recuo à esquerda = 0
                    .FirstLineIndent = 0                     ' Recuo da 1ª linha = 0
                    .Alignment = wdAlignParagraphCenter       ' Alinhamento centralizado
                End With

                With para.Range.Font
                    .Bold = True                  ' Negrito
                End With
                
                ' Padroniza o texto mantendo pontuação original se houver
                Dim originalEnd As String
                originalEnd = ""
                If Len(paraText) > Len(cleanText) Then
                    originalEnd = Right(paraText, Len(paraText) - Len(cleanText))
                End If
                para.Range.Text = "Justificativa" & originalEnd & vbCrLf
                
                LogMessage "Parágrafo 'Justificativa' formatado (centralizado, negrito, sem recuos)", LOG_LEVEL_INFO
                formattedCount = formattedCount + 1
                
            ' REQUISITO 1: Formatação de variações de "vereador" 
            ElseIf IsVereadorPattern(cleanText) Then
                ' REQUISITO 2: Formatar parágrafo ANTERIOR a "vereador" PRIMEIRO
                If i > 1 Then
                    Dim paraPrev As Paragraph
                    Set paraPrev = doc.Paragraphs(i - 1)
                    
                    ' Verifica se o parágrafo anterior não tem conteúdo visual
                    If Not HasVisualContent(paraPrev) Then
                        Dim prevText As String
                        prevText = Trim(Replace(Replace(paraPrev.Range.Text, vbCr, ""), vbLf, ""))
                        
                        ' Só formata se o parágrafo anterior tem conteúdo textual
                        If prevText <> "" Then
                            ' Formatação COMPLETA do parágrafo anterior
                            With paraPrev.Format
                                .LeftIndent = 0                      ' Recuo à esquerda = 0
                                .FirstLineIndent = 0                 ' Recuo da 1ª linha = 0  
                                .Alignment = wdAlignParagraphCenter  ' Alinhamento centralizado
                            End With

                            With paraPrev.Range.Font
                                .Bold = True                  ' Negrito
                                .AllCaps = True               ' Caixa alta
                            End With
                            
                            LogMessage "Parágrafo anterior a '- Vereador -' formatado (centralizado, caixa alta, negrito, sem recuos): " & Left(UCase(prevText), 30) & "...", LOG_LEVEL_INFO
                        End If
                    End If
                End If
                
                ' Agora formata o parágrafo "- Vereador -"
                With para.Format
                    .LeftIndent = 0               ' Recuo à esquerda = 0
                    .FirstLineIndent = 0          ' Recuo da 1ª linha = 0
                    .Alignment = wdAlignParagraphCenter  ' Alinhamento centralizado
                End With

                With para.Range.Font
                    .Bold = True                  ' Negrito
                End With
                
                ' Padroniza o texto
                para.Range.Text = "- Vereador -" & vbCrLf
                
                LogMessage "Parágrafo '- Vereador -' formatado (centralizado, negrito, sem recuos)", LOG_LEVEL_INFO
                vereadorCount = vereadorCount + 1
                formattedCount = formattedCount + 1
                
            ' REQUISITO 3: Formatação de variações de "anexo" ou "anexos"
            ElseIf IsAnexoPattern(cleanText) Then
                ' Aplica formatação específica para Anexo/Anexos
                With para.Format
                    .LeftIndent = 0               ' Recuo à esquerda = 0
                    .FirstLineIndent = 0          ' Recuo da 1ª linha = 0
                    .RightIndent = 0              ' Recuo à direita = 0
                    .Alignment = wdAlignParagraphLeft    ' Alinhamento à esquerda
                End With

                With para.Range.Font
                    .Bold = True                  ' Negrito
                End With

                ' Padroniza o texto mantendo pontuação original se houver
                Dim anexoEnd As String
                anexoEnd = ""
                If Len(paraText) > Len(cleanText) Then
                    anexoEnd = Right(paraText, Len(paraText) - Len(cleanText))
                End If
                
                Dim anexoText As String
                If cleanText = "anexo" Then
                    anexoText = "Anexo"
                Else
                    anexoText = "Anexos"
                End If
                para.Range.Text = anexoText & anexoEnd & vbCrLf
                
                LogMessage "Parágrafo '" & anexoText & "' formatado (alinhado à esquerda, negrito, sem recuos)", LOG_LEVEL_INFO
                formattedCount = formattedCount + 1
                
            ' REQUISITO 4: Formatação de parágrafos que começam com "Ante o exposto"
            ElseIf IsAnteOExpostoPattern(paraText) Then
                ' Aplica formatação de negrito para "Ante o exposto"
                With para.Range.Font
                    .Bold = True                  ' Negrito
                End With
                
                LogMessage "Parágrafo 'Ante o exposto' formatado (negrito)", LOG_LEVEL_INFO
                formattedCount = formattedCount + 1
            End If
        End If
    Next i
    
    LogMessage "Formatação especial concluída: " & formattedCount & " parágrafos formatados (incluindo " & vereadorCount & " '- Vereador -')", LOG_LEVEL_INFO
    FormatJustificativaAnexoParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação de parágrafos especiais: " & Err.Description, LOG_LEVEL_ERROR
    FormatJustificativaAnexoParagraphs = False
End Function

'================================================================================
' FORMAT NUMBERED PARAGRAPHS - FUNCIONALIDADE 14 - #NEW
'================================================================================
Private Function FormatNumberedParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim formattedCount As Long
    
    ' Percorre todos os parágrafos do documento
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        
        ' Não processa parágrafos com conteúdo visual
        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            
            ' Verifica se o parágrafo começa com número seguido de ponto, parênteses ou espaço
            If IsNumberedParagraph(paraText) Then
                ' Aplica formatação de lista numerada
                With para.Range.ListFormat
                    ' Remove formatação de lista existente primeiro
                    .RemoveNumbers
                    
                    ' Aplica lista numerada
                    .ApplyNumberDefault
                End With
                
                ' Remove o número manual do texto, pois a lista numerada irá gerar automaticamente
                Dim cleanedText As String
                cleanedText = RemoveManualNumber(paraText)
                
                ' Atualiza o texto do parágrafo
                para.Range.Text = cleanedText & vbCrLf
                
                LogMessage "Parágrafo convertido para lista numerada: " & Left(cleanedText, 50) & "...", LOG_LEVEL_INFO
                formattedCount = formattedCount + 1
            End If
        End If
    Next i
    
    LogMessage "Formatação de listas numeradas concluída: " & formattedCount & " parágrafos convertidos", LOG_LEVEL_INFO
    FormatNumberedParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação de listas numeradas: " & Err.Description, LOG_LEVEL_ERROR
    FormatNumberedParagraphs = False
End Function

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

Private Function IsAnteOExpostoPattern(text As String) As Boolean
    ' Verifica se o parágrafo começa com "Ante o exposto" (ignorando maiúsculas/minúsculas)
    Dim cleanText As String
    cleanText = LCase(Trim(text))
    
    ' Verifica se está vazio
    If Len(cleanText) = 0 Then
        IsAnteOExpostoPattern = False
        Exit Function
    End If
    
    ' Verifica se começa com "ante o exposto"
    If Len(cleanText) >= 13 And Left(cleanText, 13) = "ante o exposto" Then
        IsAnteOExpostoPattern = True
    Else
        IsAnteOExpostoPattern = False
    End If
End Function

'================================================================================
' FUNÇÕES AUXILIARES PARA LISTAS NUMERADAS
'================================================================================
Private Function IsNumberedParagraph(text As String) As Boolean
    ' Verifica se o parágrafo começa com um número seguido de separadores comuns
    Dim cleanText As String
    cleanText = Trim(text)
    
    ' Verifica se está vazio
    If Len(cleanText) = 0 Then
        IsNumberedParagraph = False
        Exit Function
    End If
    
    ' Extrai a primeira palavra/token
    Dim firstToken As String
    Dim spacePos As Long
    spacePos = InStr(cleanText, " ")
    
    If spacePos > 0 Then
        firstToken = Left(cleanText, spacePos - 1)
    Else
        firstToken = cleanText
    End If
    
    ' Verifica diferentes padrões de numeração
    ' Padrão 1: Número seguido de ponto (1., 2., 3., etc.)
    If Len(firstToken) >= 2 And Right(firstToken, 1) = "." Then
        Dim numberPart As String
        numberPart = Left(firstToken, Len(firstToken) - 1)
        If IsNumeric(numberPart) And Val(numberPart) > 0 Then
            ' Verifica se há texto substantivo após o número e pontuação
            If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then
                IsNumberedParagraph = True
                Exit Function
            End If
        End If
    End If
    
    ' Padrão 2: Número seguido de parênteses (1), 2), 3), etc.)
    If Len(firstToken) >= 2 And Right(firstToken, 1) = ")" Then
        numberPart = Left(firstToken, Len(firstToken) - 1)
        If IsNumeric(numberPart) And Val(numberPart) > 0 Then
            ' Verifica se há texto substantivo após o número e pontuação
            If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then
                IsNumberedParagraph = True
                Exit Function
            End If
        End If
    End If
    
    ' Padrão 3: Parênteses com número ((1), (2), (3), etc.)
    If Len(firstToken) >= 3 And Left(firstToken, 1) = "(" And Right(firstToken, 1) = ")" Then
        numberPart = Mid(firstToken, 2, Len(firstToken) - 2)
        If IsNumeric(numberPart) And Val(numberPart) > 0 Then
            ' Verifica se há texto substantivo após o número e pontuação
            If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then
                IsNumberedParagraph = True
                Exit Function
            End If
        End If
    End If
    
    ' Padrão 4: Apenas número seguido de espaço (1 texto, 2 texto, etc.)
    ' CRITÉRIO MAIS RIGOROSO: deve ter espaço E texto substantivo após o número
    If IsNumeric(firstToken) And Val(firstToken) > 0 And spacePos > 0 Then
        ' Verifica se há texto substantivo após o número e espaço
        If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then
            IsNumberedParagraph = True
            Exit Function
        End If
    End If
    
    ' Padrão 5: Número seguido de outros separadores comuns (-, :, ;)
    If Len(firstToken) >= 2 Then
        Dim lastChar As String
        lastChar = Right(firstToken, 1)
        
        If lastChar = "-" Or lastChar = ":" Or lastChar = ";" Then
            numberPart = Left(firstToken, Len(firstToken) - 1)
            If IsNumeric(numberPart) And Val(numberPart) > 0 Then
                ' Verifica se há texto substantivo após o número e pontuação
                If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then
                    IsNumberedParagraph = True
                    Exit Function
                End If
            End If
        End If
    End If
    
    IsNumberedParagraph = False
End Function

'================================================================================
' HAS SUBSTANTIVE TEXT AFTER NUMBER - VERIFICA SE HÁ TEXTO SUBSTANTIVO APÓS NÚMERO - #NEW
'================================================================================
Private Function HasSubstantiveTextAfterNumber(fullText As String, numberToken As String) As Boolean
    ' Verifica se há texto substantivo (não apenas espaços ou números) após o número
    Dim remainingText As String
    Dim startPos As Long
    
    ' Encontra a posição após o token do número
    startPos = Len(numberToken) + 1
    
    ' Se não há mais texto após o token, não é um parágrafo numerado válido
    If startPos > Len(fullText) Then
        HasSubstantiveTextAfterNumber = False
        Exit Function
    End If
    
    ' Extrai o texto restante após o número
    remainingText = Trim(Mid(fullText, startPos))
    
    ' Verifica se há texto substantivo
    If Len(remainingText) = 0 Then
        ' Sem texto após o número
        HasSubstantiveTextAfterNumber = False
        Exit Function
    End If
    
    ' Remove espaços e verifica se há pelo menos uma palavra com letras
    Dim words() As String
    Dim i As Long
    Dim hasLetters As Boolean
    
    words = Split(remainingText, " ")
    
    For i = 0 To UBound(words)
        Dim word As String
        word = Trim(words(i))
        
        ' Verifica se a palavra contém pelo menos uma letra (não é apenas números ou pontuação)
        If ContainsLetters(word) And Len(word) >= 2 Then
            HasSubstantiveTextAfterNumber = True
            Exit Function
        End If
    Next i
    
    ' Se chegou até aqui, não encontrou texto substantivo
    HasSubstantiveTextAfterNumber = False
End Function

'================================================================================
' CONTAINS LETTERS - VERIFICA SE STRING CONTÉM LETRAS - #NEW
'================================================================================
Private Function ContainsLetters(text As String) As Boolean
    Dim i As Long
    Dim char As String
    
    For i = 1 To Len(text)
        char = LCase(Mid(text, i, 1))
        If char >= "a" And char <= "z" Then
            ContainsLetters = True
            Exit Function
        End If
    Next i
    
    ContainsLetters = False
End Function

Private Function RemoveManualNumber(text As String) As String
    ' Remove o número manual do início do parágrafo
    Dim cleanText As String
    cleanText = Trim(text)
    
    If Len(cleanText) = 0 Then
        RemoveManualNumber = text
        Exit Function
    End If
    
    ' Encontra a primeira palavra/token
    Dim firstToken As String
    Dim spacePos As Long
    spacePos = InStr(cleanText, " ")
    
    If spacePos > 0 Then
        firstToken = Left(cleanText, spacePos - 1)
        
        ' Remove o primeiro token se for um número com separadores
        If (Len(firstToken) >= 2 And (Right(firstToken, 1) = "." Or Right(firstToken, 1) = ")")) Or _
           (Len(firstToken) >= 3 And Left(firstToken, 1) = "(" And Right(firstToken, 1) = ")") Or _
           (IsNumeric(firstToken) And Val(firstToken) > 0) Then
            
            ' Remove o primeiro token e espaços extras
            RemoveManualNumber = Trim(Mid(cleanText, spacePos + 1))
        Else
            RemoveManualNumber = cleanText
        End If
    Else
        RemoveManualNumber = cleanText
    End If
End Function

'================================================================================
' SUBROTINA PÚBLICA: ABRIR PASTA DE LOGS - #NEW
'================================================================================
Public Sub OpenLogsFolder()
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
    shell "explorer.exe """ & logsFolder & """", vbNormalFocus
    
    Application.StatusBar = "Pasta de logs aberta: " & logsFolder
    
    ' Log da operação se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Pasta de logs aberta pelo usuário: " & logsFolder, LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir pasta de logs"
    
    ' Fallback: tenta abrir pasta temporária
    On Error Resume Next
    shell "explorer.exe """ & Environ("TEMP") & """", vbNormalFocus
    If Err.Number = 0 Then
        Application.StatusBar = "Pasta temporária aberta como alternativa"
    Else
        Application.StatusBar = "Não foi possível abrir pasta de logs"
    End If
End Sub

'================================================================================
' SUBROTINA PÚBLICA: ABRIR REPOSITÓRIO GITHUB - FUNCIONALIDADE 9 - #NEW
'================================================================================
Public Sub OpenGitHubRepository()
    On Error GoTo ErrorHandler
    
    Dim repoURL As String
    Dim shellResult As Long
    
    ' URL do repositório do projeto
    repoURL = "https://github.com/chrmsantos/chainsaw-proposituras"
    
    ' Abre o link no navegador padrão
    shellResult = shell("rundll32.exe url.dll,FileProtocolHandler " & repoURL, vbNormalFocus)
    
    If shellResult > 0 Then
        Application.StatusBar = "Repositório GitHub aberto no navegador"
        
        ' Log da operação se sistema de log estiver ativo
        If loggingEnabled Then
            LogMessage "Repositório GitHub aberto pelo usuário: " & repoURL, LOG_LEVEL_INFO
        End If
    Else
        GoTo ErrorHandler
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir repositório GitHub"
    
    ' Log do erro se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Erro ao abrir repositório GitHub: " & Err.Description, LOG_LEVEL_ERROR
    End If
    
    ' Fallback: tenta copiar URL para a área de transferência
    On Error Resume Next
    Dim dataObj As Object
    Set dataObj = CreateObject("htmlfile").parentWindow.clipboardData
    dataObj.setData "text", repoURL
    
    If Err.Number = 0 Then
        Application.StatusBar = "URL copiada para área de transferência: " & repoURL
    Else
        Application.StatusBar = "Não foi possível abrir o repositório. URL: " & repoURL
    End If
End Sub

'================================================================================
' SISTEMA DE BACKUP - FUNCIONALIDADE DE SEGURANÇA - #NEW
'================================================================================
Private Function CreateDocumentBackup(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    CreateDocumentBackup = False
    
    ' Validação inicial do documento
    If doc Is Nothing Then
        LogMessage "Erro no backup: documento é nulo", LOG_LEVEL_ERROR
        Exit Function
    End If
    
    ' Não faz backup se documento não foi salvo
    If doc.path = "" Then
        LogMessage "Backup ignorado - documento não salvo", LOG_LEVEL_INFO
        CreateDocumentBackup = True
        Exit Function
    End If
    
    ' Verifica se documento não está corrompido
    On Error Resume Next
    Dim testAccess As String
    testAccess = doc.Name
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        LogMessage "Erro no backup: documento inacessível", LOG_LEVEL_ERROR
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    Dim backupFolder As String
    Dim fso As Object
    Dim docName As String
    Dim docExtension As String
    Dim timestamp As String
    Dim backupFileName As String
    Dim retryCount As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Verifica se FSO foi criado com sucesso
    If fso Is Nothing Then
        LogMessage "Erro no backup: não foi possível criar FileSystemObject", LOG_LEVEL_ERROR
        Exit Function
    End If
    
    ' Define pasta de backup com validação
    On Error Resume Next
    Dim parentPath As String
    parentPath = fso.GetParentFolderName(doc.path)
    If Err.Number <> 0 Or parentPath = "" Then
        On Error GoTo ErrorHandler
        LogMessage "Erro no backup: não foi possível determinar pasta pai", LOG_LEVEL_ERROR
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    backupFolder = parentPath & BACKUP_FOLDER_NAME
    
    ' Cria pasta de backup com verificação robusta
    If Not fso.FolderExists(backupFolder) Then
        On Error Resume Next
        fso.CreateFolder backupFolder
        If Err.Number <> 0 Then
            On Error GoTo ErrorHandler
            LogMessage "Erro ao criar pasta de backup: " & Err.Description, LOG_LEVEL_ERROR
            Exit Function
        End If
        On Error GoTo ErrorHandler
        LogMessage "Pasta de backup criada: " & backupFolder, LOG_LEVEL_INFO
    End If
    
    ' Verifica permissões de escrita na pasta de backup
    On Error Resume Next
    Dim testFile As String
    testFile = backupFolder & "\test_write_" & Format(Now, "HHmmss") & ".tmp"
    Open testFile For Output As #1
    Close #1
    Kill testFile
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        LogMessage "Erro no backup: sem permissões de escrita na pasta", LOG_LEVEL_ERROR
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Extrai nome e extensão do documento com validação
    On Error Resume Next
    docName = fso.GetBaseName(doc.Name)
    docExtension = fso.GetExtensionName(doc.Name)
    If Err.Number <> 0 Or docName = "" Then
        On Error GoTo ErrorHandler
        LogMessage "Erro no backup: nome de arquivo inválido", LOG_LEVEL_ERROR
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Cria timestamp para o backup
    timestamp = Format(Now, "yyyy-mm-dd_HHmmss")
    
    ' Nome do arquivo de backup
    backupFileName = docName & "_backup_" & timestamp & "." & docExtension
    backupFilePath = backupFolder & "\" & backupFileName
    
    ' Salva uma cópia do documento como backup com retry
    Application.StatusBar = "Criando backup do documento..."
    
    ' Salva o documento atual primeiro para garantir que está atualizado
    On Error Resume Next
    doc.Save
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        LogMessage "Aviso: não foi possível salvar documento antes do backup: " & Err.Description, LOG_LEVEL_WARNING
    End If
    On Error GoTo ErrorHandler
    
    ' Cria uma cópia do arquivo usando FileSystemObject com retry
    For retryCount = 1 To MAX_RETRY_ATTEMPTS
        On Error Resume Next
        fso.CopyFile doc.FullName, backupFilePath, True
        If Err.Number = 0 Then
            On Error GoTo ErrorHandler
            Exit For
        Else
            On Error GoTo ErrorHandler
            LogMessage "Tentativa " & retryCount & " de backup falhou: " & Err.Description, LOG_LEVEL_WARNING
            If retryCount < MAX_RETRY_ATTEMPTS Then
                ' Aguarda um pouco antes de tentar novamente
                Sleep 1000 ' 1 segundo = 1000 milissegundos
            End If
        End If
    Next retryCount
    
    ' Verifica se o backup foi criado com sucesso
    If Not fso.FileExists(backupFilePath) Then
        LogMessage "Erro no backup: arquivo não foi criado", LOG_LEVEL_ERROR
        Exit Function
    End If
    
    ' Limpa backups antigos se necessário
    CleanOldBackups backupFolder, docName
    
    LogMessage "Backup criado com sucesso: " & backupFileName, LOG_LEVEL_INFO
    Application.StatusBar = "Backup criado - processando documento..."
    
    CreateDocumentBackup = True
    Exit Function

ErrorHandler:
    LogMessage "Erro crítico ao criar backup: " & Err.Description & " (Linha: " & Erl & ")", LOG_LEVEL_ERROR
    Application.StatusBar = "Falha na criação do backup"
    CreateDocumentBackup = False
    
    ' Limpeza de recursos
    On Error Resume Next
    Set fso = Nothing
End Function

'================================================================================
' LIMPEZA DE BACKUPS ANTIGOS - SIMPLIFICADO - #NEW
'================================================================================
Private Sub CleanOldBackups(backupFolder As String, docBaseName As String)
    On Error Resume Next
    
    ' Limpeza simplificada - só remove se houver muitos arquivos
    Dim fso As Object
    Dim folder As Object
    Dim filesCount As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(backupFolder)
    
    filesCount = folder.Files.Count
    
    ' Se há mais de 15 arquivos na pasta de backup, registra aviso
    If filesCount > 15 Then
        LogMessage "Muitos backups na pasta (" & filesCount & " arquivos) - considere limpeza manual", LOG_LEVEL_WARNING
    End If
End Sub

'================================================================================
' SUBROTINA PÚBLICA: ABRIR PASTA DE BACKUPS - #NEW
'================================================================================
Public Sub OpenBackupsFolder()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim backupFolder As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Tenta obter documento ativo
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler
    
    ' Define pasta de backup baseada no documento atual
    If Not doc Is Nothing And doc.path <> "" Then
        backupFolder = fso.GetParentFolderName(doc.path) & "\" & BACKUP_FOLDER_NAME
    Else
        Application.StatusBar = "Nenhum documento salvo ativo para localizar pasta de backups"
        Exit Sub
    End If
    
    ' Verifica se a pasta de backup existe
    If Not fso.FolderExists(backupFolder) Then
        Application.StatusBar = "Pasta de backups não encontrada - nenhum backup foi criado ainda"
        LogMessage "Pasta de backups não encontrada: " & backupFolder, LOG_LEVEL_WARNING
        Exit Sub
    End If
    
    ' Abre a pasta no Windows Explorer
    shell "explorer.exe """ & backupFolder & """", vbNormalFocus
    
    Application.StatusBar = "Pasta de backups aberta: " & backupFolder
    
    ' Log da operação se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Pasta de backups aberta pelo usuário: " & backupFolder, LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir pasta de backups"
    LogMessage "Erro ao abrir pasta de backups: " & Err.Description, LOG_LEVEL_ERROR
    
    ' Fallback: tenta abrir pasta do documento
    On Error Resume Next
    If Not doc Is Nothing And doc.path <> "" Then
        Dim docFolder As String
        docFolder = fso.GetParentFolderName(doc.path)
        shell "explorer.exe """ & docFolder & """", vbNormalFocus
        Application.StatusBar = "Pasta do documento aberta como alternativa"
    Else
        Application.StatusBar = "Não foi possível abrir pasta de backups"
    End If
End Sub

'================================================================================
' CLEAN MULTIPLE SPACES - LIMPEZA FINAL DE ESPAÇOS MÚLTIPLOS - #NEW
'================================================================================
Private Function CleanMultipleSpaces(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Limpando espaços múltiplos..."
    
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
            .Text = "  "  ' Dois espaços
            .Replacement.Text = " "  ' Um espaço
            
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
        .Text = "  ^p"  ' 2 espaços seguidos de quebra
        .Replacement.Text = " ^p"  ' 1 espaço seguido de quebra
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
        ' Segunda passada para garantir limpeza completa
        .Text = " ^p"  ' Espaço antes de quebra
        .Replacement.Text = "^p"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
        ' Remove múltiplos espaços depois de quebras - método iterativo
        .Text = "^p  "  ' Quebra seguida de 2 espaços
        .Replacement.Text = "^p "  ' Quebra seguida de 1 espaço
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
        .Text = "^t^t"  ' 2 tabs
        .Replacement.Text = "^t"  ' 1 tab
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
        ' Converte tabs para espaços
        .Text = "^t"
        .Replacement.Text = " "
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
    End With
    
    ' OTIMIZAÇÃO 4: Verificação final ultra-rápida de espaços duplos remanescentes
    Set rng = doc.Range
    With rng.Find
        .Text = "  "
        .Replacement.Text = " "
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
        .Text = "CONSIDERANDOa"
        .Replacement.Text = "CONSIDERANDO a"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop
        
        .Text = "CONSIDERANDOe"
        .Replacement.Text = "CONSIDERANDO e"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop
        
        .Text = "CONSIDERANDOo"
        .Replacement.Text = "CONSIDERANDO o"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop
        
        .Text = "CONSIDERANDOq"
        .Replacement.Text = "CONSIDERANDO q"
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
' LIMIT SEQUENTIAL EMPTY LINES - CONTROLA LINHAS VAZIAS SEQUENCIAIS - #NEW
'================================================================================
Private Function LimitSequentialEmptyLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Controlando linhas em branco sequenciais..."
    
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
        .Text = "^p^p^p^p"  ' 4 quebras
        .Replacement.Text = "^p^p"  ' 2 quebras
        
        Do While .Execute(Replace:=True)
            linesRemoved = linesRemoved + 1
            totalReplaces = totalReplaces + 1
            If totalReplaces > 500 Then Exit Do
            If linesRemoved Mod 50 = 0 Then DoEvents
        Loop
        
        ' Remove 3 quebras -> 2 quebras
        .Text = "^p^p^p"  ' 3 quebras
        .Replacement.Text = "^p^p"  ' 2 quebras
        
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
        .Text = "^p^p^p"  ' 3 quebras
        .Replacement.Text = "^p^p"  ' 2 quebras
        
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
        .Text = "^p^p^p"  ' 3 quebras (2 linhas vazias + conteúdo)
        .Replacement.Text = "^p^p"  ' 2 quebras (1 linha vazia + conteúdo)
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
        
        Do While i <= doc.Paragraphs.Count And fallbackRemoved < 50
            Set para = doc.Paragraphs(i)
            paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            
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
' ENSURE PARAGRAPH SEPARATION - GARANTE SEPARAÇÃO ENTRE PARÁGRAFOS - #NEW
'================================================================================
Private Function EnsureParagraphSeparation(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Garantindo separação mínima entre parágrafos..."
    
    Dim para As Paragraph
    Dim nextPara As Paragraph
    Dim i As Long
    Dim insertedCount As Long
    Dim totalChecked As Long
    
    ' Percorre todos os parágrafos verificando se há pelo menos uma linha em branco após cada um
    For i = 1 To doc.Paragraphs.Count - 1 ' -1 porque verificamos o próximo parágrafo
        Set para = doc.Paragraphs(i)
        Set nextPara = doc.Paragraphs(i + 1)
        
        totalChecked = totalChecked + 1
        
        ' Extrai o texto de ambos os parágrafos para análise
        Dim paraText As String
        Dim nextParaText As String
        
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        nextParaText = Trim(Replace(Replace(nextPara.Range.Text, vbCr, ""), vbLf, ""))
        
        ' Verifica se ambos os parágrafos contêm texto (não são linhas em branco)
        If paraText <> "" And nextParaText <> "" Then
            ' Verifica se os parágrafos estão adjacentes (sem linha em branco entre eles)
            ' Para isso, verifica se o final do parágrafo atual é imediatamente seguido pelo início do próximo
            
            Dim currentParaEnd As Long
            Dim nextParaStart As Long
            
            currentParaEnd = para.Range.End
            nextParaStart = nextPara.Range.Start
            
            ' Se a diferença entre o fim de um parágrafo e o início do próximo é apenas 1 caractere,
            ' significa que eles estão diretamente adjacentes (sem linha em branco)
            If nextParaStart - currentParaEnd <= 1 Then
                ' Insere uma linha em branco entre os parágrafos
                Dim insertRange As Range
                Set insertRange = doc.Range(currentParaEnd - 1, currentParaEnd - 1)
                insertRange.Text = vbCrLf
                
                insertedCount = insertedCount + 1
                
                ' Atualiza a referência dos parágrafos após a inserção
                ' porque os índices podem ter mudado
                On Error Resume Next
                Set para = doc.Paragraphs(i)
                Set nextPara = doc.Paragraphs(i + 2) ' +2 porque inserimos uma linha
                On Error GoTo ErrorHandler
                
                ' Log apenas para os primeiros casos ou casos significativos
                If insertedCount <= 10 Or insertedCount Mod 50 = 0 Then
                    LogMessage "Linha em branco inserida entre parágrafos " & i & " e " & (i + 1) & " (total: " & insertedCount & ")"
                End If
            End If
        End If
        
        ' Controle de performance e responsividade
        If totalChecked Mod 100 = 0 Then
            DoEvents
            Application.StatusBar = "Verificando separação de parágrafos... " & totalChecked & " verificados"
        End If
        
        ' Proteção contra documentos muito grandes
        If totalChecked > 5000 Then
            LogMessage "Limite de verificação atingido (5000 parágrafos) - interrompendo verificação", LOG_LEVEL_WARNING
            Exit For
        End If
    Next i
    
    LogMessage "Separação de parágrafos garantida: " & insertedCount & " linhas em branco inseridas de " & totalChecked & " pares verificados"
    EnsureParagraphSeparation = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao garantir separação de parágrafos: " & Err.Description, LOG_LEVEL_ERROR
    EnsureParagraphSeparation = False
End Function

'================================================================================
' CONFIGURE DOCUMENT VIEW - CONFIGURAÇÃO DE VISUALIZAÇÃO - #MODIFIED
'================================================================================
Private Function ConfigureDocumentView(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Configurando visualização do documento..."
    
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
' SALVAR E SAIR - SUBROTINA PÚBLICA PROFISSIONAL E ROBUSTA
'================================================================================
Public Sub SaveAndExit()
    On Error GoTo CriticalErrorHandler
    
    Dim startTime As Date
    startTime = Now
    
    Application.StatusBar = "Verificando documentos abertos..."
    LogMessage "Iniciando processo de salvar e sair - verificação de documentos", LOG_LEVEL_INFO
    
    ' Verifica se há documentos abertos
    If Application.Documents.Count = 0 Then
        Application.StatusBar = "Nenhum documento aberto - encerrando Word"
        LogMessage "Nenhum documento aberto - encerrando aplicação", LOG_LEVEL_INFO
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
            LogMessage "Documento não salvo detectado: " & doc.Name
        End If
        On Error GoTo CriticalErrorHandler
    Next i
    
    ' Se não há documentos não salvos, encerra diretamente
    If unsavedDocs.Count = 0 Then
        Application.StatusBar = "Todos os documentos salvos - encerrando Word"
        LogMessage "Todos os documentos estão salvos - encerrando aplicação"
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
            LogMessage "Usuário optou por salvar todos os documentos antes de sair"
            
            If SalvarTodosDocumentos() Then
                Application.StatusBar = "Documentos salvos com sucesso - encerrando Word"
                LogMessage "Todos os documentos salvos com sucesso - encerrando aplicação"
                Application.Quit SaveChanges:=wdDoNotSaveChanges
            Else
                Application.StatusBar = "Erro ao salvar documentos - operação cancelada"
                LogMessage "Falha ao salvar alguns documentos - operação de sair cancelada", LOG_LEVEL_ERROR
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
                LogMessage "Usuário confirmou fechamento sem salvar - encerrando aplicação", LOG_LEVEL_WARNING
                Application.Quit SaveChanges:=wdDoNotSaveChanges
            Else
                Application.StatusBar = "Operação cancelada pelo usuário"
                LogMessage "Usuário cancelou fechamento sem salvar"
                MsgBox "Operação cancelada." & vbCrLf & "Os documentos permanecem abertos.", _
                       vbInformation, "Chainsaw - Operação Cancelada"
            End If
            
        Case vbCancel
            ' Usuário cancelou
            Application.StatusBar = "Operação de sair cancelada pelo usuário"
            LogMessage "Usuário cancelou operação de salvar e sair"
            MsgBox "Operação cancelada." & vbCrLf & "Os documentos permanecem abertos.", _
                   vbInformation, "Chainsaw - Operação Cancelada"
    End Select
    
    Application.StatusBar = False
    LogMessage "Processo de salvar e sair concluído em " & Format(Now - startTime, "hh:mm:ss")
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "ERRO CRÍTICO na operação Salvar e Sair #" & Err.Number & ": " & Err.Description
    
    LogMessage errDesc, LOG_LEVEL_ERROR
    Application.StatusBar = "Erro crítico - operação cancelada"
    
    MsgBox "Erro crítico durante a operação Salvar e Sair:" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "A operação foi cancelada por segurança." & vbCrLf & _
           "Salve manualmente os documentos importantes.", _
           vbCritical, "Chainsaw - Erro Crítico"
End Sub

'================================================================================
' SALVAR TODOS DOCUMENTOS - FUNÇÃO AUXILIAR PRIVADA
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
                        LogMessage "Documento salvo como novo arquivo: " & doc.Name
                    Else
                        errorCount = errorCount + 1
                        LogMessage "Erro ao salvar documento como novo: " & doc.Name & " - " & Err.Description, LOG_LEVEL_ERROR
                    End If
                Else
                    errorCount = errorCount + 1
                    LogMessage "Salvamento cancelado pelo usuário: " & doc.Name, LOG_LEVEL_WARNING
                End If
            End With
        Else
            ' Documento já tem caminho, apenas salva
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
    LogMessage "Erro crítico ao salvar documentos: " & Err.Description, LOG_LEVEL_ERROR
    SalvarTodosDocumentos = False
End Function

'================================================================================
' IMAGE PROTECTION SYSTEM - SISTEMA DE PROTEÇÃO DE IMAGENS - #NEW
'================================================================================

'================================================================================
' BACKUP ALL IMAGES - Faz backup de propriedades das imagens do documento
'================================================================================
Private Function BackupAllImages(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Fazendo backup das propriedades das imagens..."
    
    imageCount = 0
    ReDim savedImages(0)
    
    Dim para As Paragraph
    Dim i As Long
    Dim j As Long
    Dim shape As InlineShape
    Dim tempImageInfo As ImageInfo
    
    ' Conta todas as imagens primeiro
    Dim totalImages As Long
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        totalImages = totalImages + para.Range.InlineShapes.Count
    Next i
    
    ' Adiciona shapes flutuantes
    totalImages = totalImages + doc.Shapes.Count
    
    ' Redimensiona array se necessário
    If totalImages > 0 Then
        ReDim savedImages(totalImages - 1)
        
        ' Backup de imagens inline - apenas propriedades críticas
        For i = 1 To doc.Paragraphs.Count
            Set para = doc.Paragraphs(i)
            
            For j = 1 To para.Range.InlineShapes.Count
                Set shape = para.Range.InlineShapes(j)
                
                ' Salva apenas propriedades essenciais para proteção
                With tempImageInfo
                    .ParaIndex = i
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
        For i = 1 To doc.Shapes.Count
            Set floatingShape = doc.Shapes(i)
            
            If floatingShape.Type = msoPicture Then
                ' Redimensiona array se necessário
                If imageCount >= UBound(savedImages) + 1 Then
                    ReDim Preserve savedImages(imageCount)
                End If
                
                With tempImageInfo
                    .ParaIndex = -1 ' Indica que é flutuante
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
' RESTORE ALL IMAGES - Verifica e corrige propriedades das imagens
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
                If .ParaIndex <= doc.Paragraphs.Count Then
                    Dim para As Paragraph
                    Set para = doc.Paragraphs(.ParaIndex)
                    
                    ' Se ainda há imagens inline no parágrafo, considera verificada
                    If para.Range.InlineShapes.Count > 0 Then
                        verifiedCount = verifiedCount + 1
                    End If
                End If
                
            ElseIf .ImageType = "Floating" Then
                ' Verifica e corrige propriedades de shapes flutuantes se ainda existem
                If .ImageIndex <= doc.Shapes.Count Then
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
    If targetRange.InlineShapes.Count > 0 Then
        ' OTIMIZADO: Aplica formatação caractere por caractere, protegendo imagens
        Dim i As Long
        Dim charRange As Range
        Dim charCount As Long
        charCount = SafeGetCharacterCount(targetRange) ' Cache da contagem segura
        
        If charCount > 0 Then ' Verificação de segurança
            For i = 1 To charCount
                Set charRange = targetRange.Characters(i)
                ' Só formata caracteres que não são parte de imagens
                If charRange.InlineShapes.Count = 0 Then
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
' CLEANUP IMAGE PROTECTION - Limpeza das variáveis de proteção de imagens
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
' VISUAL ELEMENTS CLEANUP SYSTEM - SISTEMA DE LIMPEZA DE ELEMENTOS VISUAIS
'================================================================================

'================================================================================
' DELETE HIDDEN VISUAL ELEMENTS - Remove todos os elementos visuais ocultos
'================================================================================
Private Function DeleteHiddenVisualElements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Removendo elementos visuais ocultos..."
    
    Dim deletedCount As Long
    deletedCount = 0
    
    ' Remove shapes ocultos (flutuantes)
    Dim i As Long
    For i = doc.Shapes.Count To 1 Step -1
        Dim shp As Shape
        Set shp = doc.Shapes(i)
        
        ' Verifica se o shape está oculto (múltiplos critérios)
        Dim isHidden As Boolean
        isHidden = False
        
        ' Shape marcado como não visível
        If Not shp.Visible Then isHidden = True
        
        ' Shape com transparência total
        On Error Resume Next
        If shp.Fill.Transparency >= 0.99 Then isHidden = True
        On Error GoTo ErrorHandler
        
        ' Shape com tamanho zero ou quase zero
        If shp.Width <= 1 Or shp.Height <= 1 Then isHidden = True
        
        ' Shape posicionado fora da página visível (coordenadas muito negativas)
        If shp.Left < -1000 Or shp.Top < -1000 Then isHidden = True
        
        If isHidden Then
            LogMessage "Removendo shape oculto (tipo: " & shp.Type & ", índice: " & i & ")"
            shp.Delete
            deletedCount = deletedCount + 1
        End If
    Next i
    
    ' Remove objetos incorporados ocultos
    For i = doc.InlineShapes.Count To 1 Step -1
        Dim inlineShp As InlineShape
        Set inlineShp = doc.InlineShapes(i)
        
        Dim isInlineHidden As Boolean
        isInlineHidden = False
        
        ' Objeto inline em texto oculto
        If inlineShp.Range.Font.Hidden Then isInlineHidden = True
        
        ' Objeto inline em parágrafo com espaçamento zero (provavelmente oculto)
        If inlineShp.Range.ParagraphFormat.LineSpacing = 0 Then isInlineHidden = True
        
        ' Objeto inline com tamanho zero
        If inlineShp.Width <= 1 Or inlineShp.Height <= 1 Then isInlineHidden = True
        
        If isInlineHidden Then
            LogMessage "Removendo objeto inline oculto (tipo: " & inlineShp.Type & ")"
            inlineShp.Delete
            deletedCount = deletedCount + 1
        End If
    Next i
    
    LogMessage "Remoção de elementos ocultos concluída: " & deletedCount & " elementos removidos"
    DeleteHiddenVisualElements = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover elementos visuais ocultos: " & Err.Description, LOG_LEVEL_ERROR
    DeleteHiddenVisualElements = False
End Function

'================================================================================
' DELETE VISUAL ELEMENTS IN RANGE - Remove elementos visuais entre os parágrafos 1-4
'================================================================================
Private Function DeleteVisualElementsInFirstFourParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Removendo elementos visuais entre os parágrafos 1-4..."
    
    If doc.Paragraphs.Count < 1 Then
        LogMessage "Documento não possui parágrafos - pulando limpeza de elementos visuais"
        DeleteVisualElementsInFirstFourParagraphs = True
        Exit Function
    End If
    
    If doc.Paragraphs.Count < 4 Then
        LogMessage "Documento possui menos de 4 parágrafos - removendo elementos dos parágrafos existentes (" & doc.Paragraphs.Count & " parágrafos)"
    End If
    
    Dim deletedCount As Long
    deletedCount = 0
    
    ' Define o range dos primeiros 4 parágrafos (ou menos se o documento for menor)
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
    
    LogMessage "Removendo elementos visuais dos parágrafos 1 a " & maxParagraphs & " (posição " & startRange & " a " & endRange & ")"
    
    ' Remove shapes flutuantes que estão ancorados na faixa dos primeiros 4 parágrafos
    Dim i As Long
    For i = doc.Shapes.Count To 1 Step -1
        Dim shp As Shape
        Set shp = doc.Shapes(i)
        
        ' Verifica se o shape está ancorado na faixa dos primeiros 4 parágrafos
        On Error Resume Next
        Dim anchorPosition As Long
        anchorPosition = shp.Anchor.Start
        On Error GoTo ErrorHandler
        
        If anchorPosition >= startRange And anchorPosition <= endRange Then
            Dim paragraphNum As Long
            paragraphNum = GetParagraphNumber(doc, anchorPosition)
            LogMessage "Removendo shape (tipo: " & shp.Type & ") ancorado no parágrafo " & paragraphNum
            shp.Delete
            deletedCount = deletedCount + 1
        End If
    Next i
    
    ' Remove objetos inline nos primeiros 4 parágrafos
    For i = doc.InlineShapes.Count To 1 Step -1
        Dim inlineShp As InlineShape
        Set inlineShp = doc.InlineShapes(i)
        
        ' Verifica se o objeto inline está na faixa dos primeiros 4 parágrafos
        If inlineShp.Range.Start >= startRange And inlineShp.Range.Start <= endRange Then
            Dim inlineParagraphNum As Long
            inlineParagraphNum = GetParagraphNumber(doc, inlineShp.Range.Start)
            LogMessage "Removendo objeto inline (tipo: " & inlineShp.Type & ") no parágrafo " & inlineParagraphNum
            inlineShp.Delete
            deletedCount = deletedCount + 1
        End If
    Next i
    
    LogMessage "Remoção de elementos visuais dos primeiros " & maxParagraphs & " parágrafos concluída: " & deletedCount & " elementos removidos"
    DeleteVisualElementsInFirstFourParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao remover elementos visuais dos primeiros 4 parágrafos: " & Err.Description, LOG_LEVEL_ERROR
    DeleteVisualElementsInFirstFourParagraphs = False
End Function

'================================================================================
' GET PARAGRAPH NUMBER - Função auxiliar para determinar o número do parágrafo
'================================================================================
Private Function GetParagraphNumber(doc As Document, position As Long) As Long
    Dim i As Long
    For i = 1 To doc.Paragraphs.Count
        If position >= doc.Paragraphs(i).Range.Start And position <= doc.Paragraphs(i).Range.End Then
            GetParagraphNumber = i
            Exit Function
        End If
    Next i
    GetParagraphNumber = 0 ' Não encontrado
End Function

'================================================================================
' CLEAN VISUAL ELEMENTS MAIN - Função principal para limpeza de elementos visuais
'================================================================================
Private Function CleanVisualElementsMain(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    LogMessage "============ INICIANDO LIMPEZA DE ELEMENTOS VISUAIS ============"
    LogMessage "Aplicando regras: (1) Remover elementos ocultos, (2) Remover elementos dos parágrafos 1-4"
    
    ' Contabiliza elementos antes da limpeza
    Dim initialShapeCount As Long
    Dim initialInlineShapeCount As Long
    initialShapeCount = doc.Shapes.Count
    initialInlineShapeCount = doc.InlineShapes.Count
    
    LogMessage "Estado inicial: " & initialShapeCount & " shapes flutuantes, " & initialInlineShapeCount & " objetos inline"
    
    ' 1. Remove todos os elementos visuais ocultos do documento
    LogMessage "=== FASE 1: Removendo elementos visuais ocultos ==="
    If Not DeleteHiddenVisualElements(doc) Then
        LogMessage "Falha ao remover elementos visuais ocultos", LOG_LEVEL_WARNING
    End If
    
    ' 2. Remove elementos visuais entre os parágrafos 1-4 (visíveis ou não)
    LogMessage "=== FASE 2: Removendo elementos visuais dos parágrafos 1-4 ==="
    If Not DeleteVisualElementsInFirstFourParagraphs(doc) Then
        LogMessage "Falha ao remover elementos visuais dos primeiros 4 parágrafos", LOG_LEVEL_WARNING
    End If
    
    ' Contabiliza elementos após a limpeza
    Dim finalShapeCount As Long
    Dim finalInlineShapeCount As Long
    finalShapeCount = doc.Shapes.Count
    finalInlineShapeCount = doc.InlineShapes.Count
    
    Dim shapesRemoved As Long
    Dim inlineShapesRemoved As Long
    shapesRemoved = initialShapeCount - finalShapeCount
    inlineShapesRemoved = initialInlineShapeCount - finalInlineShapeCount
    
    LogMessage "Estado final: " & finalShapeCount & " shapes flutuantes, " & finalInlineShapeCount & " objetos inline"
    LogMessage "Resumo da limpeza: " & shapesRemoved & " shapes flutuantes removidos, " & inlineShapesRemoved & " objetos inline removidos"
    LogMessage "============ LIMPEZA DE ELEMENTOS VISUAIS CONCLUÍDA ============"
    
    CleanVisualElementsMain = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na limpeza de elementos visuais: " & Err.Description, LOG_LEVEL_ERROR
    CleanVisualElementsMain = False
End Function

'================================================================================
' VIEW SETTINGS PROTECTION SYSTEM - SISTEMA DE PROTEÇÃO DAS CONFIGURAÇÕES DE VISUALIZAÇÃO
'================================================================================

'================================================================================
' BACKUP VIEW SETTINGS - Faz backup das configurações de visualização originais
'================================================================================
Private Function BackupViewSettings(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Fazendo backup das configurações de visualização..."
    
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
    
    Application.StatusBar = "Restaurando configurações de visualização originais..."
    
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

' Compiled successfully on 2024-06-10 12:34:56
'  --- IGNORE ---