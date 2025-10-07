Attribute VB_Name = "modPipeline"
'================================================================================
' MODULE: modPipeline (formerly modMain)
' PURPOSE: Final canonical orchestrator for the Chainsaw pipeline. All domain
'          logic (formatting, replacements, validation, safety, config, UI,
'          constants, messages) lives in their dedicated modules. This module
'          ONLY sequences calls, manages application state, and handles any
'          high-level fail‑soft error trapping.
'================================================================================
Option Explicit

' Public entrypoint (stable)
Public Function RunChainsawPipeline() As Boolean
    Dim doc As Document
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As WdAlertLevel
    Dim hadError As Boolean
    Dim cfgLoaded As Boolean

    hadError = False
    RunChainsawPipeline = False
    On Error GoTo FatalPipelineError

    ' Acquire active document safely
    Set doc = Nothing
    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox NormalizeForUI("Nenhum documento ativo encontrado."), vbExclamation, NormalizeForUI("Chainsaw - Documento ausente")
        GoTo Finalize
    End If

    ' Load configuration (idempotent)
    Call modConfig_LoadConfigIfNeeded(cfgLoaded)

    ' Runtime environment hardening (respect configuration flags)
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    If Config.disableScreenUpdating Then Application.ScreenUpdating = False
    If Config.disableDisplayAlerts Then Application.DisplayAlerts = wdAlertsNone

    ' Preliminary validation stage
    If Not Pipeline_PreviousChecking(doc) Then GoTo Finalize

    ' Formatting & replacement stage
    If Not Pipeline_PreviousFormatting(doc) Then GoTo Finalize

    RunChainsawPipeline = True
    GoTo Finalize

FatalPipelineError:
    hadError = True
    RunChainsawPipeline = False

Finalize:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    If hadError Then
        Application.StatusBar = "Chainsaw: erro inesperado durante o processamento"
    ElseIf RunChainsawPipeline Then
        If Config.showStatusBarUpdates Then Application.StatusBar = "Chainsaw: processamento concluído"
    Else
        If Config.showStatusBarUpdates Then Application.StatusBar = "Chainsaw: processamento interrompido"
    End If
End Function

'================================================================================
' INTERNAL: PREVIOUS CHECKING (migrated from legacy modMain)
'================================================================================
Private Function Pipeline_PreviousChecking(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        Application.StatusBar = "Error: Document not accessible for verification"
        Pipeline_PreviousChecking = False: Exit Function
    End If
    If doc.Type <> wdTypeDocument Then
        Application.StatusBar = "Error: Unsupported document type (Type: " & doc.Type & ")"
        Pipeline_PreviousChecking = False: Exit Function
    End If
    If doc.protectionType <> wdNoProtection Then
        Application.StatusBar = "Error: Document is protected (" & Pipeline_GetProtectionType(doc) & ")"
        Pipeline_PreviousChecking = False: Exit Function
    End If
    If doc.ReadOnly Then
        Application.StatusBar = "Error: Document is read-only"
        Pipeline_PreviousChecking = False: Exit Function
    End If
    If Not Pipeline_CheckDiskSpace(doc) Then
        Application.StatusBar = "Error: Not enough disk space"
        Pipeline_PreviousChecking = False: Exit Function
    End If
    ' Structural validation placeholder (currently benign)
    If Not ValidateDocumentStructure(doc) Then
        ' Intentionally ignoring result; legacy hook
    End If
    Pipeline_PreviousChecking = True
    Exit Function
ErrorHandler:
    Application.StatusBar = "Error during security checks"
    Pipeline_PreviousChecking = False
End Function

'================================================================================
' INTERNAL: DISK SPACE CHECK (simplified, fail-open)
'================================================================================
Private Function Pipeline_CheckDiskSpace(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim fso As Object, drive As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If doc.Path <> "" Then
        Set drive = fso.GetDrive(Left(doc.Path, 3))
    Else
        Set drive = fso.GetDrive(Left(Environ("TEMP"), 3))
    End If
    If drive.AvailableSpace < 10485760 Then ' < 10MB
        Pipeline_CheckDiskSpace = False
    Else
        Pipeline_CheckDiskSpace = True
    End If
    Exit Function
ErrorHandler:
    Pipeline_CheckDiskSpace = True ' fail-open
End Function

'================================================================================
' INTERNAL: PROTECTION TYPE TEXT
'================================================================================
Private Function Pipeline_GetProtectionType(doc As Document) As String
    On Error Resume Next
    Select Case doc.protectionType
        Case wdNoProtection: Pipeline_GetProtectionType = "No protection"
        Case 1: Pipeline_GetProtectionType = "Tracked changes protection"
        Case 2: Pipeline_GetProtectionType = "Comments protection"
        Case 3: Pipeline_GetProtectionType = "Forms protection"
        Case 4: Pipeline_GetProtectionType = "Read-only protection"
        Case Else: Pipeline_GetProtectionType = "Unknown type (" & doc.protectionType & ")"
    End Select
End Function

'================================================================================
' INTERNAL: MAIN FORMATTING & REPLACEMENT SEQUENCE (migrated)
'================================================================================
Private Function Pipeline_PreviousFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ' Page setup
    If Not ApplyPageSetup(doc) Then GoTo FailSoft
    ' Structural cleanup
    If Not CleanDocumentStructure(doc) Then GoTo FailSoft
    ' Proposition type (informational)
    Call ValidatePropositionType(doc)
    ' Content consistency (may abort)
    If Not modValidation.ValidateContentConsistency(doc) Then GoTo Fail
    ' Title formatting
    Call FormatDocumentTitle(doc)
    ' Standard font
    If Not ApplyStdFont(doc) Then GoTo FailSoft
    ' Standard paragraphs
    If Not ApplyStdParagraphs(doc) Then GoTo FailSoft
    ' First / Second paragraphs
    Call FormatFirstParagraph(doc)
    Call FormatSecondParagraph(doc)
    ' CONSIDERANDO emphasis
    If Not FormatConsiderandoParagraphs(doc) Then GoTo FailSoft
    ' Text replacements (generic + specific)
    If Not modReplacements.ApplyTextReplacements(doc) Then GoTo FailSoft
    If Not modReplacements.ApplySpecificParagraphReplacements(doc) Then GoTo FailSoft
    ' Numbered paragraphs normalization
    Call FormatNumberedParagraphs(doc)
    ' Justificativa / Anexo
    Call FormatJustificativaAnexoParagraphs(doc)
    ' Hyphenation & watermark
    Call EnableHyphenation(doc)
    Call RemoveWatermark(doc)
    ' Header image & footer page numbers
    InsertHeaderstamp doc
    If Not InsertFooterstamp(doc) Then GoTo FailSoft
    ' Final spacing / normalization passes
    Call CleanMultipleSpaces(doc)
    Call LimitSequentialEmptyLines(doc)
    Call EnsureParagraphSeparation(doc)
    Call EnsureSecondParagraphBlankLines(doc)
    Call FormatJustificativaAnexoParagraphs(doc) ' reinforce after spacing
    ' View configuration
    Call ConfigureDocumentView(doc)

    Pipeline_PreviousFormatting = True
    Exit Function

FailSoft:
    ' Non-fatal formatting issue → continue but mark as success
    Pipeline_PreviousFormatting = True
    Exit Function
Fail:
    Pipeline_PreviousFormatting = False
    Exit Function
ErrorHandler:
    Pipeline_PreviousFormatting = False
End Function

'================================================================================
' PLACEHOLDER: STRUCTURE VALIDATION (legacy hook, always true for now)
'================================================================================
Private Function ValidateDocumentStructure(doc As Document) As Boolean
    On Error GoTo ErrHandler
    ' Lightweight structural sanity checks:
    ' 1. Non-empty document.
    ' 2. Reasonable paragraph count (not absurdly high for typical proposition).
    ' 3. First non-empty paragraph length threshold.

    Dim para As Paragraph, i As Long, firstText As String
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        firstText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        If firstText <> "" Then Exit For
        If i > 200 Then Exit For ' safety cap
    Next i

    If firstText = "" Then
        MsgBox NormalizeForUI(MSG_EMPTY_DOC), vbExclamation, NormalizeForUI(TITLE_VALIDATION_ERROR)
        ValidateDocumentStructure = False: Exit Function
    End If

    If doc.Paragraphs.Count > 5000 Then
        Dim largeWarn As String
        largeWarn = ReplacePlaceholders(MSG_PARAGRAPH_EXCESS, "COUNT", CStr(doc.Paragraphs.Count))
        If MsgBox(NormalizeForUI(largeWarn), vbYesNo + vbQuestion + vbDefaultButton2, NormalizeForUI(TITLE_LARGE_DOC)) = vbNo Then
            ValidateDocumentStructure = False: Exit Function
        End If
    End If

    If Len(firstText) < 3 Then
        If MsgBox(NormalizeForUI(MSG_FIRST_PARA_SHORT), vbYesNo + vbExclamation + vbDefaultButton2, NormalizeForUI(TITLE_VALIDATION_ERROR)) = vbNo Then
            ValidateDocumentStructure = False: Exit Function
        End If
    End If

    ValidateDocumentStructure = True
    Exit Function
ErrHandler:
    ValidateDocumentStructure = True ' fail-open to avoid blocking formatting
End Function

' End of file.