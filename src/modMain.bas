' =============================================================================
' PROJECT: CHAINSAW PROPOSITURAS
' =============================================================================
' LEGACY TRANSITION MODULE (modMain.bas)
' -----------------------------------------------------------------------------
' This file still contains portions of the original monolithic implementation.
' Refactor Status:
'   * New pipeline entry: RunChainsawPipeline (public) – STABLE
'   * Formatting logic: progressively migrating to modFormatting.bas
'   * Text replacement logic: migrated to modReplacements.bas
'   * Validation routines: being isolated in modValidation.bas
'   * Safety wrappers: moved to modSafety.bas
'   * Configuration loading: modConfig.bas
'   * Logging: stubbed in modLog.bas (no file I/O in beta)
'   * Backup & view/image protection: marked LEGACY; slated for isolation or
'     removal in a future beta unless formally re‑scoped.
'
' Contributor Guidance:
'   - Do NOT add new business logic here; place it in the appropriate module.
'   - When migrating a cohesive block, move the code, update calls, then delete
'     the legacy block to avoid divergence.
'   - Keep behavior identical (formatting semantics MUST NOT change).
'   - Remove obsolete comments & dead code as you extract.
'
' Deletion Criteria for This File:
'   When only thin orchestration helpers remain, rename/merge or remove this
'   module and let the entry stub (chainsaw.bas) call directly into the final
'   orchestrator module (e.g., modPipeline.bas if created).
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
'(Removed keybd_event and virtual key constants – obsolete after ChatGPT feature removal)

'================================================================================
' PUBLIC ENTRYPOINT WRAPPER
' Provides a stable callable pipeline while refactor migration continues.
'================================================================================
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

    ' Load configuration once per session
    Call modConfig_LoadConfigIfNeeded(cfgLoaded)

    ' Runtime environment hardening
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    If Config.disableScreenUpdating Then Application.ScreenUpdating = False
    If Config.disableDisplayAlerts Then Application.DisplayAlerts = wdAlertsNone

    ' Preliminary validation stage
    If Not PreviousChecking(doc) Then GoTo Finalize

    ' Formatting & replacement stage
    If Not PreviousFormatting(doc) Then GoTo Finalize

    RunChainsawPipeline = True
    GoTo Finalize

FatalPipelineError:
    hadError = True
    ' Swallow unexpected errors – stability priority. (Optional: surface message)
    RunChainsawPipeline = False

Finalize:
    On Error Resume Next
    ' Restore UI state deterministically
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

Private Function ApplyTextReplacements(doc As Document) As Boolean: ApplyTextReplacements = modReplacements.ApplyTextReplacements(doc): End Function
    
        Application.ScreenUpdating = True
        Application.DisplayAlerts = wdAlertsAll
        Application.StatusBar = ""
        Application.EnableCancelKey = wdCancelInterrupt
    
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
    End If
    
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
            Case vbYes: doc.Save
            Case vbCancel: Exit Function
            Case vbNo: ' continue without saving
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
          ' Removed stray duplicated MsgBox (already handled in main flow)
    
End Function

Private Function SafeSetFont(targetRange As Range, fontName As String, fontSize As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' Apply font formatting safely
    With targetRange.Font
        If fontName <> "" Then .Name = fontName
        If fontSize > 0 Then .size = fontSize
        .Color = wdColorAutomatic
    End With
    
                ' Removed stray duplicated version requirement MsgBox (handled earlier)
ErrorHandler:
    SafeSetFont = False
    
End Function

Private Function SafeSetParagraphFormat(para As Paragraph, alignment As Long, leftIndent As Single, firstLineIndent As Single) As Boolean
    On Error GoTo ErrorHandler
    
    With para.Format
        If alignment >= 0 Then .alignment = alignment
        If leftIndent >= 0 Then .leftIndent = leftIndent
        If firstLineIndent >= 0 Then .firstLineIndent = firstLineIndent
    End With
            MsgBox NormalizeForUI("No document is open or accessible." & vbCrLf & _
                 "Open a document before running the standardization."), vbExclamation, NormalizeForUI("Document Not Found - Chainsaw Proposituras")
    Exit Function
    
ErrorHandler:
    SafeSetParagraphFormat = False
    
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
    ' Simpler alternative method
    SafeHasVisualContent = (para.Range.InlineShapes.count > 0)
    Exit Function
    
ErrorHandler:
    ' In case of error, assume no visual content
    SafeHasVisualContent = False
End Function

'================================================================================
' SAFE FIND/REPLACE OPERATIONS - Compatibility with all versions
'================================================================================
Private Function SafeFindReplace(doc As Document, findText As String, replaceText As String, Optional useWildcards As Boolean = False) As Long
    On Error GoTo ErrorHandler
    
    Dim findCount As Long
    findCount = 0
    
    ' Safe Find/Replace configuration
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
    .MatchWildcards = useWildcards  ' Controlled parameter
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
    ' Execute replacement and count occurrences
        Do While .Execute(Replace:=True)
            findCount = findCount + 1
            ' Safety limit to avoid infinite loops
            If findCount > 10000 Then
                
                Exit Do
            End If
        Loop
    End With
    
    SafeFindReplace = findCount
    Exit Function
    
ErrorHandler:
    SafeFindReplace = 0
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

"' Logging removed"

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
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If doc.Path <> "" Then
        Set drive = fso.GetDrive(Left(doc.Path, 3))
    Else
        Set drive = fso.GetDrive(Left(Environ("TEMP"), 3))
    End If
    
    ' Basic verification - 10MB minimum
    If drive.AvailableSpace < 10485760 Then ' 10MB in bytes
    
        CheckDiskSpace = False
    Else
        CheckDiskSpace = True
    End If
    
    Exit Function
    
ErrorHandler:
    ' If cannot verify, assume there is sufficient space
    CheckDiskSpace = True
End Function

'================================================================================
' MAIN FORMATTING ROUTINE
'================================================================================
Private Function PreviousFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    
    
    ' Apply page setup (always on)
    LogStepStart "Page setup"
    If Not ApplyPageSetup(doc) Then
    
        LogStepEnd False
    Else
        LogStepEnd True
    End If
    
    ' Clean document structure (remove blank lines above the first text and leading spaces)
    LogStepStart "Clean document structure"
    If Not CleanDocumentStructure(doc) Then
    
        LogStepEnd False
    Else
        LogStepEnd True
    End If

    ' Validate proposition type (informational, can be skipped by user)
    LogStepStart "Validate proposition type"
    Call ValidatePropositionType(doc)
    LogStepEnd True

    ' Validate content consistency (may cancel if user chooses so)
    LogStepStart "Validate content consistency"
    If Not modValidation.ValidateContentConsistency(doc) Then
    
        LogStepEnd False
        PreviousFormatting = False
        Exit Function
    Else
        LogStepEnd True
    End If
    
    ' Title formatting (uppercased, bold, underlined, centered)
    LogStepStart "Format document title"
    Call FormatDocumentTitle(doc)
    LogStepEnd True

    ' Apply standard font (always on)
    LogStepStart "Apply standard font"
    If Not ApplyStdFont(doc) Then
    
        LogStepEnd False
    Else
        LogStepEnd True
    End If
    
    ' Apply standard paragraphs (always on)
    LogStepStart "Apply standard paragraphs"
    If Not ApplyStdParagraphs(doc) Then
    
        LogStepEnd False
    Else
        LogStepEnd True
    End If

    ' Format first and second paragraphs
    LogStepStart "Format first paragraph"
    Call FormatFirstParagraph(doc)
    LogStepEnd True
    LogStepStart "Format second paragraph"
    Call FormatSecondParagraph(doc)
    LogStepEnd True
    
    ' Apply CONSIDERANDO uppercase/bold at paragraph start
    LogStepStart "Format 'CONSIDERANDO' paragraphs"
    If Not FormatConsiderandoParagraphs(doc) Then
    
        LogStepEnd False
    Else
        LogStepEnd True
    End If
    
    ' Apply text replacements (always on)
    LogStepStart "Apply text replacements"
    If Not modReplacements.ApplyTextReplacements(doc) Then
    
        LogStepEnd False
    Else
        LogStepEnd True
    End If
    
    ' Apply specific paragraph replacements (always on)
    LogStepStart "Apply specific paragraph replacements"
    If Not modReplacements.ApplySpecificParagraphReplacements(doc) Then
    
        LogStepEnd False
    Else
        LogStepEnd True
    End If
    
    ' Normalize numbered paragraphs
    LogStepStart "Format numbered paragraphs"
    Call FormatNumberedParagraphs(doc)
    LogStepEnd True
    
    ' Justificativa/Anexo formatting
    LogStepStart "Format 'Justificativa/Anexo' paragraphs"
    Call FormatJustificativaAnexoParagraphs(doc)
    LogStepEnd True

    ' Hyphenation and watermark
    LogStepStart "Enable hyphenation"
    Call EnableHyphenation(doc)
    LogStepEnd True
    LogStepStart "Remove watermark"
    Call RemoveWatermark(doc)
    LogStepEnd True

    ' Insert header image (always enabled)
    LogStepStart "Insert header image"
    InsertHeaderstamp doc
    LogStepEnd True
    
    ' Insert page numbers in footer (restored feature)
    LogStepStart "Insert footer page numbers"
    If Not InsertFooterstamp(doc) Then
    
        LogStepEnd False
    Else
        LogStepEnd True
    End If
    
    ' Final spacing and separation controls
    LogStepStart "Clean multiple spaces"
    Call CleanMultipleSpaces(doc)
    LogStepEnd True
    LogStepStart "Limit sequential empty lines"
    Call LimitSequentialEmptyLines(doc)
    LogStepEnd True
    LogStepStart "Ensure paragraph separation"
    Call EnsureParagraphSeparation(doc)
    LogStepEnd True
    LogStepStart "Reinforce 2nd paragraph blank lines"
    Call EnsureSecondParagraphBlankLines(doc)
    LogStepEnd True
    LogStepStart "Reapply 'Justificativa/Anexo' formatting"
    Call FormatJustificativaAnexoParagraphs(doc)
    LogStepEnd True
    
    ' Configure view (keeps user zoom)
    LogStepStart "Configure document view"
    Call ConfigureDocumentView(doc)
    LogStepEnd True
    
    ' Clipboard pane visibility enforcement removed per request
    
    PreviousFormatting = True
    
    Exit Function

ErrorHandler:
    
    PreviousFormatting = False
End Function

 ' (Formatting functions ApplyPageSetup, ApplyStdFont, FormatCharacterByCharacter moved to modFormatting.bas)

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
        
    ' Safety limit
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
    
    ' Check following paragraphs (maximum 5 for performance)
    For i = paraIndex + 1 To doc.Paragraphs.count
        If i > doc.Paragraphs.count Then Exit For
        
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
    ' If the paragraph is empty, count as a blank line
        If paraText = "" And Not HasVisualContent(para) Then
            count = count + 1
        Else
            ' If a paragraph with content is found, stop counting
            Exit For
        End If
        
    ' Safety limit
        If count >= 5 Then Exit For
    Next i
    
    CountBlankLinesAfter = count
    Exit Function
    
ErrorHandler:
    CountBlankLinesAfter = 0
End Function

'================================================================================
' SECOND PARAGRAPH LOCATION HELPER - Locate the second paragraph
'================================================================================
Private Function GetSecondParagraphIndex(doc As Document) As Long
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    
    actualParaIndex = 0
    
    ' Find the 2nd paragraph with content (skip empty)
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
    ' If the paragraph has text or visual content, count as a valid paragraph
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1
            
            ' Return the index of the 2nd paragraph
            If actualParaIndex = 2 Then
                GetSecondParagraphIndex = i
                Exit Function
            End If
        End If
        
    ' Protection: process up to 50 paragraphs to find the 2nd
        If i > 50 Then Exit For
    Next i
    
    GetSecondParagraphIndex = 0  ' Not found
    Exit Function
    
ErrorHandler:
    GetSecondParagraphIndex = 0
End Function

'================================================================================
' ENSURE SECOND PARAGRAPH BLANK LINES - Ensure two blank lines
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
        
    ' Check and fix blank lines BEFORE
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
            
            ' Update the index (it shifted)
            secondParaIndex = secondParaIndex + linesToAdd
            Set para = doc.Paragraphs(secondParaIndex)
        End If
        
    ' Check and fix blank lines AFTER
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
        
    
    End If
    
    EnsureSecondParagraphBlankLines = True
    Exit Function
    
ErrorHandler:
    EnsureSecondParagraphBlankLines = False
    
End Function
     ' (Removed duplicate EnsureSecondParagraphBlankLines – migrated to modFormatting.EnsureSecondParagraphBlankLines)

'================================================================================
' FORMAT FIRST PARAGRAPH
'================================================================================
Private Function FormatFirstParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim firstParaIndex As Long
    
    ' Identify the 1st paragraph (considering only paragraphs with text)
    actualParaIndex = 0
    firstParaIndex = 0
    
    ' Find the 1st paragraph with content (skip empty)
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
    ' If the paragraph has text or visual content, count as a valid paragraph
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1
            
            ' Record the index of the 1st paragraph
            If actualParaIndex = 1 Then
                firstParaIndex = i
                Exit For ' We have found the 1st paragraph
            End If
        End If
        
        ' Expanded protection: process up to 20 paragraphs to find the 1st
        If i > 20 Then Exit For
    Next i
    
    ' Apply specific formatting only to the 1st paragraph
    If firstParaIndex > 0 And firstParaIndex <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(firstParaIndex)
        
    ' NEW: Always apply formatting, protecting only images
    ' 1st paragraph formatting: uppercase, bold and underlined
        If HasVisualContent(para) Then
            ' For paragraphs with images, apply character-by-character formatting
            Dim n As Long
            Dim charCount4 As Long
            charCount4 = SafeGetCharacterCount(para.Range) ' Cache da contagem segura
            
            If charCount4 > 0 Then ' Safety check
                For n = 1 To charCount4
                    Dim charRange3 As Range
                    Set charRange3 = para.Range.Characters(n)
                    If charRange3.InlineShapes.count = 0 Then
                        With charRange3.Font
                            .AllCaps = True           ' Uppercase
                            .Bold = True              ' Bold
                            .Underline = wdUnderlineSingle ' Underlined
                        End With
                    End If
                Next n
            End If
            
        Else
            ' Normal formatting for paragraphs without images
            With para.Range.Font
                .AllCaps = True           ' Uppercase
                .Bold = True              ' Bold
                .Underline = wdUnderlineSingle ' Underlined
            End With
        End If
        
    ' Also apply paragraph formatting - ALWAYS
        With para.Format
            .alignment = wdAlignParagraphCenter       ' Centered
            .leftIndent = 0                           ' No left indent
            .firstLineIndent = 0                      ' No first-line indent
            .RightIndent = 0                          ' No right indent
        End With
    Else
    
    End If
    
    FormatFirstParagraph = True
    Exit Function

ErrorHandler:
    
    FormatFirstParagraph = False
End Function
     ' (Removed duplicate FormatFirstParagraph – migrated to modFormatting.FormatFirstParagraph)

'================================================================================
' ENABLE HYPHENATION
'================================================================================
Private Function EnableHyphenation(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    If Not doc.AutoHyphenation Then
        doc.AutoHyphenation = True
        doc.HyphenationZone = CentimetersToPoints(0.63)
        doc.HyphenateCaps = True
    ' Log removed for performance
        EnableHyphenation = True
    Else
    ' Log removed for performance
        EnableHyphenation = True
    End If
    
    Exit Function
    
ErrorHandler:
    
    EnableHyphenation = False
End Function
     ' (Removed duplicate EnableHyphenation – migrated to modFormatting.EnableHyphenation)

'================================================================================
' REMOVE WATERMARK
'================================================================================
Private Function RemoveWatermark(doc As Document) As Boolean
 ' (Removed duplicate RemoveWatermark – migrated to modFormatting.RemoveWatermark)
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
    
    End If
    ' "No watermark" log removed for performance
    
    RemoveWatermark = True
    Exit Function
    
ErrorHandler:
    
    RemoveWatermark = False
End Function

'================================================================================
' INSERT HEADER IMAGE
'================================================================================
Private Function InsertHeaderstamp(doc As Document) As Boolean
 ' (Removed duplicate InsertHeaderstamp – migrated to modFormatting.InsertHeaderstamp)
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
    imgFile = Trim(Config.headerImagePath)
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
    
        InsertHeaderstamp = False
    End If

    Exit Function

ErrorHandler:
    
    InsertHeaderstamp = False
End Function

'================================================================================
' INSERT FOOTER PAGE NUMBERS
'================================================================================
Private Function InsertFooterstamp(doc As Document) As Boolean
 ' (Removed duplicate InsertFooterstamp – migrated to modFormatting.InsertFooterstamp)
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

    ' Detailed log removed for performance
    InsertFooterstamp = True
    Exit Function

ErrorHandler:
    
    InsertFooterstamp = False
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
     ' (Removed duplicate CentimetersToPoints – migrated to modFormatting.CentimetersToPoints)

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
    rawName = "UnknownUser"
    End If
    
    For i = 1 To Len(rawName)
        c = Mid(rawName, i, 1)
        If c Like "[A-Za-z0-9_\-]" Then
            safeName = safeName & c
        ElseIf c = " " Then
            safeName = safeName & "_"
        End If
    Next i
    
    If safeName = "" Then safeName = "User"
    
    GetSafeUserName = safeName
    Exit Function
    
ErrorHandler:
    GetSafeUserName = "User"
End Function

'================================================================================
' VALIDATE DOCUMENT STRUCTURE
'================================================================================
Private Function ValidateDocumentStructure(doc As Document) As Boolean
    On Error Resume Next
    
    ' Basic and fast verification
    If doc.Range.End > 0 And doc.Sections.count > 0 Then
        ValidateDocumentStructure = True
    Else
    
         ' (Removed duplicate GetSafeUserName – migrated to modFormatting.GetSafeUserName)
        ValidateDocumentStructure = False
    End If
End Function

'================================================================================
' CRITICAL FIX: SAVE DOCUMENT BEFORE PROCESSING
' TO PREVENT CRASHES ON NEW NON SAVED DOCUMENTS
'================================================================================
Private Function SaveDocumentFirst(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Waiting for document to be saved..."
    ' Start log removed for performance
    
    Dim saveDialog As Object
    Set saveDialog = Application.Dialogs(wdDialogFileSaveAs)

    If saveDialog.Show <> -1 Then
    
    Application.StatusBar = "Save cancelled by user"
        SaveDocumentFirst = False
        Exit Function
    End If

    ' Wait for save confirmation with safety timeout
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
    Application.StatusBar = "Waiting for save... (" & waitCount & "/" & maxWait & ")"
    Next waitCount

    If doc.Path = "" Then
    
    Application.StatusBar = "Save failed - operation cancelled"
        SaveDocumentFirst = False
    Else
    ' Success log removed for performance
    Application.StatusBar = "Document saved successfully"
        SaveDocumentFirst = True
    End If

    Exit Function

ErrorHandler:
    
    Application.StatusBar = "Error during save"
    SaveDocumentFirst = False
End Function

'================================================================================
' CLEAR ALL FORMATTING - INITIAL FULL CLEANUP
'================================================================================
 ' Removed: ClearAllFormatting

'================================================================================
' CLEAN DOCUMENT STRUCTURE - FEATURES 2, 6, 7
'================================================================================
Private Function CleanDocumentStructure(doc As Document) As Boolean
    CleanDocumentStructure = modFormatting.CleanDocumentStructure(doc)
End Function

'================================================================================
' SAFE CHECK FOR VISUAL CONTENT
'================================================================================
Private Function HasVisualContent(para As Paragraph) As Boolean: HasVisualContent = modFormatting.HasVisualContent(para): End Function

'================================================================================
' VALIDATE PROPOSITION TYPE
'================================================================================
Private Function ValidatePropositionType(doc As Document) As Boolean: ValidatePropositionType = modValidation.ValidatePropositionType(doc): End Function

'================================================================================
' FORMAT DOCUMENT TITLE
'================================================================================
Private Function FormatDocumentTitle(doc As Document) As Boolean: FormatDocumentTitle = modFormatting.FormatDocumentTitle(doc): End Function

'================================================================================
' FORMAT "CONSIDERANDO" PARAGRAPHS - OPTIMIZED AND SIMPLIFIED
'================================================================================
Private Function FormatConsiderandoParagraphs(doc As Document) As Boolean: FormatConsiderandoParagraphs = modFormatting.FormatConsiderandoParagraphs(doc): End Function

''================================================================================
'' MODULE: chainsaw.bas
'' PURPOSE: Standardize Word proposition documents (font, paragraphs, title, numbering,
''          replacements, structural cleanup, validations, backups, visual element handling).
'' VERSION: v1.0.0-Beta1 (baseline after logging removal and aggressive cleanup)
'' DATE: 2025-10-06
'' NOTES:
''  - Logging system fully removed.
''  - Comment noise aggressively minimized.
''  - Preserve formatting logic exactly as validated earlier.
''  - Safe helper functions ensure Word 2010+ compatibility.
''================================================================================


'================================================================================
' UI STRING NORMALIZATION - produce ASCII-safe text for MsgBox dialogs
'================================================================================
Private Function NormalizeForUI(ByVal s As String) As String
    On Error Resume Next
    If Not dialogAsciiNormalizationEnabled Then
        NormalizeForUI = s
        Exit Function
    End If
    Dim i As Long, ch As String, code As Long
    Dim out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case code
            Case 192 To 197, 224 To 229: out = out & "a"   ' ÀÁÂÃÄÅàáâãäå
            Case 199: out = out & "C"                      ' Ç
            Case 231: out = out & "c"                      ' ç
            Case 200 To 203, 232 To 235: out = out & "e"
            Case 204 To 207, 236 To 239: out = out & "i"
            Case 210 To 214, 242 To 246: out = out & "o"
            Case 217 To 220, 249 To 252: out = out & "u"
            Case 209: out = out & "N"                      ' Ñ
            Case 241: out = out & "n"                      ' ñ
            Case 8211, 8212: out = out & "-"               ' en/em dash
            Case 8216, 8217: out = out & "'"               ' curly apostrophes
            Case 8220, 8221, 171, 187: out = out & """"    ' various quotes -> standard quote
            Case 10, 13: out = out & ch                    ' CR/LF
            Case 32 To 126: out = out & ch                 ' standard ASCII
            Case Else: out = out & "?"
        End Select
    Next i
    NormalizeForUI = out
End Function

'--------------------------------------------------------------------------------
' ReplacePlaceholders - convenience wrapper for common {{KEY}} replacements
' Example: ReplacePlaceholders(MSG_ERR_VERSION, "MIN", 14, "CUR", Application.Version)
'--------------------------------------------------------------------------------
Private Function ReplacePlaceholders(ByVal template As String, ParamArray kv()) As String
    On Error Resume Next
    Dim i As Long, result As String, k As String, val As String
    result = template
    For i = LBound(kv) To UBound(kv) Step 2
        If i + 1 <= UBound(kv) Then
            k = CStr(kv(i))
            val = CStr(kv(i + 1))
            result = Replace(result, "{{" & k & "}}", val)
        End If
    Next i
    ReplacePlaceholders = result
End Function

'================================================================================
' APPLY TEXT REPLACEMENTS
'================================================================================
Private Function ApplyTextReplacements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Dim replacementCount As Long
    
    Set rng = doc.Range
    
    ' Feature 10: Replace variants of "d'Oeste"
    Dim dOesteVariants() As String
    Dim i As Long
    
    ' Define possible variants of the first 3 characters of "d'Oeste"
    ReDim dOesteVariants(0 To 15)
    dOesteVariants(0) = "d'O"   ' Original
    dOesteVariants(1) = "d´O"   ' Acute accent
    dOesteVariants(2) = "d`O"   ' Grave accent
    dOesteVariants(3) = "d" & Chr(8220) & "O"   ' Left curly quote
    dOesteVariants(4) = "d'o"   ' Lowercase
    dOesteVariants(5) = "d´o"
    dOesteVariants(6) = "d`o"
    dOesteVariants(7) = "d" & Chr(8220) & "o"
    dOesteVariants(8) = "D'O"   ' Uppercase D
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
            .Replacement.ClearFormatting
            .text = dOesteVariants(i) & "este"
            .Replacement.text = "d'Oeste"
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            Do While .Execute(Replace:=wdReplaceOne)
                replacementCount = replacementCount + 1
                rng.Collapse wdCollapseEnd
            Loop
        End With
    Next i
    
    ' Removed: vereador variants replacement
    
    ' Feature 12: Replace isolated hyphens/en dashes with em dash (—)
    ' Normalizes hyphens (-) and en dashes (–) surrounded by spaces into em dashes (—)
    Set rng = doc.Range
    Dim dashVariants() As String
    ReDim dashVariants(0 To 2)
    
    ' Define dash types to replace when surrounded by spaces
    dashVariants(0) = " - "     ' Hyphen
    dashVariants(1) = " – "     ' En dash
    dashVariants(2) = " — "     ' Em dash (normalize)
    
    ' Replace all types with em dash
    For i = 0 To UBound(dashVariants)
    ' Only if not already an em dash
        If dashVariants(i) <> " — " Then
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .text = dashVariants(i)
                .Replacement.text = " — "    ' Em dash (travessão) com espaços
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                
                Do While .Execute(Replace:=wdReplaceOne)
                    replacementCount = replacementCount + 1
                    rng.Collapse wdCollapseEnd
                Loop
            End With
        End If
    Next i
    
    ' Special cases: hyphen/en dash at the start of the line followed by a space
    Set rng = doc.Range
    Dim lineStartDashVariants() As String
    ReDim lineStartDashVariants(0 To 1)
    
    lineStartDashVariants(0) = "^p- "   ' Hyphen at line start
    lineStartDashVariants(1) = "^p– "   ' En dash at line start
    
    For i = 0 To UBound(lineStartDashVariants)
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = lineStartDashVariants(i)
            .Replacement.text = "^p— "    ' Em dash at line start
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            Do While .Execute(Replace:=wdReplaceOne)
                replacementCount = replacementCount + 1
                rng.Collapse wdCollapseEnd
            Loop
        End With
    Next i
    
    ' Special cases: space then hyphen/en dash at the end of the line
    Set rng = doc.Range
    Dim lineEndDashVariants() As String
    ReDim lineEndDashVariants(0 To 1)
    
    lineEndDashVariants(0) = " -^p"   ' Hyphen at line end
    lineEndDashVariants(1) = " –^p"   ' En dash at line end
    
    For i = 0 To UBound(lineEndDashVariants)
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = lineEndDashVariants(i)
            .Replacement.text = " —^p"    ' Em dash at line end
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            Do While .Execute(Replace:=wdReplaceOne)
                replacementCount = replacementCount + 1
                rng.Collapse wdCollapseEnd
            Loop
        End With
    Next i
    
    ' Feature 13: Remove all manual line breaks (soft breaks); keep regular paragraph breaks
    Set rng = doc.Range
    
    ' Remove manual line breaks (Shift+Enter) - Chr(11) or ^l
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "^l"  ' Manual line break
        .Replacement.text = " "  ' Replace with space
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Do While .Execute(Replace:=wdReplaceOne)
            replacementCount = replacementCount + 1
            rng.Collapse wdCollapseEnd
        Loop
    End With
    
    ' Remove manual line breaks using character code (guarded)
    On Error Resume Next
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = Chr(11)  ' Manual line break (VT)
        .Replacement.text = " "  ' Replace with space
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Do While .Execute(Replace:=wdReplaceOne)
            If Err.Number <> 0 Then Exit Do
            replacementCount = replacementCount + 1
            rng.Collapse wdCollapseEnd
        Loop
    End With
    If Err.Number <> 0 Then
        
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Remove Line Feed chars that aren't paragraph breaks
    On Error Resume Next
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = Chr(10)  ' Line Feed (LF)
        .Replacement.text = " "  ' Replace with space
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Do While .Execute(Replace:=wdReplaceOne)
            If Err.Number <> 0 Then Exit Do
            replacementCount = replacementCount + 1
            rng.Collapse wdCollapseEnd
        Loop
    End With
    If Err.Number <> 0 Then
        
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    
    ApplyTextReplacements = True
    Exit Function

ErrorHandler:
    
    ApplyTextReplacements = False
End Function

 ' (Delegation stubs removed; using module-qualified calls where needed)

'================================================================================
' FORMAT JUSTIFICATIVA/ANEXO PARAGRAPHS - SPECIAL FORMATTING
'================================================================================
Private Function FormatJustificativaAnexoParagraphs(doc As Document) As Boolean: FormatJustificativaAnexoParagraphs = modFormatting.FormatJustificativaAnexoParagraphs(doc): End Function

'================================================================================
' FORMAT NUMBERED PARAGRAPHS
'================================================================================
Private Function FormatNumberedParagraphs(doc As Document) As Boolean: FormatNumberedParagraphs = modFormatting.FormatNumberedParagraphs(doc): End Function

'================================================================================
' HELPER FUNCTIONS FOR PATTERN DETECTION
'================================================================================
 ' Removed: IsVereadorPattern

Private Function IsAnexoPattern(text As String) As Boolean
    Dim cleanText As String
    cleanText = LCase(Trim(text))
    IsAnexoPattern = (cleanText = "anexo" Or cleanText = "anexos")
End Function

Private Function IsAnteOExpostoPattern(text As String) As Boolean
    ' Check if the paragraph starts with "Ante o exposto" (case-insensitive)
    Dim cleanText As String
    cleanText = LCase(Trim(text))
    
    ' Check if empty
    If Len(cleanText) = 0 Then
        IsAnteOExpostoPattern = False
        Exit Function
    End If
    
    ' Check if starts with token
    If Len(cleanText) >= 13 And Left(cleanText, 13) = "ante o exposto" Then
        IsAnteOExpostoPattern = True
    Else
        IsAnteOExpostoPattern = False
    End If
End Function

'================================================================================
' HELPER FUNCTIONS FOR NUMBERED LISTS
'================================================================================
Private Function IsNumberedParagraph(text As String) As Boolean
    ' Check if paragraph starts with a number followed by common separators
    Dim cleanText As String
    cleanText = Trim(text)
    
    ' Check if empty
    If Len(cleanText) = 0 Then
        IsNumberedParagraph = False
        Exit Function
    End If
    
    ' Extract first word/token
    Dim firstToken As String
    Dim spacePos As Long
    spacePos = InStr(cleanText, " ")
    
    If spacePos > 0 Then
        firstToken = Left(cleanText, spacePos - 1)
    Else
        firstToken = cleanText
    End If
    
    ' Check different numbering patterns
    ' Pattern 1: number followed by a dot (1., 2., ...)
    If Len(firstToken) >= 2 And Right(firstToken, 1) = "." Then
        Dim numberPart As String
        numberPart = Left(firstToken, Len(firstToken) - 1)
        If IsNumeric(numberPart) And Val(numberPart) > 0 Then
            ' Verify there is substantive text after the number and punctuation
            If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then
                IsNumberedParagraph = True
                Exit Function
            End If
        End If
    End If
    
    ' Pattern 2: number followed by right parenthesis (1), 2), ...)
    If Len(firstToken) >= 2 And Right(firstToken, 1) = ")" Then
        numberPart = Left(firstToken, Len(firstToken) - 1)
        If IsNumeric(numberPart) And Val(numberPart) > 0 Then
            ' Verify there is substantive text after the number and punctuation
            If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then
                IsNumberedParagraph = True
                Exit Function
            End If
        End If
    End If
    
    ' Pattern 3: parenthesized number ((1), (2), ...)
    If Len(firstToken) >= 3 And Left(firstToken, 1) = "(" And Right(firstToken, 1) = ")" Then
        numberPart = Mid(firstToken, 2, Len(firstToken) - 2)
        If IsNumeric(numberPart) And Val(numberPart) > 0 Then
            ' Verify there is substantive text after the number and punctuation
            If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then
                IsNumberedParagraph = True
                Exit Function
            End If
        End If
    End If
    
    ' Pattern 4: just a number followed by space (1 text, 2 text, ...)
    ' Stricter: must have space AND substantive text after the number
    If IsNumeric(firstToken) And Val(firstToken) > 0 And spacePos > 0 Then
    ' Verify there is substantive text after the number and space
        If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then
            IsNumberedParagraph = True
            Exit Function
        End If
    End If
    
    ' Pattern 5: number followed by other common separators (-, :, ;)
    If Len(firstToken) >= 2 Then
        Dim lastChar As String
        lastChar = Right(firstToken, 1)
        
        If lastChar = "-" Or lastChar = ":" Or lastChar = ";" Then
            numberPart = Left(firstToken, Len(firstToken) - 1)
            If IsNumeric(numberPart) And Val(numberPart) > 0 Then
                ' Verify there is substantive text after the number and punctuation
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
' HAS SUBSTANTIVE TEXT AFTER NUMBER
'================================================================================
Private Function HasSubstantiveTextAfterNumber(fullText As String, numberToken As String) As Boolean
    ' Verify there is substantive text (not only spaces or digits) after the number
    Dim remainingText As String
    Dim startPos As Long
    
    ' Find the position after the number token
    startPos = Len(numberToken) + 1
    
    ' If there is no text after the token, not a valid numbered paragraph
    If startPos > Len(fullText) Then
        HasSubstantiveTextAfterNumber = False
        Exit Function
    End If
    
    ' Extract the remaining text after the number
    remainingText = Trim(Mid(fullText, startPos))
    
    ' Verify there is substantive text
    If Len(remainingText) = 0 Then
        ' No text after the number
        HasSubstantiveTextAfterNumber = False
        Exit Function
    End If
    
    ' Remove spaces and verify there is at least one word with letters
    Dim words() As String
    Dim i As Long
    Dim hasLetters As Boolean
    
    words = Split(remainingText, " ")
    
    For i = 0 To UBound(words)
        Dim word As String
        word = Trim(words(i))
        
    ' Verify the word contains at least one letter (not just digits/punctuation)
        If ContainsLetters(word) And Len(word) >= 2 Then
            HasSubstantiveTextAfterNumber = True
            Exit Function
        End If
    Next i
    
    ' If we got here, no substantive text was found
    HasSubstantiveTextAfterNumber = False
End Function

'================================================================================
' CONTAINS LETTERS
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
    ' Remove the manual number from the beginning of the paragraph
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
        
        ' Remove the first token if it is a number with separators
        If (Len(firstToken) >= 2 And (Right(firstToken, 1) = "." Or Right(firstToken, 1) = ")")) Or _
           (Len(firstToken) >= 3 And Left(firstToken, 1) = "(" And Right(firstToken, 1) = ")") Or _
           (IsNumeric(firstToken) And Val(firstToken) > 0) Then
            
            ' Remove the first token and extra spaces
            RemoveManualNumber = Trim(Mid(cleanText, spacePos + 1))
        Else
            RemoveManualNumber = cleanText
        End If
    Else
        RemoveManualNumber = cleanText
    End If
End Function

'================================================================================
' PUBLIC SUB: OPEN LOGS FOLDER
'================================================================================
Public Sub OpenLogsFolder()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim logsFolder As String
    Dim defaultLogsFolder As String
    
    ' Try to get the active document
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler
    
    ' Define logs folder based on the document folder or TEMP
    If Not doc Is Nothing And doc.Path <> "" Then
        logsFolder = doc.Path
    Else
        logsFolder = Environ("TEMP")
    End If
    
    ' Ensure the folder exists
    If Dir(logsFolder, vbDirectory) = "" Then
        logsFolder = Environ("TEMP")
    End If
    
    ' Open the folder in Windows Explorer
    Shell "explorer.exe """ & logsFolder & """", vbNormalFocus
    
    Application.StatusBar = "Logs folder opened: " & logsFolder
    
    
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Error opening logs folder"
    
    ' Fallback: try opening TEMP folder
    On Error Resume Next
    Shell "explorer.exe """ & Environ("TEMP") & """", vbNormalFocus
    If Err.Number = 0 Then
        Application.StatusBar = "TEMP folder opened as fallback"
    Else
        Application.StatusBar = "Could not open logs folder"
    End If
End Sub


'================================================================================
' BACKUP SYSTEM - SAFETY FEATURE
'================================================================================
Private Function CreateDocumentBackup(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    CreateDocumentBackup = False
    
    ' Initial document validation
    If doc Is Nothing Then
    
        Exit Function
    End If
    
    ' Do not create backup if the document has not been saved yet
    If doc.Path = "" Then
    ' (Logging removed) Backup skipped - unsaved document
        CreateDocumentBackup = True
        Exit Function
    End If
    
    ' Check if document isn't corrupted/inaccessible
    On Error Resume Next
    Dim testAccess As String
    testAccess = doc.Name
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
    ' (Logging removed) Backup error: document inaccessible
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
    
    ' Verify FSO was created successfully
    If fso Is Nothing Then
    ' (Logging removed) Backup error: could not create FileSystemObject
        Exit Function
    End If
    
    ' Choose backup folder same as logs folder: document folder or TEMP
    If doc.Path <> "" Then
        backupFolder = doc.Path
    Else
        backupFolder = Environ("TEMP")
    End If
    
    ' Verify write permissions in the chosen folder
    On Error Resume Next
    Dim testFile As String
    testFile = backupFolder & "\test_write_" & Format(Now, "HHmmss") & ".tmp"
    Open testFile For Output As #1
    Close #1
    Kill testFile
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
    ' (Logging removed) Backup error: no write permissions to folder
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Extract document base name and extension with validation
    On Error Resume Next
    docName = fso.GetBaseName(doc.Name)
    docExtension = fso.GetExtensionName(doc.Name)
    If Err.Number <> 0 Or docName = "" Then
        On Error GoTo ErrorHandler
    ' (Logging removed) Backup error: invalid file name
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Create timestamp for backup
    timestamp = Format(Now, "yyyy-mm-dd_HHmmss")
    
    ' Backup file name
    backupFileName = docName & "_backup_" & timestamp & "." & docExtension
    backupFilePath = backupFolder & "\" & backupFileName
    
    ' Save a copy of the document as backup with retry
    Application.StatusBar = "Creating document backup..."
    
    ' Save the current document first to ensure it's up to date
    On Error Resume Next
    doc.Save
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
    ' (Logging removed) Warning: could not save document before backup: " & Err.Description
    End If
    On Error GoTo ErrorHandler
    
    ' Create a file copy using FileSystemObject with retry
    For retryCount = 1 To Config.maxRetryAttempts
        On Error Resume Next
        fso.CopyFile doc.FullName, backupFilePath, True
        If Err.Number = 0 Then
            On Error GoTo ErrorHandler
            Exit For
        Else
            On Error GoTo ErrorHandler
            ' (Logging removed) Backup attempt " & retryCount & " failed: " & Err.Description
            If retryCount < Config.maxRetryAttempts Then
                ' Wait briefly before trying again
                Sleep Config.retryDelayMs ' according to config
            End If
        End If
    Next retryCount
    
    ' Verify backup was created
    If Not fso.FileExists(backupFilePath) Then
    ' (Logging removed) Backup error: file was not created
        Exit Function
    End If
    
    ' Clean old backups if needed (now in the same folder as logs)
    CleanOldBackups backupFolder, docName
    
    ' (Logging removed) Backup created successfully: " & backupFileName
    Application.StatusBar = "Backup created - processing document..."
    
    CreateDocumentBackup = True
    Exit Function

ErrorHandler:
    ' (Logging removed) Critical error creating backup: " & Err.Description & " (Line: " & Erl & ")"
    Application.StatusBar = "Backup creation failed"
    CreateDocumentBackup = False
    
    ' Limpeza de recursos
    On Error Resume Next
    Set fso = Nothing
End Function

'================================================================================
' CLEAN OLD BACKUPS - SIMPLIFIED
'================================================================================
Private Sub CleanOldBackups(backupFolder As String, docBaseName As String)
    On Error Resume Next
    
    ' Simplified cleanup - only warns when there are many files
    Dim fso As Object
    Dim folder As Object
    Dim filesCount As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(backupFolder)
    
    filesCount = folder.Files.count
    
    ' If more than 15 files exist in the backups folder, log a warning
    If filesCount > 15 Then
    ' (Logging removed) Too many backups in folder (" & filesCount & " files) - consider manual cleanup
    End If
End Sub

'================================================================================
' PUBLIC SUB: OPEN BACKUPS FOLDER
'================================================================================
Public Sub OpenBackupsFolder()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim backupFolder As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Try to get active document
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler
    
    ' Define backups folder to match logs folder: document folder or TEMP
    If Not doc Is Nothing And doc.Path <> "" Then
        backupFolder = doc.Path
    Else
        backupFolder = Environ("TEMP")
    End If
    
    ' Check if folder exists
    If Not fso.FolderExists(backupFolder) Then
        Application.StatusBar = "Folder not found"
    ' (Logging removed) Backups/logs folder not found: " & backupFolder
        Exit Sub
    End If
    
    ' Open the folder in Windows Explorer
    Shell "explorer.exe """ & backupFolder & """", vbNormalFocus
    
    Application.StatusBar = "Backups folder opened: " & backupFolder
    
    ' (Logging removed)
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Error opening backups folder"
    ' (Logging removed) Error opening backups folder: " & Err.Description
    
    ' Fallback: try to open the TEMP folder or document folder
    On Error Resume Next
    If fso.FolderExists(Environ("TEMP")) Then
        Shell "explorer.exe """ & Environ("TEMP") & """", vbNormalFocus
        Application.StatusBar = "TEMP folder opened as fallback"
    ElseIf Not doc Is Nothing And doc.Path <> "" Then
        Shell "explorer.exe """ & doc.Path & """", vbNormalFocus
        Application.StatusBar = "Document folder opened as fallback"
    Else
        Application.StatusBar = "Could not open backups folder"
    End If
End Sub

'================================================================================
' CLEAN MULTIPLE SPACES - FINAL PASS
'================================================================================
Private Function CleanMultipleSpaces(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Cleaning multiple spaces..."
    
    Dim rng As Range
    Dim spacesRemoved As Long
    Dim totalOperations As Long
    
    ' SUPER OPTIMIZED: Operations consolidated into a single Find configuration
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
        
    ' OPTIMIZATION 1: Remove multiple spaces (2 or more) in one operation
    ' Uses an optimized loop that progressively reduces spaces
        Do
            .text = "  "  ' Two spaces
            .Replacement.text = " "  ' One space
            
            Dim currentReplaceCount As Long
            currentReplaceCount = 0
            
            ' Execute until no more doubles are found
            Do While .Execute(Replace:=wdReplaceOne)
                currentReplaceCount = currentReplaceCount + 1
                spacesRemoved = spacesRemoved + 1
                rng.Collapse wdCollapseEnd
                ' Optimized protection - check every 200 operations
                If currentReplaceCount Mod 200 = 0 Then
                    DoEvents
                    If spacesRemoved > 2000 Then Exit Do
                End If
            Loop
            
            totalOperations = totalOperations + 1
            ' If no more doubles found or limit reached, stop
            If currentReplaceCount = 0 Or totalOperations > 10 Then Exit Do
        Loop
    End With
    
    ' OPTIMIZATION 2: Consolidated line break cleanup operations
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
    .MatchWildcards = False  ' Use simple Find/Replace for compatibility
        
    ' Remove multiple spaces before breaks - iterative method
    .text = "  ^p"  ' 2 spaces followed by paragraph break
    .Replacement.text = " ^p"  ' 1 space followed by paragraph break
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
    ' Second pass to ensure complete cleanup
    .text = " ^p"  ' Space before break
        .Replacement.text = "^p"
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
    ' Remove multiple spaces after breaks - iterative method
    .text = "^p  "  ' Break followed by 2 spaces
    .Replacement.text = "^p "  ' Break followed by 1 space
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2000 Then Exit Do
        Loop
    End With
    
    ' OPTIMIZATION 3: Consolidated and optimized tab cleanup
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
    .MatchWildcards = False  ' Use simple Find/Replace
        
    ' Remove multiple tabs iteratively
    .text = "^t^t"  ' 2 tabs
    .Replacement.text = "^t"  ' 1 tab
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
    ' Convert tabs to spaces
        .text = "^t"
        .Replacement.text = " "
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
    End With
    
    ' OPTIMIZATION 4: Final ultra-fast check for remaining double spaces
    Set rng = doc.Range
    With rng.Find
        .text = "  "
        .Replacement.text = " "
        .MatchWildcards = False
        .Forward = True
    .Wrap = wdFindStop  ' Faster than wdFindContinue
        
        Dim finalCleanCount As Long
        Do While .Execute(Replace:=wdReplaceOne) And finalCleanCount < 100
            finalCleanCount = finalCleanCount + 1
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
        Loop
    End With
    
    ' SPECIFIC PROTECTION: Ensure space after CONSIDERANDO
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchCase = False
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        
    ' Fix CONSIDERANDO stuck to the next word
        .text = "CONSIDERANDOa"
        .Replacement.text = "CONSIDERANDO a"
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2100 Then Exit Do
        Loop
        
        .text = "CONSIDERANDOe"
        .Replacement.text = "CONSIDERANDO e"
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2100 Then Exit Do
        Loop
        
        .text = "CONSIDERANDOo"
        .Replacement.text = "CONSIDERANDO o"
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2100 Then Exit Do
        Loop
        
        .text = "CONSIDERANDOq"
        .Replacement.text = "CONSIDERANDO q"
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2100 Then Exit Do
        Loop
    End With
    
    ' (Logging removed) Space cleanup complete: " & spacesRemoved & " corrections applied (with CONSIDERANDO protection)
    CleanMultipleSpaces = True
    Exit Function

ErrorHandler:
    ' (Logging removed) Error cleaning multiple spaces: " & Err.Description
    CleanMultipleSpaces = False ' Do not fail the process because of this
End Function

'================================================================================
' LIMIT SEQUENTIAL EMPTY LINES - CONTROL CONSECUTIVE BLANK LINES
'================================================================================
Private Function LimitSequentialEmptyLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Controlling consecutive blank lines..."
    
    ' Identify the second paragraph for protection
    Dim secondParaIndex As Long
    secondParaIndex = GetSecondParagraphIndex(doc)
    
    ' SUPER OPTIMIZED: Use Find/Replace without wildcard for faster operation and compatibility
    Dim rng As Range
    Dim linesRemoved As Long
    Dim totalReplaces As Long
    Dim passCount As Long
    
    passCount = 1 ' Initialize pass counter
    
    Set rng = doc.Range
    
    ' ULTRA-FAST METHOD: Remove multiple consecutive breaks
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
    .MatchWildcards = False  ' Use simple Find/Replace for compatibility
        
    ' Remove multiple consecutive breaks iteratively
    .text = "^p^p^p^p"  ' 4 breaks
    .Replacement.text = "^p^p"  ' 2 breaks
        
        Do While .Execute(Replace:=True)
            linesRemoved = linesRemoved + 1
            totalReplaces = totalReplaces + 1
            If totalReplaces > 500 Then Exit Do
            If linesRemoved Mod 50 = 0 Then DoEvents
        Loop
        
    ' Convert 3 breaks -> 2 breaks
    .text = "^p^p^p"  ' 3 breaks
    .Replacement.text = "^p^p"  ' 2 breaks
        
        Do While .Execute(Replace:=True)
            linesRemoved = linesRemoved + 1
            totalReplaces = totalReplaces + 1
            If totalReplaces > 500 Then Exit Do
            If linesRemoved Mod 50 = 0 Then DoEvents
        Loop
    End With
    
    ' SECOND PASS: Remove remaining double breaks (2 breaks -> 1 break)
    If totalReplaces > 0 Then passCount = passCount + 1
    
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindContinue
        
    ' Convert triple breaks to double
    .text = "^p^p^p"  ' 3 breaks
    .Replacement.text = "^p^p"  ' 2 breaks
        
        Dim secondPassCount As Long
        Do While .Execute(Replace:=True) And secondPassCount < 200
            secondPassCount = secondPassCount + 1
            linesRemoved = linesRemoved + 1
        Loop
    End With
    
    ' FINAL CHECK: Ensure there is no more than 1 consecutive blank line
    If secondPassCount > 0 Then passCount = passCount + 1
    
    ' Hybrid method: Find/Replace for simple cases + loop only if necessary
    Set rng = doc.Range
    With rng.Find
    .text = "^p^p^p"  ' 3 breaks (2 blank lines + content)
    .Replacement.text = "^p^p"  ' 2 breaks (1 blank line + content)
        .MatchWildcards = False
        
        Dim finalPassCount As Long
        Do While .Execute(Replace:=True) And finalPassCount < 100
            finalPassCount = finalPassCount + 1
            linesRemoved = linesRemoved + 1
        Loop
    End With
    
    If finalPassCount > 0 Then passCount = passCount + 1
    
    ' OPTIMIZED FALLBACK: If issues remain, use limited traditional method
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
            
            ' Check if the paragraph is empty
            If paraText = "" And Not HasVisualContent(para) Then
                emptyLineCount = emptyLineCount + 1
                
                ' If we already have more than 1 consecutive blank line, remove this one
                If emptyLineCount > 1 Then
                    para.Range.Delete
                    fallbackRemoved = fallbackRemoved + 1
                    linesRemoved = linesRemoved + 1
                    ' Do not increment i because we removed a paragraph
                Else
                    i = i + 1
                End If
            Else
                ' If content found, reset the counter
                emptyLineCount = 0
                i = i + 1
            End If
            
            ' Responsiveness and optimized protections
            If fallbackRemoved Mod 10 = 0 Then DoEvents
            If i > 500 Then Exit Do ' Additional protection
        Loop
    End If
    
    ' (Logging removed) Consecutive blank lines control completed in " & passCount & " pass(es): " & linesRemoved & " extra line(s) removed (max 1 in a row)
    LimitSequentialEmptyLines = True
    Exit Function

ErrorHandler:
    ' (Logging removed) Error controlling blank lines: " & Err.Description
    LimitSequentialEmptyLines = False ' Do not fail the process because of this
End Function

'================================================================================
' ENSURE PARAGRAPH SEPARATION
'================================================================================
Private Function EnsureParagraphSeparation(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Ensuring minimum separation between paragraphs..."
    
    Dim para As Paragraph
    Dim nextPara As Paragraph
    Dim i As Long
    Dim insertedCount As Long
    Dim totalChecked As Long
    
    ' Iterate all paragraphs ensuring at least one blank line between non-empty ones
    For i = 1 To doc.Paragraphs.count - 1 ' -1 because we check the next paragraph
        Set para = doc.Paragraphs(i)
        Set nextPara = doc.Paragraphs(i + 1)
        
        totalChecked = totalChecked + 1
        
    ' Extract the text of both paragraphs for analysis
        Dim paraText As String
        Dim nextParaText As String
        
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        nextParaText = Trim(Replace(Replace(nextPara.Range.text, vbCr, ""), vbLf, ""))
        
    ' Check if both paragraphs contain text (not blank lines)
        If paraText <> "" And nextParaText <> "" Then
            ' Check if paragraphs are adjacent (no blank line between them)
            ' For that, check if the end of the current paragraph is immediately followed by the start of the next
            
            Dim currentParaEnd As Long
            Dim nextParaStart As Long
            
            currentParaEnd = para.Range.End
            nextParaStart = nextPara.Range.Start
            
            ' If the difference between the end of one paragraph and the start of the next is only 1 char,
            ' they are directly adjacent (no blank line)
            If nextParaStart - currentParaEnd <= 1 Then
                ' Insert a blank line between the paragraphs
                Dim insertRange As Range
                Set insertRange = doc.Range(currentParaEnd - 1, currentParaEnd - 1)
                insertRange.text = vbCrLf
                
                insertedCount = insertedCount + 1
                
                ' Update paragraph references after insertion
                ' because indices may have changed
                On Error Resume Next
                Set para = doc.Paragraphs(i)
                Set nextPara = doc.Paragraphs(i + 2) ' +2 because we inserted a line
                On Error GoTo ErrorHandler
                
                ' Log only for the first cases or significant cases
                If insertedCount <= 10 Or insertedCount Mod 50 = 0 Then
                    ' (Logging removed) Blank line inserted between paragraphs " & i & " and " & (i + 1) & " (total: " & insertedCount & ")"
                End If
            End If
        End If
        
    ' Performance and responsiveness control
        If totalChecked Mod 100 = 0 Then
            DoEvents
            Application.StatusBar = "Checking paragraph separation... " & totalChecked & " checked"
        End If
        
        ' Protection against very large documents
        If totalChecked > 5000 Then
            ' (Logging removed) Verification limit reached (5000 paragraphs) - stopping verification
            Exit For
        End If
    Next i
    
    ' (Logging removed) Paragraph separation ensured: " & insertedCount & " blank line(s) inserted out of " & totalChecked & " pairs checked"
    EnsureParagraphSeparation = True
    Exit Function

ErrorHandler:
    ' (Logging removed) Error ensuring paragraph separation: " & Err.Description
    EnsureParagraphSeparation = False
End Function

'================================================================================
' CONFIGURE DOCUMENT VIEW
'================================================================================
Private Function ConfigureDocumentView(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Configuring document view..."
    
    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow
    
    ' Configure ONLY the zoom to 110% - all other settings are preserved
    With docWindow.View
        .Zoom.Percentage = 110
    ' Do NOT change the view type - preserve original
    End With
    
    ' Remove settings that changed global Word settings (now preserved)
    
    ' (Logging removed) View configured: zoom set to 110%, other settings preserved
    ConfigureDocumentView = True
    Exit Function

ErrorHandler:
    ' (Logging removed) Error configuring view: " & Err.Description
    ConfigureDocumentView = False ' Do not fail the process because of this
End Function

'================================================================================
' SAVE AND EXIT - ROBUST PUBLIC SUBROUTINE
'================================================================================
Public Sub SaveAndExit()
    On Error GoTo CriticalErrorHandler
    
    Dim startTime As Date
    startTime = Now
    
    Application.StatusBar = "Checking open documents..."
    ' (Logging removed) Starting save-and-exit process - checking documents
    
    ' Check if there are open documents
    If Application.Documents.count = 0 Then
    Application.StatusBar = "No documents open - closing Word"
    ' (Logging removed) No documents open - closing application
        Application.Quit SaveChanges:=wdDoNotSaveChanges
        Exit Sub
    End If
    
    ' Collect info about unsaved documents
    Dim unsavedDocs As Collection
    Set unsavedDocs = New Collection
    
    Dim doc As Document
    Dim i As Long
    
    ' Check each open document
    For i = 1 To Application.Documents.count
        Set doc = Application.Documents(i)
        
        On Error Resume Next
    ' Check if the document has unsaved changes
        If doc.Saved = False Or doc.Path = "" Then
            unsavedDocs.Add doc.Name
            ' (Logging removed) Unsaved document detected: " & doc.Name
        End If
        On Error GoTo CriticalErrorHandler
    Next i
    
    ' If no unsaved documents, close directly
    If unsavedDocs.count = 0 Then
    Application.StatusBar = "All documents saved - closing Word"
    ' (Logging removed) All documents are saved - closing application
        Application.Quit SaveChanges:=wdDoNotSaveChanges
        Exit Sub
    End If
    
    ' Build detailed message about unsaved documents
    Dim message As String
    Dim docList As String
    
    For i = 1 To unsavedDocs.count
        docList = docList & "• " & unsavedDocs(i) & vbCrLf
    Next i
    
    message = "ATTENTION: There are " & unsavedDocs.count & " document(s) with unsaved changes:" & vbCrLf & vbCrLf
    message = message & docList & vbCrLf
    message = message & "Do you want to save all documents before exiting?" & vbCrLf & vbCrLf
    message = message & "• YES: Save all and close Word" & vbCrLf
    message = message & "• NO: Close without saving (you will LOSE changes)" & vbCrLf
    message = message & "• CANCEL: Cancel the operation"
    
    ' Present options to the user
    Application.StatusBar = "Waiting for user decision about unsaved documents..."
    
    Dim userChoice As VbMsgBoxResult
    userChoice = MsgBox(NormalizeForUI(message), vbYesNoCancel + vbExclamation + vbDefaultButton1, _
                        NormalizeForUI(SYSTEM_NAME & " - Save and Exit (" & unsavedDocs.count & " unsaved document(s))"))
    
    Select Case userChoice
        Case vbYes
            ' User chose to save all
            Application.StatusBar = "Saving all documents..."
            ' (Logging removed) User chose to save all documents before exiting
            
            If SalvarTodosDocumentos() Then
                Application.StatusBar = "Documents saved successfully - closing Word"
                ' (Logging removed) All documents saved successfully - closing application
                Application.Quit SaveChanges:=wdDoNotSaveChanges
            Else
          Application.StatusBar = "Error saving documents - operation cancelled"
          ' (Logging removed) Failed to save some documents - exit operation cancelled
          MsgBox NormalizeForUI(MSG_SAVE_ERROR), _
              vbCritical, NormalizeForUI(TITLE_SAVE_ERROR)
            End If
            
        Case vbNo
            ' User chose not to save
            Dim confirmMessage As String
            confirmMessage = "FINAL CONFIRMATION:" & vbCrLf & vbCrLf & _
                "You are about to CLOSE WORD WITHOUT SAVING " & unsavedDocs.count & " document(s)." & vbCrLf & vbCrLf & _
                "ALL UNSAVED CHANGES WILL BE LOST!" & vbCrLf & vbCrLf & _
                "Are you absolutely sure?"
            
            Dim finalConfirm As VbMsgBoxResult
            finalConfirm = MsgBox(NormalizeForUI(confirmMessage), vbYesNo + vbCritical + vbDefaultButton2, _
                                  NormalizeForUI(TITLE_FINAL_CONFIRM))
            
            If finalConfirm = vbYes Then
                Application.StatusBar = "Closing Word without saving changes..."
                ' (Logging removed) User confirmed closing without saving - closing application
                Application.Quit SaveChanges:=wdDoNotSaveChanges
            Else
          Application.StatusBar = "Operation cancelled by user"
          ' (Logging removed) User cancelled closing without saving
          MsgBox NormalizeForUI(MSG_OPERATION_CANCELLED), _
              vbInformation, NormalizeForUI(TITLE_OPERATION_CANCELLED)
            End If
            
        Case vbCancel
         ' User cancelled
         Application.StatusBar = "Exit operation cancelled by user"
         ' (Logging removed) User cancelled save and exit operation
         MsgBox NormalizeForUI(MSG_OPERATION_CANCELLED), _
             vbInformation, NormalizeForUI(TITLE_OPERATION_CANCELLED)
    End Select
    
    Application.StatusBar = False
    ' (Logging removed) Save-and-exit process completed in " & Format(Now - startTime, "hh:mm:ss")
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "CRITICAL ERROR in Save and Exit operation #" & Err.Number & ": " & Err.Description
    
    ' (Logging removed) " & errDesc
    Application.StatusBar = "Critical error - operation cancelled"
    
    Dim critMsg As String
    critMsg = ReplacePlaceholders(MSG_CRITICAL_SAVE_EXIT, "ERR", Err.Description)
    MsgBox NormalizeForUI(critMsg), _
        vbCritical, NormalizeForUI(TITLE_CRITICAL_SAVE_EXIT)
End Sub

'================================================================================
' SAVE ALL DOCUMENTS - PRIVATE HELPER FUNCTION
'================================================================================
Private Function SalvarTodosDocumentos() As Boolean
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim i As Long
    Dim savedCount As Long
    Dim errorCount As Long
    Dim totalDocs As Long
    
    totalDocs = Application.Documents.count
    
    ' Save each document individually
    For i = 1 To totalDocs
        Set doc = Application.Documents(i)
        
    Application.StatusBar = "Saving document " & i & " of " & totalDocs & ": " & doc.Name
        
        On Error Resume Next
        
    ' If the document has never been saved (no path), open dialog
        If doc.Path = "" Then
            Dim saveDialog As Object
            Set saveDialog = Application.FileDialog(msoFileDialogSaveAs)
            
            With saveDialog
                .Title = "Save document: " & doc.Name
                .InitialFileName = doc.Name
                
                If .Show = -1 Then
                    doc.SaveAs2 .SelectedItems(1)
                    If Err.Number = 0 Then
                        savedCount = savedCount + 1
                        ' (Logging removed) Document saved as new file: " & doc.Name
                    Else
                        errorCount = errorCount + 1
                        ' (Logging removed) Error saving document as new: " & doc.Name & " - " & Err.Description
                    End If
                Else
                    errorCount = errorCount + 1
                    ' (Logging removed) Save cancelled by user: " & doc.Name
                End If
            End With
        Else
            ' Document already has a path, just save it
            doc.Save
            If Err.Number = 0 Then
                savedCount = savedCount + 1
                ' (Logging removed) Document saved: " & doc.Name
            Else
                errorCount = errorCount + 1
                ' (Logging removed) Error saving document: " & doc.Name & " - " & Err.Description
            End If
        End If
        
        On Error GoTo ErrorHandler
    Next i
    
    ' Verify result
    If errorCount = 0 Then
    ' (Logging removed) All documents saved successfully: " & savedCount & " of " & totalDocs
        SalvarTodosDocumentos = True
    Else
    ' (Logging removed) Partial save failure: " & savedCount & " saved, " & errorCount & " errors"
        SalvarTodosDocumentos = False
    End If
    
    Exit Function

ErrorHandler:
    ' (Logging removed) Critical error saving documents: " & Err.Description
    SalvarTodosDocumentos = False
End Function

'================================================================================

'================================================================================
' GET CLIPBOARD DATA - Get data from the clipboard
'================================================================================
Private Function GetClipboardData() As Variant
    On Error GoTo ErrorHandler
    
    ' Placeholder for clipboard data
    ' In a complete implementation, Windows APIs or advanced methods would be needed
    ' to capture binary data
    GetClipboardData = "ImageDataPlaceholder"
    Exit Function

ErrorHandler:
    GetClipboardData = Empty
End Function

'================================================================================
' ENHANCED IMAGE PROTECTION - Improved protection during formatting
'================================================================================
Private Function ProtectImagesInRange(targetRange As Range) As Boolean
    On Error GoTo ErrorHandler
    
    ' Check if there are images in the range before applying formatting
    If targetRange.InlineShapes.count > 0 Then
    ' OPTIMIZED: Apply formatting character by character, protecting images
        Dim i As Long
        Dim charRange As Range
        Dim charCount As Long
        charCount = SafeGetCharacterCount(targetRange) ' Cache da contagem segura
        
        If charCount > 0 Then ' Safety check
            For i = 1 To charCount
                Set charRange = targetRange.Characters(i)
                ' Only format characters that are not part of images
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
        ' Range without images - full normal formatting
        With targetRange.Font
            .Name = STANDARD_FONT
            .size = STANDARD_FONT_SIZE
            .Color = wdColorAutomatic
        End With
    End If
    
    ' (Legacy image protection removed) Always succeed
    ProtectImagesInRange = True: Exit Function

ErrorHandler:
    ProtectImagesInRange = True ' Fail-open; formatting continues
End Function

'================================================================================
' VISUAL ELEMENTS CLEANUP SYSTEM
'================================================================================

'================================================================================
' DELETE HIDDEN VISUAL ELEMENTS - Remove all hidden visual elements
'================================================================================
Private Function DeleteHiddenVisualElements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Removing hidden visual elements..."
    
    Dim deletedCount As Long
    deletedCount = 0
    
    ' Remove hidden shapes (floating)
    Dim i As Long
    For i = doc.Shapes.count To 1 Step -1
        Dim shp As shape
        Set shp = doc.Shapes(i)
        
    ' Check if the shape is hidden (multiple criteria)
        Dim isHidden As Boolean
        isHidden = False
        
    ' Shape marked as not visible
        If Not shp.Visible Then isHidden = True
        
    ' Shape with total transparency
        On Error Resume Next
        If shp.Fill.Transparency >= 0.99 Then isHidden = True
        On Error GoTo ErrorHandler
        
    ' Shape with zero or nearly zero size
        If shp.Width <= 1 Or shp.Height <= 1 Then isHidden = True
        
    ' Shape positioned outside the visible page (very negative coordinates)
        If shp.Left < -1000 Or shp.Top < -1000 Then isHidden = True
        
        If isHidden Then
            ' (Logging removed) Removing hidden shape (type: " & shp.Type & ", index: " & i & ")
            shp.Delete
            deletedCount = deletedCount + 1
        End If
    Next i
    
    ' Remove hidden inline objects
    For i = doc.InlineShapes.count To 1 Step -1
        Dim inlineShp As InlineShape
        Set inlineShp = doc.InlineShapes(i)
        
        Dim isInlineHidden As Boolean
        isInlineHidden = False
        
    ' Inline object in hidden text
        If inlineShp.Range.Font.Hidden Then isInlineHidden = True
        
    ' Inline object in paragraph with zero spacing (likely hidden)
        If inlineShp.Range.ParagraphFormat.LineSpacing = 0 Then isInlineHidden = True
        
    ' Inline object with zero size
        If inlineShp.Width <= 1 Or inlineShp.Height <= 1 Then isInlineHidden = True
        
        If isInlineHidden Then
            ' (Logging removed) Removing hidden inline object (type: " & inlineShp.Type & ")
            inlineShp.Delete
            deletedCount = deletedCount + 1
        End If
    Next i
    
    ' (Logging removed) Removal of hidden elements completed: " & deletedCount & " element(s) removed
    DeleteHiddenVisualElements = True
    Exit Function

ErrorHandler:
    ' (Logging removed) Error removing hidden visual elements: " & Err.Description
    DeleteHiddenVisualElements = False
End Function

'================================================================================
' DELETE VISUAL ELEMENTS IN RANGE - Remove visual elements between paragraphs 1-4
'================================================================================
Private Function DeleteVisualElementsInFirstFourParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Removing visual elements between paragraphs 1-4..."
    
    If doc.Paragraphs.count < 1 Then
    ' (Logging removed) Document has no paragraphs - skipping visual elements cleanup
        DeleteVisualElementsInFirstFourParagraphs = True
        Exit Function
    End If
    
    If doc.Paragraphs.count < 4 Then
    ' (Logging removed) Document has less than 4 paragraphs - removing elements from existing paragraphs (" & doc.Paragraphs.count & " paragraphs)
    End If
    
    Dim deletedCount As Long
    deletedCount = 0
    
    ' Define the range of the first 4 paragraphs (or less if the document is shorter)
    Dim maxParagraphs As Long
    If doc.Paragraphs.count < 4 Then
        maxParagraphs = doc.Paragraphs.count
    Else
        maxParagraphs = 4
    End If
    
    Dim startRange As Long
    Dim endRange As Long
    startRange = doc.Paragraphs(1).Range.Start
    endRange = doc.Paragraphs(maxParagraphs).Range.End
    
    ' (Logging removed) Removing visual elements from paragraphs 1 to " & maxParagraphs & " (position " & startRange & " to " & endRange & ")
    
    ' Remove floating shapes anchored within the first 4 paragraphs' range
    Dim i As Long
    For i = doc.Shapes.count To 1 Step -1
        Dim shp As shape
        Set shp = doc.Shapes(i)
        
    ' Check if the shape is anchored within the first 4 paragraphs' range
        On Error Resume Next
        Dim anchorPosition As Long
        anchorPosition = shp.Anchor.Start
        On Error GoTo ErrorHandler
        
        If anchorPosition >= startRange And anchorPosition <= endRange Then
            Dim paragraphNum As Long
            paragraphNum = GetParagraphNumber(doc, anchorPosition)
            ' (Logging removed) Removing shape (type: " & shp.Type & ") anchored at paragraph " & paragraphNum
            shp.Delete
            deletedCount = deletedCount + 1
        End If
    Next i
    
    ' Remove inline objects in the first 4 paragraphs
    For i = doc.InlineShapes.count To 1 Step -1
        Dim inlineShp As InlineShape
        Set inlineShp = doc.InlineShapes(i)
        
    ' Check if the inline object is within the first 4 paragraphs
        If inlineShp.Range.Start >= startRange And inlineShp.Range.Start <= endRange Then
            Dim inlineParagraphNum As Long
            inlineParagraphNum = GetParagraphNumber(doc, inlineShp.Range.Start)
            ' (Logging removed) Removing inline object (type: " & inlineShp.Type & ") at paragraph " & inlineParagraphNum
            inlineShp.Delete
            deletedCount = deletedCount + 1
        End If
    Next i
    
    ' (Logging removed) Removal of visual elements from the first " & maxParagraphs & " paragraphs completed: " & deletedCount & " element(s) removed
    DeleteVisualElementsInFirstFourParagraphs = True
    Exit Function

ErrorHandler:
    ' (Logging removed) Error removing visual elements from the first 4 paragraphs: " & Err.Description
    DeleteVisualElementsInFirstFourParagraphs = False
End Function

'================================================================================
' GET PARAGRAPH NUMBER - Helper to determine paragraph number
'================================================================================
Private Function GetParagraphNumber(doc As Document, position As Long) As Long
    Dim i As Long
    For i = 1 To doc.Paragraphs.count
        If position >= doc.Paragraphs(i).Range.Start And position <= doc.Paragraphs(i).Range.End Then
            GetParagraphNumber = i
            Exit Function
        End If
    Next i
    GetParagraphNumber = 0 ' Not found
End Function

'================================================================================
' VIEW SETTINGS PROTECTION SYSTEM
'================================================================================

'================================================================================
' BACKUP VIEW SETTINGS - Save original view settings
'================================================================================
Private Function BackupViewSettings(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Backing up view settings..."
    
    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow
    
    ' Backup view settings
    With originalViewSettings
        .ViewType = docWindow.View.Type
    ' Rulers are controlled by Window, not by View
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
    ' .ShowAnimation removed - may not exist in all versions
        .DraftFont = docWindow.View.Draft
        .WrapToWindow = docWindow.View.WrapToWindow
        .ShowPicturePlaceHolders = docWindow.View.ShowPicturePlaceHolders
        .ShowFieldShading = docWindow.View.FieldShading
        .TableGridlines = docWindow.View.TableGridlines
    ' .EnlargeFontsLessThan removed - may not exist in all versions
    End With
    
    ' (Logging removed) Backup of view settings completed
    BackupViewSettings = True
    Exit Function

ErrorHandler:
    ' (Logging removed) Error backing up view settings: " & Err.Description
    BackupViewSettings = False
End Function

'================================================================================
' RESTORE VIEW SETTINGS - Restore original view settings
'================================================================================
Private Function RestoreViewSettings(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Restoring original view settings..."
    
    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow
    
    ' Restore all original settings, EXCEPT the zoom
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
    ' .ShowAnimation removed for compatibility
        .Draft = originalViewSettings.DraftFont
        .WrapToWindow = originalViewSettings.WrapToWindow
        .ShowPicturePlaceHolders = originalViewSettings.ShowPicturePlaceHolders
        .FieldShading = originalViewSettings.ShowFieldShading
        .TableGridlines = originalViewSettings.TableGridlines
    ' .EnlargeFontsLessThan removed for compatibility
        
    ' ZOOM kept at 110% - the only setting that remains changed
        .Zoom.Percentage = 110
    End With
    
    ' Window-specific settings (for rulers)
    docWindow.DisplayRulers = originalViewSettings.ShowHorizontalRuler
    docWindow.DisplayVerticalRuler = originalViewSettings.ShowVerticalRuler
    
    ' (Logging removed) Original view settings restored (zoom kept at 110%)
    RestoreViewSettings = True
    Exit Function

ErrorHandler:
    ' (Logging removed) Error restoring view settings: " & Err.Description
    RestoreViewSettings = False
End Function

'================================================================================
' CLEANUP VIEW SETTINGS - Reset stored view settings variables
'================================================================================
Private Sub CleanupViewSettings()
    On Error Resume Next
    
    ' Reset the settings structure
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
    ' .ShowAnimation removed for compatibility
        .DraftFont = False
        .WrapToWindow = False
        .ShowPicturePlaceHolders = False
        .ShowFieldShading = 0
        .TableGridlines = False
    ' .EnlargeFontsLessThan removed for compatibility
    End With
    
    ' (Logging removed) View settings variables cleaned
End Sub
