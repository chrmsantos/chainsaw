' SPDX-License-Identifier: GPL-3.0-or-later
' =============================================================================
' PROJECT: CHAINSAW PROPOSITURAS
' =============================================================================
'
' Automated system for standardizing legislative proposers in Microsoft Word
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
' Version: 1.0.0-Beta3 | Date: 2025-10-10
' Repository: github.com/chrmsantos/chainsaw-proposituras
' Author: Christian Martin dos Santos <chrmsantos@gmail.com>

Option Explicit

'================================================================================
' CONSTANTS
'================================================================================

' System constants
Private Const version As String = "v1.0.0-Beta3"
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

' Dialog/MsgBox title constants (centralized UI titles)
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
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 21
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.7
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.19

' Performance batching constants (added to fix undefined symbol errors)
' When total paragraphs exceed OPTIMIZATION_THRESHOLD we process them in
' chunks of MAX_PARAGRAPH_BATCH_SIZE to balance speed and UI responsiveness.
Private Const OPTIMIZATION_THRESHOLD As Long = 400
Private Const MAX_PARAGRAPH_BATCH_SIZE As Long = 120

' Fixed application constants (replacing dynamic configuration)
Private Const MIN_WORD_VERSION As Double = 14#
Private Const HEADER_IMAGE_RELATIVE_PATH As String = "assets\stamp.png"

'================================================================================
' GLOBAL VARIABLES
'================================================================================
Private undoGroupEnabled As Boolean
Private formattingCancelled As Boolean
Private processingStartTime As Single ' Stores Timer() value at start of processing

'================================================================================
' UNIT CONVERSION UTILITIES
'================================================================================
' Word uses points (1 point = 1/72 inch). 1 inch = 2.54 cm. So cm = points * 2.54 / 72.
Private Function CmFromPoints(ByVal pts As Double) As Double
    CmFromPoints = (pts * 2.54#) / 72#
End Function

'================================================================================
' TIMING UTILITIES
'================================================================================
' Returns whole seconds elapsed since the stored processingStartTime.
' Safe if called before initialization (returns 0). Placed after UDT per VBA ordering rules.
Private Function ElapsedSeconds() As Long
    If processingStartTime <= 0 Then
        ElapsedSeconds = 0
    Else
        ElapsedSeconds = CLng(Timer - processingStartTime)
        If ElapsedSeconds < 0 Then ' Timer wraps at midnight
            ElapsedSeconds = ElapsedSeconds + 86400&
        End If
    End If
End Function

'================================================================================
' PERFORMANCE OPTIMIZATION SYSTEM
'================================================================================

Private Function InitializePerformanceOptimization() As Boolean
    On Error GoTo ErrorHandler
    
    InitializePerformanceOptimization = False
    
    ' Apply standard performance optimizations (always on)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    
    ' Word-specific optimizations
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

Private Function OptimizedFindReplace(findText As String, replaceText As String, Optional searchRange As Range = Nothing) As Long
    On Error GoTo ErrorHandler
    
    OptimizedFindReplace = 0
    
    ' Always use optimized bulk replace
    OptimizedFindReplace = BulkFindReplace(findText, replaceText, searchRange)
    
    Exit Function
    
ErrorHandler:
    OptimizedFindReplace = 0
End Function

Private Function BulkFindReplace(findText As String, replaceText As String, Optional searchRange As Range = Nothing) As Long
    On Error GoTo ErrorHandler
    
    BulkFindReplace = 0
    
    Dim targetRange As Range
    Set targetRange = IIf(searchRange Is Nothing, ActiveDocument.Content, searchRange)
    
    ' Optimization: use Word's native bulk operation method
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
        
    ' Execute all replacements at once
        BulkFindReplace = .Execute(Replace:=wdReplaceAll)
    End With
    
    Exit Function
    
ErrorHandler:
    BulkFindReplace = 0
End Function

Private Function StandardFindReplace(findText As String, replaceText As String, Optional searchRange As Range = Nothing) As Long
    On Error GoTo ErrorHandler
    
    StandardFindReplace = 0
    
    Dim targetRange As Range
    Set targetRange = IIf(searchRange Is Nothing, ActiveDocument.Content, searchRange)
    
    ' Standard compatible implementation
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
    StandardFindReplace = 0
End Function

Private Function OptimizedParagraphProcessing(processingFunction As String) As Boolean
    On Error GoTo ErrorHandler
    
    OptimizedParagraphProcessing = False
    
    ' Always use batch paragraph processing
    OptimizedParagraphProcessing = BatchProcessParagraphs(processingFunction)
    
    Exit Function
    
ErrorHandler:
    OptimizedParagraphProcessing = False
End Function

Private Function BatchProcessParagraphs(processingFunction As String) As Boolean
    On Error GoTo ErrorHandler
    
    BatchProcessParagraphs = False
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim paragraphCount As Long
    paragraphCount = doc.Paragraphs.count
    
    Dim batchSize As Long
    batchSize = IIf(paragraphCount > OPTIMIZATION_THRESHOLD, MAX_PARAGRAPH_BATCH_SIZE, paragraphCount)
    
    
    Dim i As Long
    For i = 1 To paragraphCount Step batchSize
        Dim endIndex As Long
        endIndex = IIf(i + batchSize - 1 > paragraphCount, paragraphCount, i + batchSize - 1)
        
        ' Process paragraph batch
        If Not ProcessParagraphBatch(i, endIndex, processingFunction) Then
            ' Batch error
            Exit Function
        End If
        
        ' Optionally yield to UI during long batches
        If i Mod (batchSize * 5) = 0 Then DoEvents
    Next i
    
    BatchProcessParagraphs = True
    Exit Function
    
ErrorHandler:
    BatchProcessParagraphs = False
End Function

Private Function StandardProcessParagraphs(processingFunction As String) As Boolean
    On Error GoTo ErrorHandler
    
    StandardProcessParagraphs = False
    
    ' Standard implementation - process paragraph by paragraph
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim para As Paragraph
    For Each para In doc.Paragraphs
    ' Apply specific processing function
        Select Case processingFunction
            Case "FORMAT"
                Call FormatParagraph(para)
            Case "CLEAN"
                Call CleanParagraph(para)
        End Select
    Next para
    
    StandardProcessParagraphs = True
    Exit Function
    
ErrorHandler:
    StandardProcessParagraphs = False
End Function

Private Function ProcessParagraphBatch(startIndex As Long, endIndex As Long, processingFunction As String) As Boolean
    On Error GoTo ErrorHandler
    
    ProcessParagraphBatch = False
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim i As Long
    For i = startIndex To endIndex
        If i <= doc.Paragraphs.count Then
            Dim para As Paragraph
            Set para = doc.Paragraphs(i)
            
            ' Apply specific processing function
            Select Case processingFunction
                Case "FORMAT"
                    Call FormatParagraph(para)
                Case "CLEAN"
                    Call CleanParagraph(para)
            End Select
        End If
    Next i
    
    ProcessParagraphBatch = True
    Exit Function
    
ErrorHandler:
    ProcessParagraphBatch = False
End Function

Private Sub FormatParagraph(para As Paragraph)
    On Error Resume Next
    ' Basic normalization: remove leading/trailing spaces in paragraph text (without touching internal spacing)
    ' Operates only on non-empty paragraphs of plain text (skips those containing tables or shapes inline)
    If para Is Nothing Then Exit Sub
    If para.Range.Tables.Count > 0 Then Exit Sub
    Dim txt As String
    txt = para.Range.Text
    ' Word paragraphs end with vbCr; preserve the final terminator
    If Len(txt) > 1 Then
        Dim body As String
        body = Left$(txt, Len(txt) - 1)
        body = Trim$(body)
        para.Range.Text = body & vbCr
    End If
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
    s = r.Text
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
        r.Text = tmp
    End If
End Sub



Private Sub ForceGarbageCollection()
    On Error Resume Next
    
    ' Yield to UI to keep Word responsive during long operations
    DoEvents
End Sub

'================================================================================
' MAIN ENTRY POINT
'================================================================================
Public Sub StandardizeDocumentMain()
    On Error GoTo CriticalErrorHandler
    
    ' ========================================
    ' INITIALIZATION AND CONFIG LOAD
    ' ========================================
    
    processingStartTime = Timer
    formattingCancelled = False
    
    
    
    ' ========================================
    ' PRELIMINARY VALIDATIONS BASED ON CONFIGURATION
    ' ========================================
    
    ' Word version validation (always on)
    If Not CheckWordVersion() Then
        Application.StatusBar = "Error: Word version not supported (minimum: Word " & MIN_WORD_VERSION & ")"
        Dim verMsg As String
        verMsg = ReplacePlaceholders(MSG_ERR_VERSION, _
                    "MIN", CStr(MIN_WORD_VERSION), _
                    "CUR", CStr(Application.Version))
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
    
    ' Ensure the document is editable (not Protected View, not read-only, not marked as Final)
    If Not EnsureDocumentEditable(doc) Then
        Application.StatusBar = "Document is not editable - operation cancelled"
        GoTo CleanUp
    End If
    
    ' ========================================
    ' PERFORMANCE OPTIMIZATION INITIALIZATION
    ' ========================================
    
    If Not InitializePerformanceOptimization() Then
        ' Continue execution even if optimizations fail
    End If
        
    ' Configure undo group
    StartUndoGroup "Document Standardization - " & doc.Name
    
    ' Configure application state
    If Not SetAppState(False, "Formatting document...") Then
        ' Ignore failure
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
    
    
    Application.StatusBar = "Processing document structure..."

    If Not PreviousFormatting(doc) Then
        GoTo CleanUp
    End If

    

    If formattingCancelled Then
        GoTo CleanUp
    End If

    Application.StatusBar = "Document standardized successfully!"

CleanUp:
    ' Restore performance settings
    If Not RestorePerformanceSettings() Then
        ' Ignore failure
    End If
    
    SafeCleanup
    CleanupImageProtection ' Cleanup image protection variables
    
    
    If Not SetAppState(True, "Document standardized successfully!") Then
        ' Ignore failure
    End If
    
    ' Final status / elapsed time (non-blocking)
    Dim elapsedSec As Long
    elapsedSec = CLng(Timer - processingStartTime)
    Application.StatusBar = "Chainsaw: concluído em " & ElapsedSeconds() & "s"
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "CRITICAL ERROR #" & Err.Number & ": " & Err.Description & _
              " in " & Err.Source & " (Line: " & Erl & ")"
    
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
    curVer = CDbl(Val(Application.Version))

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
    If doc.ProtectionType <> wdNoProtection Or doc.ReadOnly Then
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
    If doc.ProtectionType = wdNoProtection And Not doc.ReadOnly Then
        EnsureDocumentEditable = True
    End If
    Exit Function

ErrorHandler:
    EnsureDocumentEditable = False
End Function

'================================================================================
' DOCUMENT INTEGRITY VALIDATION
'================================================================================

'================================================================================
' EMERGENCY RECOVERY (stub)
'================================================================================
' Attempts minimal recovery actions after a critical failure.
' In the simplified build this only restores basic app state safely.
Private Sub EmergencyRecovery()
    On Error Resume Next
    ' Attempt to re-enable screen updating / alerts if disabled
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    ' Optionally could add a forced document save attempt here if desired.
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
    isProtected = (doc.ProtectionType <> wdNoProtection)
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
    SafeGetCharacterCount = Len(targetRange.Text)
    Exit Function
    
ErrorHandler:
          
End Function

Private Function SafeSetFont(targetRange As Range, fontName As String, fontSize As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' Apply font formatting safely
    With targetRange.Font
        If fontName <> "" Then .Name = fontName
        If fontSize > 0 Then .Size = fontSize
        .Color = wdColorAutomatic
    End With
    
    SafeSetFont = True
    Exit Function
ErrorHandler:
    SafeSetFont = False
End Function

Private Function SafeSetParagraphFormat(para As Paragraph, alignment As Long, leftIndent As Single, firstLineIndent As Single) As Boolean
    On Error GoTo ErrorHandler
    
    With para.Format
        If alignment >= 0 Then .Alignment = alignment
        If leftIndent >= 0 Then .LeftIndent = leftIndent
        If firstLineIndent >= 0 Then .FirstLineIndent = firstLineIndent
    End With
    SafeSetParagraphFormat = True
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
    .Text = findText
    .Replacement.Text = replaceText
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
    SafeGetLastCharacter = rng.Characters(charCount).Text
    Else
        SafeGetLastCharacter = ""
    End If
    Exit Function
    
ErrorHandler:
    ' Alternative method using Right()
    On Error GoTo FinalFallback
    SafeGetLastCharacter = Right(rng.Text, 1)
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
' UTILITY: GET PROTECTION TYPE
'================================================================================
Private Function GetProtectionType(doc As Document) As String
    On Error Resume Next
    
    Select Case doc.ProtectionType
        Case wdNoProtection: GetProtectionType = "No protection"
        Case 1: GetProtectionType = "Tracked changes protection"
        Case 2: GetProtectionType = "Comments protection"
        Case 3: GetProtectionType = "Forms protection"
        Case 4: GetProtectionType = "Read-only protection"
    Case Else: GetProtectionType = "Unknown type (" & doc.ProtectionType & ")"
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

    If doc.ProtectionType <> wdNoProtection Then
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
    
    ' Justificativa/Anexo formatting
    Call FormatJustificativaAnexoParagraphs(doc)

    ' Hyphenation and watermark
    Call EnableHyphenation(doc)
    Call RemoveWatermark(doc)

    ' Insert header image (always enabled)
    InsertHeaderstamp doc
    
    ' Insert page numbers in footer (restored feature)
    Call InsertFooterstamp(doc)
    
    ' Final spacing and separation controls
    Call CleanMultipleSpaces(doc)
    Call LimitSequentialEmptyLines(doc)
    Call EnsureParagraphSeparation(doc)
    Call EnsureSecondParagraphBlankLines(doc)
    Call FormatJustificativaAnexoParagraphs(doc)
    
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
    needsFontFormatting = (paraFont.Name <> STANDARD_FONT) Or _
                 (paraFont.Size <> STANDARD_FONT_SIZE) Or _
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
            paraFullText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            
            ' Check if it is the first paragraph with text (title) - optimized
            If i <= 3 And para.Format.Alignment = wdAlignParagraphCenter And paraFullText <> "" Then
                isTitle = True
            End If
                     
            ' Check if it is a special paragraph - optimized
            Dim cleanParaText As String
            cleanParaText = paraFullText
            ' Remove ending punctuation for analysis
            Do While Len(cleanParaText) >  0 And (Right(cleanParaText, 1) = "." Or Right(cleanParaText, 1) = "," Or Right(cleanParaText, 1) = ":" Or Right(cleanParaText, 1) = ";")
                cleanParaText = Left(cleanParaText, Len(cleanParaText) - 1)
            Loop
            cleanParaText = Trim(LCase(cleanParaText))

            If cleanParaText = "justificativa:" Or IsAnexoPattern(cleanParaText) Then
                isSpecialParagraph = True
            End If
            
            
            If i < doc.Paragraphs.count Then
                Dim nextPara As Paragraph
                Set nextPara = doc.Paragraphs(i + 1)
                If Not HasVisualContent(nextPara) Then
                    Dim nextParaText As String
                    nextParaText = Trim(Replace(Replace(nextPara.Range.Text, vbCr, ""), vbLf, ""))
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
                        .Name = STANDARD_FONT
                        .Size = STANDARD_FONT_SIZE
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
    
    If charCount > 0 Then ' Safety check
        For j = 1 To charCount
            Set charRange = para.Range.Characters(j)
            If charRange.InlineShapes.count = 0 Then
                With charRange.Font
                    ' Apply font formatting if specified
                    If fontName <> "" Then .Name = fontName
                    If fontSize > 0 Then .Size = fontSize
                    If fontColor >= 0 Then .Color = fontColor
                    
                    ' Remove special formats if requested
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
    cleanText = para.Range.Text
        
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
        If cleanText <> para.Range.Text And Not hasInlineImage Then
            para.Range.Text = cleanText
        End If

    paraText = Trim(LCase(Replace(Replace(Replace(para.Range.Text, ".", ""), ",", ""), ";", "")))
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
    paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
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
            .LeftIndent = CentimetersToPoints(9)      ' 9 cm left indent
            .FirstLineIndent = 0                      ' No first-line indent
            .RightIndent = 0                          ' No right indent
            .Alignment = wdAlignParagraphJustify      ' Justified
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
    paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
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
    paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
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
    paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
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
        
    ' If it has images, just log (but do not skip formatting)
        If HasVisualContent(para) Then
        Else
        End If
    Else
    End If
    
    EnsureSecondParagraphBlankLines = True
    Exit Function

ErrorHandler:
    EnsureSecondParagraphBlankLines = False
End Function

'================================================================================
' ENABLE HYPHENATION
'================================================================================
Private Function EnableHyphenation(doc As Document) As Boolean
    On Error GoTo ErrHandler
    ' Real implementation: enable automatic hyphenation for the active document.
    ' Keeps previous compatibility mode assumptions minimal.
    doc.Hyphenation = True
    EnableHyphenation = True
    Exit Function
ErrHandler:
    ' Fail-soft: do not abort formatting sequence if not supported.
    EnableHyphenation = True ' treat as success to avoid pipeline interruption
End Function

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
    paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
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
    ' 1st paragraph formatting: uppercase, bold, and underlined
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
            .Alignment = wdAlignParagraphCenter       ' Centered
            .LeftIndent = 0                           ' No left indent
            .FirstLineIndent = 0                      ' No first-line indent
            .RightIndent = 0                          ' No right indent
        End With
    End If
    
    Exit Function
    
ErrorHandler:
    FormatFirstParagraph = False
End Function

'================================================================================
' REMOVE WATERMARK
'================================================================================
Private Function RemoveWatermark(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim shp As Shape
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
    
    
    RemoveWatermark = True
    Exit Function
    
ErrorHandler:
    RemoveWatermark = False
End Function

'================================================================================
' INSERT HEADER IMAGE
'================================================================================
Private Function InsertHeaderstamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim header As HeaderFooter
    Dim imgFile As String
    Dim username As String
    Dim imgWidth As Single
    Dim imgHeight As Single
    Dim shp As Shape
    Dim imgFound As Boolean
    Dim sectionsProcessed As Long

    ' Resolve image file path (no extra fallbacks): use fixed relative path; if relative, resolve from document folder
    imgFile = HEADER_IMAGE_RELATIVE_PATH
    If InStr(1, imgFile, ":", vbTextCompare) = 0 And Left(imgFile, 2) <> "\\" Then
        Dim baseFolder As String
        On Error Resume Next
        baseFolder = IIf(doc.Path <> "", doc.Path, CurDir$)
        On Error GoTo ErrorHandler
        If Right(baseFolder, 1) <> "\" Then baseFolder = baseFolder & "\"
        imgFile = baseFolder & imgFile
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
                .Font.Size = FOOTER_FONT_SIZE
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Fields.Update
            End With
            
            sectionsProcessed = sectionsProcessed + 1
        End If
    Next sec

    
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
 

'================================================================================
' CLEAN DOCUMENT STRUCTURE - FEATURES 2, 6, 7
'================================================================================
Private Function CleanDocumentStructure(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim i As Long
    Dim firstTextParaIndex As Long
    Dim emptyLinesRemoved As Long
    Dim leadingSpacesRemoved As Long
    Dim paraCount As Long
    
    ' Cache the total paragraph count
    paraCount = doc.Paragraphs.count
    
    ' OPTIMIZED: Feature 2 - Remove blank lines above the title
    ' Optimized search for the first paragraph with text
    firstTextParaIndex = -1
    For i = 1 To paraCount
    If i > doc.Paragraphs.count Then Exit For ' Dynamic protection
        
        Set para = doc.Paragraphs(i)
        Dim paraTextCheck As String
    paraTextCheck = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
    ' Find the first paragraph with real text
        If paraTextCheck <> "" Then
            firstTextParaIndex = i
            Exit For
        End If
        
    ' Protection against very large documents
    If i > 50 Then Exit For ' Limit search to the first 50 paragraphs
    Next i
    
    ' OPTIMIZED: Remove empty lines BEFORE the first text in a single pass
    If firstTextParaIndex > 1 Then
    ' Process backwards to avoid index issues
        For i = firstTextParaIndex - 1 To 1 Step -1
            If i > doc.Paragraphs.count Or i < 1 Then Exit For ' Dynamic protection
            
            Set para = doc.Paragraphs(i)
            Dim paraTextEmpty As String
            paraTextEmpty = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            
            ' OPTIMIZED: Visual check only if necessary
            If paraTextEmpty = "" Then
                If Not HasVisualContent(para) Then
                    para.Range.Delete
                    emptyLinesRemoved = emptyLinesRemoved + 1
                    ' Update cache after removal
                    paraCount = paraCount - 1
                End If
            End If
        Next i
    End If
    
    ' SUPER OPTIMIZED: Feature 7 - Remove leading spaces with Find/Replace
    ' Use Find/Replace which is much faster than looping through paragraphs
    Dim rng As Range
    Set rng = doc.Range
    
    ' Remove spaces at the start of lines using Find/Replace
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        
    ' Remove spaces/tabs at the start of lines using simple Find/Replace
    .Text = "^p "  ' Paragraph break followed by space
        .Replacement.Text = "^p"
        
        Do While .Execute(Replace:=True)
            leadingSpacesRemoved = leadingSpacesRemoved + 1
            ' Protection against infinite loop
            If leadingSpacesRemoved > 1000 Then Exit Do
        Loop
        
    ' Remove tabs at the start of lines
    .Text = "^p^t"  ' Paragraph break followed by tab
        .Replacement.Text = "^p"
        
        Do While .Execute(Replace:=True)
            leadingSpacesRemoved = leadingSpacesRemoved + 1
            If leadingSpacesRemoved > 1000 Then Exit Do
        Loop
    End With
    
    ' Second pass for spaces at the absolute start of the document (no preceding ^p)
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
    .MatchWildcards = False  ' Do not use wildcards in this section
        
    ' Position at the start of the document
        rng.Start = 0
        rng.End = 1
        
    ' Remove spaces/tabs at the absolute start of the document
    If rng.Text = " " Or rng.Text = vbTab Then
            ' Expand the range to capture all leading spaces using safe method
            Do While rng.End <= doc.Range.End And (SafeGetLastCharacter(rng) = " " Or SafeGetLastCharacter(rng) = vbTab)
                rng.End = rng.End + 1
                leadingSpacesRemoved = leadingSpacesRemoved + 1
                If leadingSpacesRemoved > 100 Then Exit Do ' Protection
            Loop
            
            If rng.Start < rng.End - 1 Then
                rng.Delete
            End If
        End If
    End With
    
    ' Simplified log only if there was significant cleanup
    If emptyLinesRemoved > 0 Then
    End If
    
    CleanDocumentStructure = True
    Exit Function

ErrorHandler:
    CleanDocumentStructure = False
End Function

'================================================================================
' SAFE CHECK FOR VISUAL CONTENT
'================================================================================
Private Function HasVisualContent(para As Paragraph) As Boolean
    ' Use the safe function implemented for full compatibility
    HasVisualContent = SafeHasVisualContent(para)
End Function

'================================================================================
' VALIDATE PROPOSITION TYPE
'================================================================================
Private Function ValidatePropositionType(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim firstPara As Paragraph
    Dim firstWord As String
    Dim paraText As String
    Dim i As Long
    Dim userResponse As VbMsgBoxResult
    
    ' Find the first non-empty paragraph
    For i = 1 To doc.Paragraphs.count
        Set firstPara = doc.Paragraphs(i)
    paraText = Trim(Replace(Replace(firstPara.Range.Text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then
            Exit For
        End If
    Next i
    
    If paraText = "" Then
        ValidatePropositionType = True
        Exit Function
    End If
    
    ' Extract the first word
    Dim words() As String
    words = Split(paraText, " ")
    If UBound(words) >= 0 Then
        firstWord = LCase(Trim(words(0)))
    End If
    
    ' Check if it's one of the expected proposition types (Portuguese terms)
    If firstWord = "indica��o" Or firstWord = "requerimento" Or firstWord = "mo��o" Then
        ValidatePropositionType = True
    Else
        ' Not a standard proposition document � ask the user for confirmation
        Application.StatusBar = "Waiting for user confirmation about document type..."
        
    ' Build a detailed message for the user
        Dim confirmationMessage As String
    confirmationMessage = ReplacePlaceholders(MSG_DOC_TYPE_WARNING, _
                         "FIRSTWORD", UCase(firstWord), _
                         "DOCSTART", Left(paraText, 150))
        userResponse = MsgBox(NormalizeForUI(confirmationMessage), vbYesNo + vbQuestion + vbDefaultButton2, _
                 NormalizeForUI(TITLE_DOC_TYPE))
        
        If userResponse = vbYes Then
            Application.StatusBar = "Processing non-standard document as requested..."
            ValidatePropositionType = True
        Else
            Application.StatusBar = "Processing cancelled by user"
            
            ' Final cancellation message
            MsgBox NormalizeForUI(MSG_PROCESSING_CANCELLED), _
                vbInformation, NormalizeForUI(TITLE_OPERATION_CANCELLED)
            
            ValidatePropositionType = False
        End If
    End If
    
    Exit Function

ErrorHandler:
    ValidatePropositionType = False
End Function

'================================================================================
' FORMAT DOCUMENT TITLE
'================================================================================
Private Function FormatDocumentTitle(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim firstPara As Paragraph
    Dim paraText As String
    Dim words() As String
    Dim i As Long
    Dim newText As String
    
    ' Find the first non-empty paragraph (after skipping blank lines)
    For i = 1 To doc.Paragraphs.count
        Set firstPara = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(firstPara.Range.Text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then
            Exit For
        End If
    Next i
    
    If paraText = "" Then
        FormatDocumentTitle = True
        Exit Function
    End If
    
    ' Remove trailing period if present
    If Right(paraText, 1) = "." Then
        paraText = Left(paraText, Len(paraText) - 1)
    End If
    
    ' Check if it's a proposition (to apply $NUMERO$/$ANO$ substitution)
    Dim isProposition As Boolean
    Dim firstWord As String
    
    words = Split(paraText, " ")
    If UBound(words) >= 0 Then
        firstWord = LCase(Trim(words(0)))
        If firstWord = "indica��o" Or firstWord = "requerimento" Or firstWord = "mo��o" Then
            isProposition = True
        End If
    End If
    
    ' If proposition, replace the last word with $NUMERO$/$ANO$
    If isProposition And UBound(words) >= 0 Then
    ' Rebuild the text replacing the last word
        newText = ""
        For i = 0 To UBound(words) - 1
            If i > 0 Then newText = newText & " "
            newText = newText & words(i)
        Next i
        
        ' Add $NUMERO$/$ANO$ instead of the last word
        If newText <> "" Then newText = newText & " "
        newText = newText & "$NUMERO$/$ANO$"
    Else
        ' Not a proposition: keep the original text
        newText = paraText
    End If
    
    ' Always apply title formatting: uppercase, bold, underline
    firstPara.Range.Text = UCase(newText) & vbCrLf
    
    ' Full title formatting (first line)
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
        .SpaceAfter = 6  ' Small space after the title
    End With
    
    If isProposition Then
    Else
    End If
    
    FormatDocumentTitle = True
    Exit Function

ErrorHandler:
    FormatDocumentTitle = False
End Function

'================================================================================
' FORMAT "CONSIDERANDO" PARAGRAPHS - OPTIMIZED AND SIMPLIFIED
'================================================================================
Private Function FormatConsiderandoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim rawText As String
    Dim textNoCrLf As String
    Dim i As Long
    Dim totalFormatted As Long
    Dim startIdx As Long
    Dim n As Long
    Dim ch As String
    Dim code As Long
    Dim allowedPrefix As String
    
    ' Characters we can ignore before "considerando" at paragraph start
    ' spaces/tabs, quotes, dashes, hyphen, parentheses, and a set of invisible/control chars
    ' Note: we keep allowedPrefix defined above for readability, but detection relies on code-point checks here
    allowedPrefix = " " ' (kept for documentation only)
    
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
    rawText = para.Range.Text
        textNoCrLf = Replace(Replace(rawText, vbCr, ""), vbLf, "")
        
        If Len(textNoCrLf) >= 12 Then
            ' Find first non-prefix character index, skipping spaces, punctuation, and invisible/control marks
            startIdx = 1
            For n = 1 To Len(textNoCrLf)
                ch = Mid$(textNoCrLf, n, 1)
                code = AscW(ch)

                    startIdx = n
                    Exit For
            Next n
            
            If startIdx + 11 <= Len(textNoCrLf) Then
                If LCase$(Mid$(textNoCrLf, startIdx, 12)) = "considerando" Then
                    Dim rng As Range
                    Set rng = para.Range.Duplicate
                    rng.SetRange rng.Start + (startIdx - 1), rng.Start + (startIdx - 1) + 12
                    
                    ' Replace token and apply bold preserving following spacing/punctuation
                    rng.Text = "CONSIDERANDO"
                    rng.Font.Bold = True
                    totalFormatted = totalFormatted + 1
                End If
            End If
        End If
    Next i
    
    If totalFormatted > 0 Then
    Else
    End If
    
    FormatConsiderandoParagraphs = True
    Exit Function

ErrorHandler:
    FormatConsiderandoParagraphs = False
End Function

'================================================================================
' LOG STRING NORMALIZATION - avoid encoding issues in log files
'================================================================================
Private Function NormalizeForLog(ByVal s As String) As String
    On Error Resume Next
    Dim i As Long, ch As String, code As Long
    Dim out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        ' Keep ASCII printable and common punctuation; replace others with '?'
        If code >= 32 And code <= 126 Then
            out = out & ch
        ElseIf ch = vbCr Or ch = vbLf Or ch = vbTab Then
            out = out & ch
        Else
            out = out & "?"
        End If
    Next i
    NormalizeForLog = out
End Function

'================================================================================
' UI STRING NORMALIZATION - produce ASCII-safe text for MsgBox dialogs
'================================================================================
Private Function NormalizeForUI(ByVal s As String) As String
    On Error Resume Next
    Dim i As Long, ch As String, code As Long
    Dim out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case code
            Case 192 To 197, 224 To 229: out = out & "a"   ' ������������
            Case 199: out = out & "C"                      ' �
            Case 231: out = out & "c"                      ' �
            Case 200 To 203, 232 To 235: out = out & "e"
            Case 204 To 207, 236 To 239: out = out & "i"
            Case 210 To 214, 242 To 246: out = out & "o"
            Case 217 To 220, 249 To 252: out = out & "u"
            Case 209: out = out & "N"                      ' �
            Case 241: out = out & "n"                      ' �
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
    dOesteVariants(1) = "d�O"   ' Acute accent
    dOesteVariants(2) = "d`O"   ' Grave accent
    dOesteVariants(3) = "d" & Chr(8220) & "O"   ' Left curly quote
    dOesteVariants(4) = "d'o"   ' Lowercase
    dOesteVariants(5) = "d�o"
    dOesteVariants(6) = "d`o"
    dOesteVariants(7) = "d" & Chr(8220) & "o"
    dOesteVariants(8) = "D'O"   ' Uppercase D
    dOesteVariants(9) = "D�O"
    dOesteVariants(10) = "D`O"
    dOesteVariants(11) = "D" & Chr(8220) & "O"
    dOesteVariants(12) = "D'o"
    dOesteVariants(13) = "D�o"
    dOesteVariants(14) = "D`o"
    dOesteVariants(15) = "D" & Chr(8220) & "o"
    
    For i = 0 To UBound(dOesteVariants)
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = dOesteVariants(i) & "este"
            .Replacement.Text = "d'Oeste"
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
    
    
    
    ' Feature 12: Replace isolated hyphens/en dashes with em dash (�)
    ' Normalizes hyphens (-) and en dashes (�) surrounded by spaces into em dashes (�)
    Set rng = doc.Range
    Dim dashVariants() As String
    ReDim dashVariants(0 To 2)
    
    ' Define dash types to replace when surrounded by spaces
    dashVariants(0) = " - "     ' Hyphen
    dashVariants(1) = " � "     ' En dash
    dashVariants(2) = " � "     ' Em dash (normalize)
    
    ' Replace all types with em dash
    For i = 0 To UBound(dashVariants)
    ' Only if not already an em dash
        If dashVariants(i) <> " � " Then
            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = dashVariants(i)
                .Replacement.Text = " � "    ' Em dash (travess�o) com espa�os
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
    lineStartDashVariants(1) = "^p� "   ' En dash at line start
    
    For i = 0 To UBound(lineStartDashVariants)
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = lineStartDashVariants(i)
            .Replacement.Text = "^p� "    ' Em dash at line start
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
    lineEndDashVariants(1) = " �^p"   ' En dash at line end
    
    For i = 0 To UBound(lineEndDashVariants)
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = lineEndDashVariants(i)
            .Replacement.Text = " �^p"    ' Em dash at line end
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
    .Text = "^l"  ' Manual line break
    .Replacement.Text = " "  ' Replace with space
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
    .Text = Chr(11)  ' Manual line break (VT)
    .Replacement.Text = " "  ' Replace with space
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
    .Text = Chr(10)  ' Line Feed (LF)
    .Replacement.Text = " "  ' Replace with space
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

'================================================================================
' APPLY SPECIFIC PARAGRAPH REPLACEMENTS
'================================================================================
Private Function ApplySpecificParagraphReplacements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Applying specific replacements (per-paragraph and global)..."
    
    Dim replacementCount As Long
    Dim secondParaIndex As Long
    Dim thirdParaIndex As Long
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    
    replacementCount = 0
    
    ' Find the 2nd and 3rd paragraphs with content (skip empty ones)
    actualParaIndex = 0
    secondParaIndex = 0
    thirdParaIndex = 0
    
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
    paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
    ' Count as an actual paragraph if it has content
        If paraText <> "" Then
            actualParaIndex = actualParaIndex + 1
            
            If actualParaIndex = 2 Then
                secondParaIndex = i
            ElseIf actualParaIndex = 3 Then
                thirdParaIndex = i
                Exit For ' Found both needed paragraphs
            End If
        End If
        
    ' Safety: protect against very large documents
        If i > 50 Then Exit For
    Next i
    
    ' REQUIREMENT 1: If the 2nd paragraph starts with "Sugiro " or "Sugere ", replace accordingly
    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(secondParaIndex)
    paraText = para.Range.Text
        
        ' Skip leading spaces/tabs before checking the start token
        Dim idxStart As Long
        idxStart = 1
        Do While idxStart <= Len(paraText) And (Mid$(paraText, idxStart, 1) = " " Or Mid$(paraText, idxStart, 1) = vbTab)
            idxStart = idxStart + 1
        Loop
        
        If Len(paraText) >= idxStart + 6 Then
            Dim token As String
            token = Mid$(paraText, idxStart, 7) ' includes trailing space
            If token = "Sugiro " Then
                Dim r1 As Range
                Set r1 = para.Range.Duplicate
                r1.SetRange r1.Start + (idxStart - 1), r1.Start + (idxStart - 1) + 7
                r1.Text = "Requeiro "
                replacementCount = replacementCount + 1
            ElseIf token = "Sugere " Then
                Dim r2 As Range
                Set r2 = para.Range.Duplicate
                r2.SetRange r2.Start + (idxStart - 1), r2.Start + (idxStart - 1) + 7
                r2.Text = "Indica "
                replacementCount = replacementCount + 1
            End If
        End If
    End If
    
    ' REQUIREMENTS 2 and 3: Replacements in the 3rd paragraph
    If thirdParaIndex > 0 And thirdParaIndex <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(thirdParaIndex)
    paraText = para.Range.Text
        Dim originalText As String
        originalText = paraText
        
    ' REQUIREMENT 2: Replace " sugerir " with " indicar " anywhere
        If InStr(paraText, " sugerir ") > 0 Then
            paraText = Replace(paraText, " sugerir ", " indicar ")
            replacementCount = replacementCount + 1
        End If
        
    ' REQUIREMENT 3: Replace " Setor, " with " setor competente, "
        If InStr(paraText, " Setor, ") > 0 Then
            paraText = Replace(paraText, " Setor, ", " setor competente, ")
            replacementCount = replacementCount + 1
        End If
        
    ' Apply changes if replacements occurred
        If paraText <> originalText Then
            para.Range.Text = paraText
        End If
    End If
    
    ' GLOBAL REQUIREMENTS: Replacements across the whole document
    Dim rng As Range
    Set rng = doc.Range
    
    ' GLOBAL 1: Replace specific uppercase institutional phrase
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
    .Text = " A C�MARA MUNICIPAL DE SANTA B�RBARA D'OESTE, ESTADO DE S�O PAULO "
    .Replacement.Text = " a C�mara Municipal de Santa B�rbara d'Oeste, estado de S�o Paulo, "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
    .MatchCase = True  ' Case-sensitive for this specific replacement
        .MatchWholeWord = False
        .MatchWildcards = False
        
        Do While .Execute(Replace:=wdReplaceOne)
            replacementCount = replacementCount + 1
            rng.Collapse wdCollapseEnd
        Loop
    End With
    
    ' GLOBAL 2: Uppercase specific words
    Dim wordsToUppercase() As String
    Dim j As Long
    
    ' Define array with all word variations to be uppercased
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
    
    ' Apply uppercase conversion for each word
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
            .MatchCase = True  ' Case-sensitive to detect the exact variation
            .MatchWholeWord = True  ' Whole word only
            .MatchWildcards = False
            
            Do While .Execute(Replace:=wdReplaceOne)
                replacementCount = replacementCount + 1
                If replacementCount <= 20 Then  ' Log only the first cases for performance
                End If
                rng.Collapse wdCollapseEnd
            Loop
        End With
    Next j
    
    ApplySpecificParagraphReplacements = True
    Exit Function

ErrorHandler:
    ApplySpecificParagraphReplacements = False
End Function

'================================================================================
' VALIDATE CONTENT CONSISTENCY - VALIDA��O DE CONSIST�NCIA ENTRE EMENTA E TEOR
'================================================================================
Private Function ValidateContentConsistency(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Validating consistency between summary and body..."
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim secondParaIndex As Long
    Dim secondParaText As String
    Dim restOfDocumentText As String
    
    ' Find the 2nd paragraph with content (summary)
    actualParaIndex = 0
    secondParaIndex = 0
    
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
    paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
    ' Count as a real paragraph if it has content
        If paraText <> "" Then
            actualParaIndex = actualParaIndex + 1
            
            If actualParaIndex = 2 Then
                secondParaIndex = i
                secondParaText = paraText
                Exit For
            End If
        End If
        
    ' Safety guard for very large documents
        If i > 50 Then Exit For
    Next i
    
    ' If 2nd paragraph not found, skip validation
    If secondParaIndex = 0 Or secondParaText = "" Then
        ValidateContentConsistency = True
        Exit Function
    End If
    
    ' Collect the remainder of the document text (from 3rd paragraph onward)
    restOfDocumentText = ""
    actualParaIndex = 0
    
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
    paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
        If paraText <> "" Then
            actualParaIndex = actualParaIndex + 1
            
            ' Collect text from the 3rd paragraph onward
            If actualParaIndex >= 3 Then
                restOfDocumentText = restOfDocumentText & " " & paraText
            End If
        End If
    Next i
    
    ' If there isn't enough content to compare, skip validation
    If restOfDocumentText = "" Then
        ValidateContentConsistency = True
        Exit Function
    End If
    
    ' Analyze consistency between the 2nd paragraph and the rest
    Dim commonWordsCount As Long
    commonWordsCount = CountCommonWords(secondParaText, restOfDocumentText)
    
    
    ' If fewer than 2 common words, alert about possible inconsistency
    If commonWordsCount < 2 Then
    ' Show a warning to the user
        Dim inconsistencyMessage As String
        Dim userResponse As VbMsgBoxResult
    inconsistencyMessage = ReplacePlaceholders(MSG_INCONSISTENCY_WARNING, _
                           "Ementa", Left(secondParaText, 200), _
                           "COMMON", CStr(commonWordsCount))
        userResponse = MsgBox(NormalizeForUI(inconsistencyMessage), vbYesNo + vbExclamation + vbDefaultButton2, _
                 NormalizeForUI(TITLE_CONSISTENCY))
        
        If userResponse = vbNo Then
            Application.StatusBar = "Formatting stopped - inconsistency detected"
            ValidateContentConsistency = False
            Exit Function
        Else
        End If
    Else
    End If
    
    ValidateContentConsistency = True
    Exit Function

ErrorHandler:
    ValidateContentConsistency = False
End Function

'================================================================================
' COUNT COMMON WORDS - CONTA PALAVRAS COMUNS ENTRE DOIS TEXTOS
'================================================================================
Private Function CountCommonWords(text1 As String, text2 As String) As Long
    On Error GoTo ErrorHandler
    
    Dim words1() As String
    Dim words2() As String
    Dim i As Long, j As Long
    Dim commonCount As Long
    Dim word1 As String, word2 As String
    
    ' Clean and normalize texts
    text1 = CleanTextForComparison(text1)
    text2 = CleanTextForComparison(text2)
    
    ' Split into words
    words1 = Split(text1, " ")
    words2 = Split(text2, " ")
    
    commonCount = 0
    
    ' Compare each word of the first text with those of the second
    For i = 0 To UBound(words1)
        word1 = Trim(words1(i))
        
    ' Ignore very short words (<4 chars) or common words
        If Len(word1) >= 4 And Not IsCommonWord(word1) Then
            For j = 0 To UBound(words2)
                word2 = Trim(words2(j))
                
                ' If equal, count and break (avoid duplicates)
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
' CLEAN TEXT FOR COMPARISON - LIMPA TEXTO PARA COMPARA��O
'================================================================================
Private Function CleanTextForComparison(text As String) As String
    Dim cleanedText As String
    Dim i As Long
    Dim char As String
    
    ' Convert to lowercase
    cleanedText = LCase(text)
    
    ' Remove punctuation and special chars, keep only letters, numbers and spaces
    Dim result As String
    result = ""
    
    For i = 1 To Len(cleanedText)
        char = Mid(cleanedText, i, 1)
        
    ' Keep only letters, numbers and spaces
        If (char >= "a" And char <= "z") Or (char >= "0" And char <= "9") Or char = " " Then
            result = result & char
        Else
            ' Replace punctuation with space
            result = result & " "
        End If
    Next i
    
    ' Remove multiple spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    CleanTextForComparison = Trim(result)
End Function

'================================================================================
' IS COMMON WORD - VERIFICA SE � PALAVRA MUITO COMUM
'================================================================================
Private Function IsCommonWord(word As String) As Boolean
    ' List of very common Portuguese words to ignore in comparison
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
    commonWords(19) = "ap�s"
    commonWords(20) = "antes"
    commonWords(21) = "durante"
    commonWords(22) = "atrav�s"
    commonWords(23) = "mediante"
    commonWords(24) = "junto"
    commonWords(25) = "desde"
    commonWords(26) = "at�"
    commonWords(27) = "contra"
    commonWords(28) = "favor"
    commonWords(29) = "deve"
    commonWords(30) = "devem"
    commonWords(31) = "pode"
    commonWords(32) = "podem"
    commonWords(33) = "ser�"
    commonWords(34) = "ser�o"
    commonWords(35) = "est�"
    commonWords(36) = "est�o"
    commonWords(37) = "foram"
    commonWords(38) = "sendo"
    commonWords(39) = "tendo"
    commonWords(40) = "onde"
    commonWords(41) = "quando"
    commonWords(42) = "como"
    commonWords(43) = "porque"
    commonWords(44) = "portanto"
    commonWords(45) = "assim"
    commonWords(46) = "ent�o"
    commonWords(47) = "ainda"
    commonWords(48) = "tamb�m"
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
' FORMAT JUSTIFICATIVA/ANEXO PARAGRAPHS - SPECIAL FORMATTING
'================================================================================
Private Function FormatJustificativaAnexoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim cleanText As String
    Dim i As Long
    Dim formattedCount As Long
    
    ' Iterate all document paragraphs
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        
    ' Don't process paragraphs with visual content
        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            
            ' Remove trailing punctuation for better analysis
            cleanText = paraText
            ' Remove trailing . , : ;
            Do While Len(cleanText) > 0 And (Right(cleanText, 1) = "." Or Right(cleanText, 1) = "," Or Right(cleanText, 1) = ":" Or Right(cleanText, 1) = ";")
                cleanText = Left(cleanText, Len(cleanText) - 1)
            Loop
            cleanText = Trim(LCase(cleanText))
            
            ' Format "Justificativa"
            If cleanText = "Justificativa" Then
                ' Apply specific formatting for "Justificativa"
                With para.Format
                    .LeftIndent = 0
                    .FirstLineIndent = 0
                    .Alignment = wdAlignParagraphCenter
                End With

                With para.Range.Font
                    ' Clear previous formatting explicitly if needed; set Bold
                    .Bold = True
                End With
                
                ' Normalize text keeping original trailing punctuation if present
                Dim originalEnd As String
                originalEnd = ""
                If Len(paraText) > Len(cleanText) Then
                    originalEnd = Right(paraText, Len(paraText) - Len(cleanText))
                End If
                para.Range.Text = "Justificativa" & originalEnd & vbCrLf
                
                formattedCount = formattedCount + 1
                
            
                
            ' Format variations of "Anexo" or "Anexos"
            ElseIf IsAnexoPattern(cleanText) Then
                ' Apply specific formatting for Anexo/Anexos
                With para.Format
                    .LeftIndent = 0
                    .FirstLineIndent = 0
                    .RightIndent = 0
                    .Alignment = wdAlignParagraphLeft
                End With

                With para.Range.Font
                    .Bold = True
                End With

                ' Normalize text keeping original trailing punctuation if present
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
                
                formattedCount = formattedCount + 1
                
            ' Format paragraphs starting with "Ante o exposto"
            ElseIf IsAnteOExpostoPattern(paraText) Then
                ' Apply bold formatting to the token
                With para.Range.Font
                    .Bold = True
                End With
                
                formattedCount = formattedCount + 1
            End If
        End If
    Next i
    
    FormatJustificativaAnexoParagraphs = True
    Exit Function

ErrorHandler:
    FormatJustificativaAnexoParagraphs = False
End Function

'================================================================================
' FORMAT NUMBERED PARAGRAPHS
'================================================================================
Private Function FormatNumberedParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim formattedCount As Long
    
    ' Iterate all paragraphs of the document
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        
    ' Skip paragraphs with visual content
        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            
            ' Check if the paragraph starts with a number followed by ., ) or space
            If IsNumberedParagraph(paraText) Then
                ' Apply numbered list formatting
                With para.Range.ListFormat
                    ' Remove existing list formatting first
                    .RemoveNumbers
                    
                    ' Aplica lista numerada
                    .ApplyNumberDefault
                End With
                
                ' Remove the manual number as the list will generate it
                Dim cleanedText As String
                cleanedText = RemoveManualNumber(paraText)
                
                ' Update the paragraph text
                para.Range.Text = cleanedText & vbCrLf
                
                formattedCount = formattedCount + 1
            End If
        End If
    Next i
    
    FormatNumberedParagraphs = True
    Exit Function

ErrorHandler:
    FormatNumberedParagraphs = False
End Function

'================================================================================
' HELPER FUNCTIONS FOR PATTERN DETECTION
'================================================================================
 

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
        If IsNumeric(numberPart) And val(numberPart) > 0 Then
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
        If IsNumeric(numberPart) And val(numberPart) > 0 Then
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
        If IsNumeric(numberPart) And val(numberPart) > 0 Then
            ' Verify there is substantive text after the number and punctuation
            If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then
                IsNumberedParagraph = True
                Exit Function
            End If
        End If
    End If
    
    ' Pattern 4: just a number followed by space (1 text, 2 text, ...)
    ' Stricter: must have space AND substantive text after the number
    If IsNumeric(firstToken) And val(firstToken) > 0 And spacePos > 0 Then
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
            If IsNumeric(numberPart) And val(numberPart) > 0 Then
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
           (IsNumeric(firstToken) And val(firstToken) > 0) Then
            
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
            .Text = "  "  ' Two spaces
            .Replacement.Text = " "  ' One space
            
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
    .Text = "  ^p"  ' 2 spaces followed by paragraph break
    .Replacement.Text = " ^p"  ' 1 space followed by paragraph break
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
    ' Second pass to ensure complete cleanup
    .Text = " ^p"  ' Space before break
        .Replacement.Text = "^p"
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
    ' Remove multiple spaces after breaks - iterative method
    .Text = "^p  "  ' Break followed by 2 spaces
    .Replacement.Text = "^p "  ' Break followed by 1 space
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
    .Text = "^t^t"  ' 2 tabs
    .Replacement.Text = "^t"  ' 1 tab
        Do While .Execute(Replace:=wdReplaceOne)
            spacesRemoved = spacesRemoved + 1
            rng.Collapse wdCollapseEnd
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
    ' Convert tabs to spaces
    .Text = "^t"
    .Replacement.Text = " "
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
    End With
    
    ' OPTIMIZATION 4: Final ultra-fast check for remaining double spaces
    Set rng = doc.Range
    With rng.Find
    .Text = "  "
    .Replacement.Text = " "
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
    Exit Function

ErrorHandler:
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
    .Text = "^p^p^p^p"  ' 4 breaks
    .Replacement.Text = "^p^p"  ' 2 breaks
        
        Do While .Execute(Replace:=True)
            linesRemoved = linesRemoved + 1
            totalReplaces = totalReplaces + 1
            If totalReplaces > 500 Then Exit Do
            If linesRemoved Mod 50 = 0 Then DoEvents
        Loop
        
    ' Convert 3 breaks -> 2 breaks
    .Text = "^p^p^p"  ' 3 breaks
    .Replacement.Text = "^p^p"  ' 2 breaks
        
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
    .Text = "^p^p^p"  ' 3 breaks
    .Replacement.Text = "^p^p"  ' 2 breaks
        
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
    .Text = "^p^p^p"  ' 3 breaks (2 blank lines + content)
    .Replacement.Text = "^p^p"  ' 2 breaks (1 blank line + content)
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
            paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            
            ' Check if the paragraph is empty
            If paraText = "" And Not HasVisualContent(para) Then
                emptyLineCount = emptyLineCount + 1
                
                ' If we already have more than 1 consecutive blank line, remove this one
                If emptyLineCount > 1 Then
                    para.Range.Delete
                    fallbackRemoved = fallbackRemoved + 1
                    linesRemoved = linesRemoved + 1
                    
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
    
    LimitSequentialEmptyLines = True
    Exit Function

ErrorHandler:
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
        
    paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
    nextParaText = Trim(Replace(Replace(nextPara.Range.Text, vbCr, ""), vbLf, ""))
        nextParaText = Trim(Replace(Replace(nextPara.Range.Text, vbCr, ""), vbLf, ""))
        
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
                insertRange.Text = vbCrLf
                
                insertedCount = insertedCount + 1
                
                ' Update paragraph references after insertion
                ' because indices may have changed
                On Error Resume Next
                Set para = doc.Paragraphs(i)
                Set nextPara = doc.Paragraphs(i + 2) ' +2 because we inserted a line
                On Error GoTo ErrorHandler
                
                ' Log only for the first cases or significant cases
                If insertedCount <= 10 Or insertedCount Mod 50 = 0 Then
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
            Exit For
        End If
    Next i
    
    EnsureParagraphSeparation = True
    Exit Function

ErrorHandler:
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
    
    ConfigureDocumentView = True
    Exit Function

ErrorHandler:
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
    
    ' Check if there are open documents
    If Application.Documents.count = 0 Then
    Application.StatusBar = "No documents open - closing Word"
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
        End If
        On Error GoTo CriticalErrorHandler
    Next i
    
    ' If no unsaved documents, close directly
    If unsavedDocs.count = 0 Then
    Application.StatusBar = "All documents saved - closing Word"
        Application.Quit SaveChanges:=wdDoNotSaveChanges
        Exit Sub
    End If
    
    ' Build detailed message about unsaved documents
    Dim message As String
    Dim docList As String
    
    For i = 1 To unsavedDocs.count
        docList = docList & "� " & unsavedDocs(i) & vbCrLf
    Next i
    
    message = "ATTENTION: There are " & unsavedDocs.count & " document(s) with unsaved changes:" & vbCrLf & vbCrLf
    message = message & docList & vbCrLf
    message = message & "Do you want to save all documents before exiting?" & vbCrLf & vbCrLf
    message = message & "� YES: Save all and close Word" & vbCrLf
    message = message & "� NO: Close without saving (you will LOSE changes)" & vbCrLf
    message = message & "� CANCEL: Cancel the operation"
    
    ' Present options to the user
    Application.StatusBar = "Waiting for user decision about unsaved documents..."
    
    Dim userChoice As VbMsgBoxResult
    userChoice = MsgBox(NormalizeForUI(message), vbYesNoCancel + vbExclamation + vbDefaultButton1, _
                        NormalizeForUI(SYSTEM_NAME & " - Save and Exit (" & unsavedDocs.count & " unsaved document(s))"))
    
    Select Case userChoice
        Case vbYes
            ' User chose to save all
            Application.StatusBar = "Saving all documents..."
            
            If SalvarTodosDocumentos() Then
                Application.StatusBar = "Documents saved successfully - closing Word"
                Application.Quit SaveChanges:=wdDoNotSaveChanges
            Else
          Application.StatusBar = "Error saving documents - operation cancelled"
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
                Application.Quit SaveChanges:=wdDoNotSaveChanges
            Else
          Application.StatusBar = "Operation cancelled by user"
          MsgBox NormalizeForUI(MSG_OPERATION_CANCELLED), _
              vbInformation, NormalizeForUI(TITLE_OPERATION_CANCELLED)
            End If
            
        Case vbCancel
         ' User cancelled
         Application.StatusBar = "Exit operation cancelled by user"
         MsgBox NormalizeForUI(MSG_OPERATION_CANCELLED), _
             vbInformation, NormalizeForUI(TITLE_OPERATION_CANCELLED)
    End Select
    
    Application.StatusBar = False
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "CRITICAL ERROR in Save and Exit operation #" & Err.Number & ": " & Err.Description
    
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
                    Else
                        errorCount = errorCount + 1
                    End If
                Else
                    errorCount = errorCount + 1
                End If
            End With
        Else
            ' Document already has a path, just save it
            doc.Save
            If Err.Number = 0 Then
                savedCount = savedCount + 1
            Else
                errorCount = errorCount + 1
            End If
        End If
        
        On Error GoTo ErrorHandler
    Next i
    
    ' Verify result
    If errorCount = 0 Then
        SalvarTodosDocumentos = True
    Else
        SalvarTodosDocumentos = False
    End If
    
    Exit Function

ErrorHandler:
    SalvarTodosDocumentos = False
End Function


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
 

'================================================================================
' CLEANUP IMAGE PROTECTION - Cleanup image protection variables
'================================================================================
Private Sub CleanupImageProtection()
    ' Stub: nothing to cleanup
End Sub

'================================================================================
' SAFE CLEANUP (stub)
'================================================================================
 
Private Sub SafeCleanup()
    On Error Resume Next
    
End Sub

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
    Dim shp As Shape
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
            inlineShp.Delete
            deletedCount = deletedCount + 1
        End If
    Next i
    
    DeleteHiddenVisualElements = True
    Exit Function

ErrorHandler:
    DeleteHiddenVisualElements = False
End Function

'================================================================================
' DELETE VISUAL ELEMENTS IN RANGE - Remove visual elements between paragraphs 1-4
'================================================================================
Private Function DeleteVisualElementsInFirstFourParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Removing visual elements between paragraphs 1-4..."
    
    If doc.Paragraphs.count < 1 Then
        DeleteVisualElementsInFirstFourParagraphs = True
        Exit Function
    End If
    
    If doc.Paragraphs.count < 4 Then
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
    
    
    ' Remove floating shapes anchored within the first 4 paragraphs' range
    Dim i As Long
    For i = doc.Shapes.count To 1 Step -1
    Dim shp As Shape
        Set shp = doc.Shapes(i)
        
    ' Check if the shape is anchored within the first 4 paragraphs' range
        On Error Resume Next
        Dim anchorPosition As Long
        anchorPosition = shp.Anchor.Start
        On Error GoTo ErrorHandler
        
        If anchorPosition >= startRange And anchorPosition <= endRange Then
            Dim paragraphNum As Long
            paragraphNum = GetParagraphNumber(doc, anchorPosition)
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
            inlineShp.Delete
            deletedCount = deletedCount + 1
        End If
    Next i
    
    DeleteVisualElementsInFirstFourParagraphs = True
    Exit Function

ErrorHandler:
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
    ' SELF-TEST / QUICK VALIDATION MACRO
    '================================================================================
    ' ChainsawSelfTest runs a minimal in-document validation after formatting to ensure
    ' that critical invariants hold. It does NOT modify substantive content; it only
    ' reports findings via MsgBox. This helps detect regressions if future edits are made.
    '
    ' Invariants Checked:
    '  1. Base font family is Arial and size 12 for body paragraphs.
    '  2. First page margins roughly match configured CM values (±0.1cm tolerance).
    '  3. Hyphenation is enabled.
    '  4. No more than one consecutive blank (truly empty) paragraph sequence.
    '  5. Footer contains page numbering field.
    '
    ' Returns: Displays a summary report. Use for manual QA after running PreviousFormatting.
    '
    Public Sub ChainsawSelfTest()
        On Error GoTo FailHard
        Dim doc As Document
        If Documents.Count = 0 Then
            MsgBox "No open document to test.", vbExclamation, SYSTEM_NAME & " - Self-Test"
            Exit Sub
        End If
        Set doc = ActiveDocument

        Dim issues As Collection
        Set issues = New Collection

        ' 1. Font invariant (sample up to first 100 paragraphs to avoid slowdown)
        Dim pCount As Long: pCount = doc.Paragraphs.Count
        Dim sampleMax As Long: sampleMax = IIf(pCount > 100, 100, pCount)
        Dim i As Long
        For i = 1 To sampleMax
            With doc.Paragraphs(i).Range.Font
                If .Name <> STANDARD_FONT Or .Size <> STANDARD_FONT_SIZE Then
                    issues.Add "Paragraph " & i & " font mismatch: '" & .Name & "' size " & .Size
                    Exit For ' one example is enough
                End If
            End With
        Next i

        ' 2. Margin tolerance (first section)
        Const TOL As Double = 0.1 ' cm tolerance
        With doc.PageSetup
            If Abs(CmFromPoints(.TopMargin) - TOP_MARGIN_CM) > TOL Then issues.Add "Top margin off: " & Format(CmFromPoints(.TopMargin), "0.00") & "cm"
            If Abs(CmFromPoints(.BottomMargin) - BOTTOM_MARGIN_CM) > TOL Then issues.Add "Bottom margin off: " & Format(CmFromPoints(.BottomMargin), "0.00") & "cm"
            If Abs(CmFromPoints(.LeftMargin) - LEFT_MARGIN_CM) > TOL Then issues.Add "Left margin off: " & Format(CmFromPoints(.LeftMargin), "0.00") & "cm"
            If Abs(CmFromPoints(.RightMargin) - RIGHT_MARGIN_CM) > TOL Then issues.Add "Right margin off: " & Format(CmFromPoints(.RightMargin), "0.00") & "cm"
        End With

        ' 3. Hyphenation flag
        If Not doc.Hyphenation.AutoHyphenation Then
            issues.Add "Hyphenation not enabled"
        End If

        ' 4. Consecutive blank paragraphs >1
        Dim consecutiveBlanks As Long: consecutiveBlanks = 0
        For i = 1 To sampleMax
            If Trim$(Replace(doc.Paragraphs(i).Range.Text, vbCr, "")) = "" Then
                consecutiveBlanks = consecutiveBlanks + 1
                If consecutiveBlanks > 1 Then
                    issues.Add "Multiple consecutive blank paragraphs (first seen ending at paragraph " & i & ")"
                    Exit For
                End If
            Else
                consecutiveBlanks = 0
            End If
        Next i

        ' 5. Footer page numbering
        Dim hasPageField As Boolean: hasPageField = False
        Dim f As Field
        For Each f In doc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Fields
            If f.Type = wdFieldPage Then hasPageField = True: Exit For
        Next f
        If Not hasPageField Then issues.Add "Primary footer missing page number field"

        Dim report As String
        If issues.Count = 0 Then
            report = "All self-test checks passed."
        Else
            report = "Self-test found " & issues.Count & " issue(s):" & vbCrLf
            For i = 1 To issues.Count
                report = report & " - " & issues(i) & vbCrLf
            Next i
        End If
        MsgBox report, IIf(issues.Count = 0, vbInformation, vbExclamation), SYSTEM_NAME & " - Self-Test"
        Exit Sub

    FailHard:
        MsgBox "Self-test error: " & Err.Description, vbCritical, SYSTEM_NAME & " - Self-Test"
    End Sub

