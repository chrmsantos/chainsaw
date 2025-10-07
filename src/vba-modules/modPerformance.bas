' SPDX-License-Identifier: GPL-3.0-or-later
' =============================================================================
' MODULE: modPerformance
' PURPOSE: Performance optimization helpers and paragraph batch processing
'          extracted from monolithic module.
' =============================================================================
Option Explicit

' Threshold constants remain in monolith; called code assumes they are available.
' If needed later we can relocate OPTIMIZATION_THRESHOLD / MAX_PARAGRAPH_BATCH_SIZE here.

Public Function InitializePerformanceOptimization() As Boolean
    On Error GoTo ErrorHandler
    InitializePerformanceOptimization = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    OptimizeWordSettings
    InitializePerformanceOptimization = True
    Exit Function
ErrorHandler:
    InitializePerformanceOptimization = False
End Function

Public Sub OptimizeWordSettings()
    On Error Resume Next
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
End Sub

Public Function RestorePerformanceSettings() As Boolean
    On Error GoTo ErrorHandler
    RestorePerformanceSettings = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    RestorePerformanceSettings = True
    Exit Function
ErrorHandler:
    RestorePerformanceSettings = False
End Function

Public Function OptimizedFindReplace(findText As String, replaceText As String, Optional searchRange As Range = Nothing) As Long
    On Error GoTo ErrorHandler
    OptimizedFindReplace = 0
    OptimizedFindReplace = BulkFindReplace(findText, replaceText, searchRange)
    Exit Function
ErrorHandler:
    OptimizedFindReplace = 0
End Function

Public Function BulkFindReplace(findText As String, replaceText As String, Optional searchRange As Range = Nothing) As Long
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
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        BulkFindReplace = .Execute(Replace:=wdReplaceAll)
    End With
    Exit Function
ErrorHandler:
    BulkFindReplace = 0
End Function

Public Function StandardFindReplace(findText As String, replaceText As String, Optional searchRange As Range = Nothing) As Long
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
    StandardFindReplace = 0
End Function

Public Function OptimizedParagraphProcessing(processingFunction As String) As Boolean
    On Error GoTo ErrorHandler
    OptimizedParagraphProcessing = False
    OptimizedParagraphProcessing = BatchProcessParagraphs(processingFunction)
    Exit Function
ErrorHandler:
    OptimizedParagraphProcessing = False
End Function

Public Function BatchProcessParagraphs(processingFunction As String) As Boolean
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
        If Not ProcessParagraphBatch(i, endIndex, processingFunction) Then Exit Function
        If i Mod (batchSize * 5) = 0 Then DoEvents
    Next i
    BatchProcessParagraphs = True
    Exit Function
ErrorHandler:
    BatchProcessParagraphs = False
End Function

Public Function StandardProcessParagraphs(processingFunction As String) As Boolean
    On Error GoTo ErrorHandler
    StandardProcessParagraphs = False
    Dim doc As Document
    Set doc = ActiveDocument
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        Select Case processingFunction
            Case "FORMAT":      FormatParagraph para
            Case "CLEAN":       CleanParagraph para
            Case "VALIDATE":    ValidateParagraph para
        End Select
    Next para
    StandardProcessParagraphs = True
    Exit Function
ErrorHandler:
    StandardProcessParagraphs = False
End Function

Public Function ProcessParagraphBatch(startIndex As Long, endIndex As Long, processingFunction As String) As Boolean
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
                Case "FORMAT":      FormatParagraph para
                Case "CLEAN":       CleanParagraph para
                Case "VALIDATE":    ValidateParagraph para
            End Select
        End If
    Next i
    ProcessParagraphBatch = True
    Exit Function
ErrorHandler:
    ProcessParagraphBatch = False
End Function

' NOTE: FormatParagraph / CleanParagraph / ValidateParagraph remain in the monolith
' (formatting module extraction will move them later). This module only batches.
