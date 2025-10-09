' SPDX-License-Identifier: GPL-3.0-or-later
' =============================================================================
' MODULE: modSmokeTests
' PURPOSE: Lightweight smoke test harness to quickly verify that the Chainsaw
'          pipeline executes without runtime errors and that key formatting
'          operations return success flags. Not a replacement for full tests.
' =============================================================================
Option Explicit

' Runs a basic end-to-end pipeline on the active document.
Public Sub ChainsawSmokeTest()
    Dim doc As Document
    On Error Resume Next
    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "No active document for smoke test", vbExclamation, "Chainsaw SmokeTest"
        Exit Sub
    End If
    On Error GoTo FailFast

    Dim results As Collection
    Set results = New Collection

    ' Load config
    Call LoadConfiguration

    ' Core validations (non-fatal collection of booleans)
    results.Add KeyValResult("CheckWordVersion", CheckWordVersion())
    results.Add KeyValResult("EnsureDocumentEditable", EnsureDocumentEditable(doc))
    results.Add KeyValResult("ValidateDocumentIntegrity", ValidateDocumentIntegrity(doc))

    ' Apply a minimal formatting subset (independent; we avoid heavy ones to keep it fast)
    results.Add KeyValResult("ApplyPageSetup", modFormatting.ApplyPageSetup(doc))
    results.Add KeyValResult("ApplyStdFont", modFormatting.ApplyStdFont(doc))
    results.Add KeyValResult("ApplyStdParagraphs", modFormatting.ApplyStdParagraphs(doc))

    ' Summarize
    Dim report As String
    report = "Chainsaw Smoke Test Results:" & vbCrLf & String(32, "-") & vbCrLf
    Dim i As Long
    For i = 1 To results.Count
        report = report & results(i) & vbCrLf
    Next i
    MsgBox report, vbInformation, "Chainsaw SmokeTest"
    Exit Sub

FailFast:
    MsgBox "Smoke test aborted: " & Err.Description, vbCritical, "Chainsaw SmokeTest"
End Sub

Private Function KeyValResult(name As String, ok As Boolean) As String
    KeyValResult = name & ": " & IIf(ok, "OK", "FAIL")
End Function
