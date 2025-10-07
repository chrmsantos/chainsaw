' SPDX-License-Identifier: GPL-3.0-or-later
' =============================================================================
' MODULE: modValidation
' PURPOSE: Document validation routines extracted from monolithic module.
' =============================================================================
Option Explicit

Public Function CheckWordVersion() As Boolean
    On Error GoTo ErrorHandler
    CheckWordVersion = False
    If Not Config.CheckWordVersion Then
        CheckWordVersion = True
        Exit Function
    End If
    Dim curVer As Double
    curVer = CDbl(Val(Application.Version))
    CheckWordVersion = (curVer >= Config.minWordVersion)
    Exit Function
ErrorHandler:
    CheckWordVersion = True ' fail-soft
End Function

Public Function EnsureDocumentEditable(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    EnsureDocumentEditable = False
    If doc Is Nothing Then Exit Function
    On Error Resume Next: doc.Final = False: On Error GoTo ErrorHandler
    On Error Resume Next
    If Not Application.ActiveProtectedViewWindow Is Nothing Then
        Application.ActiveProtectedViewWindow.Edit
    End If
    On Error GoTo ErrorHandler
    If doc.ProtectionType <> wdNoProtection Or doc.ReadOnly Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox(NormalizeForUI("Documento protegido ou somente leitura. Deseja salvar uma cópia para editar?"), _
                      vbYesNo + vbQuestion, NormalizeForUI(TITLE_ENABLE_EDITING))
        If resp = vbYes Then
            On Error Resume Next
            If Application.Dialogs(wdDialogFileSaveAs).Show <> -1 Then
                On Error GoTo ErrorHandler
                Exit Function
            End If
            On Error GoTo ErrorHandler
        Else
            Exit Function
        End If
    End If
    If doc.ProtectionType = wdNoProtection And Not doc.ReadOnly Then
        EnsureDocumentEditable = True
    End If
    Exit Function
ErrorHandler:
    EnsureDocumentEditable = False
End Function

Public Function ValidateDocumentIntegrity(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    ValidateDocumentIntegrity = False
    If doc Is Nothing Then
        MsgBox NormalizeForUI(MSG_INACCESSIBLE), vbCritical, NormalizeForUI(TITLE_INTEGRITY_ERROR)
        Exit Function
    End If
    On Error Resume Next
    Dim isProtected As Boolean
    isProtected = (doc.ProtectionType <> wdNoProtection)
    If Err.Number <> 0 Then isProtected = False
    On Error GoTo ErrorHandler
    If isProtected Then
        Dim protMsg As String
        protMsg = ReplacePlaceholders(MSG_PROTECTED, "PROT", GetProtectionType(doc))
        If vbNo = MsgBox(NormalizeForUI(protMsg), vbYesNo + vbExclamation, NormalizeForUI(TITLE_PROTECTED)) Then Exit Function
    End If
    If doc.Paragraphs.Count < 1 Then
        MsgBox NormalizeForUI(MSG_EMPTY_DOC), vbExclamation, NormalizeForUI(TITLE_EMPTY_DOC)
        Exit Function
    End If
    Dim docSize As Long
    On Error Resume Next: docSize = doc.Range.Characters.Count: If Err.Number <> 0 Then docSize = 0: On Error GoTo ErrorHandler
    If docSize > 500000 Then
        Dim continueResponse As VbMsgBoxResult
        Dim largeMsg As String
        largeMsg = ReplacePlaceholders(MSG_LARGE_DOC, "SIZE", Format(docSize, "#,##0"))
        continueResponse = MsgBox(NormalizeForUI(largeMsg), vbYesNo + vbQuestion, NormalizeForUI(TITLE_LARGE_DOC))
        If continueResponse = vbNo Then Exit Function
    End If
    If Not doc.Saved And doc.Path <> "" Then
        Dim saveResponse As VbMsgBoxResult
        saveResponse = MsgBox(NormalizeForUI(MSG_UNSAVED), vbYesNoCancel + vbQuestion, NormalizeForUI(TITLE_UNSAVED))
        Select Case saveResponse
            Case vbYes: doc.Save
            Case vbCancel: Exit Function
        End Select
    End If
    ValidateDocumentIntegrity = True
    Exit Function
ErrorHandler:
    Dim valErr As String
    valErr = ReplacePlaceholders(MSG_VALIDATION_ERROR, "ERR", Err.Description)
    MsgBox NormalizeForUI(valErr), vbCritical, NormalizeForUI(TITLE_VALIDATION_ERROR)
    ValidateDocumentIntegrity = False
End Function

Public Function ValidatePropositionType(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim firstPara As Paragraph, firstWord As String, paraText As String, i As Long, userResponse As VbMsgBoxResult
    For i = 1 To doc.Paragraphs.Count
        Set firstPara = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(firstPara.Range.Text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then Exit For
    Next i
    If paraText = "" Then
        ValidatePropositionType = True
        Exit Function
    End If
    Dim words() As String: words = Split(paraText, " ")
    If UBound(words) >= 0 Then firstWord = LCase(Trim(words(0)))
    If firstWord = "indica��o" Or firstWord = "requerimento" Or firstWord = "mo��o" Then
        ValidatePropositionType = True
    Else
        Application.StatusBar = "Waiting for user confirmation about document type..."
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
            MsgBox NormalizeForUI(MSG_PROCESSING_CANCELLED), vbInformation, NormalizeForUI(TITLE_OPERATION_CANCELLED)
            ValidatePropositionType = False
        End If
    End If
    Exit Function
ErrorHandler:
    ValidatePropositionType = False
End Function

Public Function ValidateContentConsistency(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Application.StatusBar = "Validating consistency between summary and body..."
    Dim para As Paragraph, paraText As String, i As Long, actualParaIndex As Long
    Dim secondParaIndex As Long, secondParaText As String, restOfDocumentText As String
    actualParaIndex = 0: secondParaIndex = 0
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then
            actualParaIndex = actualParaIndex + 1
            If actualParaIndex = 2 Then
                secondParaIndex = i: secondParaText = paraText
                Exit For
            End If
        End If
        If i > 50 Then Exit For
    Next i
    If secondParaIndex = 0 Or secondParaText = "" Then
        ValidateContentConsistency = True
        Exit Function
    End If
    restOfDocumentText = "": actualParaIndex = 0
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then
            actualParaIndex = actualParaIndex + 1
            If actualParaIndex >= 3 Then restOfDocumentText = restOfDocumentText & " " & paraText
        End If
    Next i
    If restOfDocumentText = "" Then
        ValidateContentConsistency = True
        Exit Function
    End If
    Dim commonWordsCount As Long
    commonWordsCount = CountCommonWords(secondParaText, restOfDocumentText)
    If commonWordsCount < 2 Then
        Dim inconsistencyMessage As String, userResponse As VbMsgBoxResult
        inconsistencyMessage = ReplacePlaceholders(MSG_INCONSISTENCY_WARNING, _
                               "Ementa", Left(secondParaText, 200), _
                               "COMMON", CStr(commonWordsCount))
        userResponse = MsgBox(NormalizeForUI(inconsistencyMessage), vbYesNo + vbExclamation + vbDefaultButton2, _
                               NormalizeForUI(TITLE_CONSISTENCY))
        If userResponse = vbNo Then
            Application.StatusBar = "Formatting stopped - inconsistency detected"
            ValidateContentConsistency = False
            Exit Function
        End If
    End If
    ValidateContentConsistency = True
    Exit Function
ErrorHandler:
    ValidateContentConsistency = False
End Function
