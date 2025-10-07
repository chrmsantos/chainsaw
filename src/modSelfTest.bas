Attribute VB_Name = "modSelfTest"
'================================================================================
' MODULE: modSelfTest
' PURPOSE: Lightweight regression/self-test to validate core formatting invariants
'          without altering document semantics beyond normal processing.
'================================================================================
Option Explicit

Public Sub ChainsawSelfTest()
    On Error GoTo Fatal
    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "Nenhum documento ativo.", vbExclamation, "Chainsaw SelfTest"
        Exit Sub
    End If

    Dim beforeParaCount As Long, beforeWordCount As Long, beforeInlineShapes As Long
    Dim beforeChars As Long
    beforeParaCount = doc.Paragraphs.Count
    beforeWordCount = doc.Words.Count
    beforeInlineShapes = doc.InlineShapes.Count
    beforeChars = SafeCountChars(doc)

    Dim ok As Boolean
    ok = RunChainsawPipeline()

    Dim afterParaCount As Long, afterWordCount As Long, afterInlineShapes As Long
    Dim afterChars As Long
    afterParaCount = doc.Paragraphs.Count
    afterWordCount = doc.Words.Count
    afterInlineShapes = doc.InlineShapes.Count
    afterChars = SafeCountChars(doc)

    Dim deltaParas As Long, deltaWords As Long, deltaChars As Long
    deltaParas = afterParaCount - beforeParaCount
    deltaWords = afterWordCount - beforeWordCount
    deltaChars = afterChars - beforeChars

    Dim report As String
    report = "Chainsaw Self-Test" & vbCrLf & _
             "Status: " & IIf(ok, "OK", "FALHOU") & vbCrLf & _
             "Parágrafos: " & beforeParaCount & " -> " & afterParaCount & " (" & FormatDelta(deltaParas) & ")" & vbCrLf & _
             "Palavras: " & beforeWordCount & " -> " & afterWordCount & " (" & FormatDelta(deltaWords) & ")" & vbCrLf & _
             "Caracteres: " & beforeChars & " -> " & afterChars & " (" & FormatDelta(deltaChars) & ")" & vbCrLf & _
             "Imagens inline: " & beforeInlineShapes & " -> " & afterInlineShapes & IIf(beforeInlineShapes <> afterInlineShapes, " *ALTERADO*", "") & vbCrLf & _
             "Observação: variações pequenas de caracteres são esperadas devido a limpeza de espaços e normalização." & vbCrLf & _
             "Use este relatório para detectar mudanças inesperadas em refactors."

    MsgBox report, IIf(ok, vbInformation, vbExclamation), "Chainsaw SelfTest"
    Exit Sub
Fatal:
    MsgBox "Erro na self-test: " & Err.Description, vbCritical, "Chainsaw SelfTest"
End Sub

Private Function SafeCountChars(doc As Document) As Long
    On Error Resume Next
    SafeCountChars = doc.Range.Characters.Count
    If Err.Number <> 0 Then SafeCountChars = Len(doc.Range.Text)
End Function

Private Function FormatDelta(v As Long) As String
    If v > 0 Then
        FormatDelta = "+" & v
    Else
        FormatDelta = CStr(v)
    End If
End Function
