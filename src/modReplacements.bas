Attribute VB_Name = "modReplacements"
'================================================================================
' MODULE: modReplacements
' PURPOSE: Text and specific paragraph replacement routines.
'================================================================================
Option Explicit

' Text replacement and paragraph-specific replacements migrated from chainsaw.bas

Public Function ApplyTextReplacements(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim rng As Range, replacementCount As Long, i As Long
	Set rng = doc.Range
	Dim dOesteVariants() As String: ReDim dOesteVariants(0 To 15)
	dOesteVariants(0) = "d'O": dOesteVariants(1) = "d´O": dOesteVariants(2) = "d`O": dOesteVariants(3) = "d" & Chr(8220) & "O"
	dOesteVariants(4) = "d'o": dOesteVariants(5) = "d´o": dOesteVariants(6) = "d`o": dOesteVariants(7) = "d" & Chr(8220) & "o"
	dOesteVariants(8) = "D'O": dOesteVariants(9) = "D´O": dOesteVariants(10) = "D`O": dOesteVariants(11) = "D" & Chr(8220) & "O"
	dOesteVariants(12) = "D'o": dOesteVariants(13) = "D´o": dOesteVariants(14) = "D`o": dOesteVariants(15) = "D" & Chr(8220) & "o"
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
			Do While .Execute(Replace:=wdReplaceOne)
				replacementCount = replacementCount + 1
				rng.Collapse wdCollapseEnd
			Loop
		End With
	Next i
	Set rng = doc.Range
	Dim dashVariants() As String: ReDim dashVariants(0 To 2)
	dashVariants(0) = " - ": dashVariants(1) = " – ": dashVariants(2) = " — "
	For i = 0 To UBound(dashVariants)
		If dashVariants(i) <> " — " Then
			With rng.Find
				.ClearFormatting
				.Replacement.ClearFormatting
				.Text = dashVariants(i)
				.Replacement.Text = " — "
				.Forward = True
				.Wrap = wdFindStop
				Do While .Execute(Replace:=wdReplaceOne)
					replacementCount = replacementCount + 1
					rng.Collapse wdCollapseEnd
				Loop
			End With
		End If
	Next i
	Set rng = doc.Range
	Dim lineStartDashVariants() As String: ReDim lineStartDashVariants(0 To 1)
	lineStartDashVariants(0) = "^p- ": lineStartDashVariants(1) = "^p– "
	For i = 0 To UBound(lineStartDashVariants)
		With rng.Find
			.ClearFormatting
			.Replacement.ClearFormatting
			.Text = lineStartDashVariants(i)
			.Replacement.Text = "^p— "
			.Forward = True
			.Wrap = wdFindStop
			Do While .Execute(Replace:=wdReplaceOne)
				replacementCount = replacementCount + 1
				rng.Collapse wdCollapseEnd
			Loop
		End With
	Next i
	Set rng = doc.Range
	Dim lineEndDashVariants() As String: ReDim lineEndDashVariants(0 To 1)
	lineEndDashVariants(0) = " -^p": lineEndDashVariants(1) = " –^p"
	For i = 0 To UBound(lineEndDashVariants)
		With rng.Find
			.ClearFormatting
			.Replacement.ClearFormatting
			.Text = lineEndDashVariants(i)
			.Replacement.Text = " —^p"
			.Forward = True
			.Wrap = wdFindStop
			Do While .Execute(Replace:=wdReplaceOne)
				replacementCount = replacementCount + 1
				rng.Collapse wdCollapseEnd
			Loop
		End With
	Next i
	Set rng = doc.Range
	With rng.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Text = "^l"
		.Replacement.Text = " "
		.Forward = True
		.Wrap = wdFindStop
		Do While .Execute(Replace:=wdReplaceOne)
			replacementCount = replacementCount + 1
			rng.Collapse wdCollapseEnd
		Loop
	End With
	On Error Resume Next: Set rng = doc.Range
	With rng.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Text = Chr(11)
		.Replacement.Text = " "
		.Forward = True
		.Wrap = wdFindStop
		Do While .Execute(Replace:=wdReplaceOne)
			If Err.Number <> 0 Then Exit Do
			replacementCount = replacementCount + 1
			rng.Collapse wdCollapseEnd
		Loop
	End With
	If Err.Number <> 0 Then Err.Clear: On Error GoTo ErrHandler
	On Error Resume Next: Set rng = doc.Range
	With rng.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Text = Chr(10)
		.Replacement.Text = " "
		.Forward = True
		.Wrap = wdFindStop
		Do While .Execute(Replace:=wdReplaceOne)
			If Err.Number <> 0 Then Exit Do
			replacementCount = replacementCount + 1
			rng.Collapse wdCollapseEnd
		Loop
	End With
	If Err.Number <> 0 Then Err.Clear: On Error GoTo ErrHandler
	ApplyTextReplacements = True: Exit Function
ErrHandler:
	ApplyTextReplacements = False
End Function

Public Function ApplySpecificParagraphReplacements(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim replacementCount As Long, secondParaIndex As Long, thirdParaIndex As Long
	Dim para As Paragraph, paraText As String, i As Long, actualParaIndex As Long
	For i = 1 To doc.Paragraphs.Count
		Set para = doc.Paragraphs(i)
		paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
		If paraText <> "" Then
			actualParaIndex = actualParaIndex + 1
			If actualParaIndex = 2 Then secondParaIndex = i
			If actualParaIndex = 3 Then thirdParaIndex = i: Exit For
		End If
		If i > 50 Then Exit For
	Next i
	If secondParaIndex > 0 Then
		Set para = doc.Paragraphs(secondParaIndex): paraText = para.Range.Text
		Dim idxStart As Long: idxStart = 1
		Do While idxStart <= Len(paraText) And _
			(Mid$(paraText, idxStart, 1) = " " Or Mid$(paraText, idxStart, 1) = vbTab)
			idxStart = idxStart + 1
		Loop
		If Len(paraText) >= idxStart + 6 Then
			Dim token As String: token = Mid$(paraText, idxStart, 7)
			If token = "Sugiro " Then
				Dim r1 As Range: Set r1 = para.Range.Duplicate
				r1.SetRange r1.Start + (idxStart - 1), r1.Start + (idxStart - 1) + 7: r1.Text = "Requeiro ": replacementCount = replacementCount + 1
			ElseIf token = "Sugere " Then
				Dim r2 As Range: Set r2 = para.Range.Duplicate
				r2.SetRange r2.Start + (idxStart - 1), r2.Start + (idxStart - 1) + 7: r2.Text = "Indica ": replacementCount = replacementCount + 1
			End If
		End If
	End If
	If thirdParaIndex > 0 Then
		Set para = doc.Paragraphs(thirdParaIndex): paraText = para.Range.Text
		Dim originalText As String: originalText = paraText
		If InStr(paraText, " sugerir ") > 0 Then paraText = Replace(paraText, " sugerir ", " indicar "): replacementCount = replacementCount + 1
		If InStr(paraText, " Setor, ") > 0 Then paraText = Replace(paraText, " Setor, ", " setor competente, "): replacementCount = replacementCount + 1
		If paraText <> originalText Then para.Range.Text = paraText
	End If
	Dim rng As Range: Set rng = doc.Range
	With rng.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Text = " A CÂMARA MUNICIPAL DE SANTA BÁRBARA D'OESTE, ESTADO DE SÃO PAULO "
		.Replacement.Text = " a Câmara Municipal de Santa Bárbara d'Oeste, estado de São Paulo, "
		.MatchCase = True
		.Wrap = wdFindContinue
		Do While .Execute(Replace:=wdReplaceOne)
			replacementCount = replacementCount + 1
			rng.Collapse wdCollapseEnd
		Loop
	End With
	Dim wordsToUppercase() As String, j As Long: ReDim wordsToUppercase(0 To 15)
	wordsToUppercase(0) = "aplaude": wordsToUppercase(1) = "Aplaude": wordsToUppercase(2) = "aplauso": wordsToUppercase(3) = "Aplauso"
	wordsToUppercase(4) = "protesta": wordsToUppercase(5) = "Protesta": wordsToUppercase(6) = "protesto": wordsToUppercase(7) = "Protesto"
	wordsToUppercase(8) = "apela": wordsToUppercase(9) = "Apela": wordsToUppercase(10) = "apelo": wordsToUppercase(11) = "Apelo"
	wordsToUppercase(12) = "apoia": wordsToUppercase(13) = "Apoia": wordsToUppercase(14) = "apoio": wordsToUppercase(15) = "Apoio"
	For j = 0 To UBound(wordsToUppercase)
		Set rng = doc.Range
		With rng.Find
			.ClearFormatting
			.Replacement.ClearFormatting
			.Text = wordsToUppercase(j)
			.Replacement.Text = UCase(wordsToUppercase(j))
			.MatchCase = True
			.MatchWholeWord = True
			.Wrap = wdFindContinue
			Do While .Execute(Replace:=wdReplaceOne)
				replacementCount = replacementCount + 1
				rng.Collapse wdCollapseEnd
			Loop
		End With
	Next j
	ApplySpecificParagraphReplacements = True: Exit Function
ErrHandler:
	ApplySpecificParagraphReplacements = False
End Function
