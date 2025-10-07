'Attribute VB_Name = "modValidation"
'================================================================================
' MODULE: modValidation
' PURPOSE: Document integrity, proposition type, and consistency validation functions.
'================================================================================
Option Explicit

Public Function ValidateContentConsistency(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim para As Paragraph, paraText As String, i As Long, actualParaIndex As Long
	Dim secondParaIndex As Long, secondParaText As String, restOfDocumentText As String
	actualParaIndex = 0
	For i = 1 To doc.Paragraphs.Count
		Set para = doc.Paragraphs(i)
		paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
		If paraText <> "" Then
			actualParaIndex = actualParaIndex + 1
			If actualParaIndex = 2 Then secondParaIndex = i: secondParaText = paraText: Exit For
		End If
		If i > 50 Then Exit For
	Next i
	If secondParaIndex = 0 Or secondParaText = "" Then ValidateContentConsistency = True: Exit Function
	restOfDocumentText = "": actualParaIndex = 0
	For i = 1 To doc.Paragraphs.Count
		Set para = doc.Paragraphs(i)
		paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
		If paraText <> "" Then
			actualParaIndex = actualParaIndex + 1
			If actualParaIndex >= 3 Then restOfDocumentText = restOfDocumentText & " " & paraText
		End If
	Next i
	If restOfDocumentText = "" Then ValidateContentConsistency = True: Exit Function
	Dim commonWordsCount As Long: commonWordsCount = CountCommonWords(secondParaText, restOfDocumentText)
	If commonWordsCount < 2 Then
		Dim inconsistencyMessage As String, userResponse As VbMsgBoxResult
		inconsistencyMessage = ReplacePlaceholders(MSG_INCONSISTENCY_WARNING, "Ementa", Left(secondParaText, 200), "COMMON", CStr(commonWordsCount))
		userResponse = MsgBox(NormalizeForUI(inconsistencyMessage), vbYesNo + vbExclamation + vbDefaultButton2, NormalizeForUI(TITLE_CONSISTENCY))
		If userResponse = vbNo Then ValidateContentConsistency = False: Exit Function
	End If
	ValidateContentConsistency = True: Exit Function
ErrHandler:
	ValidateContentConsistency = False
End Function

' Proposition type validation migrated from modMain.
Public Function ValidatePropositionType(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim firstPara As Paragraph, paraText As String, i As Long, firstWord As String
	Dim userResponse As VbMsgBoxResult
	For i = 1 To doc.Paragraphs.Count
		Set firstPara = doc.Paragraphs(i)
		paraText = Trim(Replace(Replace(firstPara.Range.Text, vbCr, ""), vbLf, ""))
		If paraText <> "" Then Exit For
	Next i
	If paraText = "" Then ValidatePropositionType = True: Exit Function
	Dim words() As String: words = Split(paraText, " ")
	If UBound(words) >= 0 Then firstWord = LCase(Trim(words(0)))
	If firstWord = "indicação" Or firstWord = "requerimento" Or firstWord = "moção" Then
		ValidatePropositionType = True: Exit Function
	End If
	Application.StatusBar = "Confirmar tipo de documento..."
	Dim confirmationMessage As String
	confirmationMessage = ReplacePlaceholders(MSG_DOC_TYPE_WARNING, _
						"FIRSTWORD", UCase(firstWord), _
						"DOCSTART", Left(paraText, 150))
	userResponse = MsgBox(NormalizeForUI(confirmationMessage), vbYesNo + vbQuestion + vbDefaultButton2, _
				NormalizeForUI(TITLE_DOC_TYPE))
	If userResponse = vbYes Then
		Application.StatusBar = "Continuando com documento não padronizado..."
		ValidatePropositionType = True
	Else
		Application.StatusBar = "Processamento cancelado pelo usuário"
		MsgBox NormalizeForUI(MSG_PROCESSING_CANCELLED), vbInformation, NormalizeForUI(TITLE_OPERATION_CANCELLED)
		ValidatePropositionType = False
	End If
	Exit Function
ErrHandler:
	ValidatePropositionType = False
End Function

Private Function CountCommonWords(text1 As String, text2 As String) As Long
	On Error GoTo ErrHandler
	Dim words1() As String, words2() As String, i As Long, j As Long, commonCount As Long
	Dim word1 As String, word2 As String
	text1 = CleanTextForComparison(text1): text2 = CleanTextForComparison(text2)
	words1 = Split(text1, " "): words2 = Split(text2, " ")
	For i = 0 To UBound(words1)
		word1 = Trim(words1(i))
		If Len(word1) >= 4 And Not IsCommonWord(word1) Then
			For j = 0 To UBound(words2)
				word2 = Trim(words2(j))
				If word1 = word2 Then commonCount = commonCount + 1: Exit For
			Next j
		End If
	Next i
	CountCommonWords = commonCount: Exit Function
ErrHandler:
	CountCommonWords = 0
End Function

Private Function CleanTextForComparison(text As String) As String
	Dim cleanedText As String, i As Long, ch As String, result As String
	cleanedText = LCase(text)
	For i = 1 To Len(cleanedText)
		ch = Mid$(cleanedText, i, 1)
		If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = " " Then
			result = result & ch
		Else
			result = result & " "
		End If
	Next i
	Do While InStr(result, "  ") > 0: result = Replace(result, "  ", " "): Loop
	CleanTextForComparison = Trim(result)
End Function

Private Function IsCommonWord(word As String) As Boolean
	Dim commonWords() As String, i As Long
	ReDim commonWords(0 To 49)
	commonWords(0) = "que": commonWords(1) = "para": commonWords(2) = "com": commonWords(3) = "uma": commonWords(4) = "por"
	commonWords(5) = "dos": commonWords(6) = "das": commonWords(7) = "este": commonWords(8) = "esta": commonWords(9) = "essa"
	commonWords(10) = "esse": commonWords(11) = "seu": commonWords(12) = "sua": commonWords(13) = "seus": commonWords(14) = "suas"
	commonWords(15) = "mais": commonWords(16) = "muito": commonWords(17) = "entre": commonWords(18) = "sobre": commonWords(19) = "após"
	commonWords(20) = "antes": commonWords(21) = "durante": commonWords(22) = "através": commonWords(23) = "mediante": commonWords(24) = "junto"
	commonWords(25) = "desde": commonWords(26) = "até": commonWords(27) = "contra": commonWords(28) = "favor": commonWords(29) = "deve"
	commonWords(30) = "devem": commonWords(31) = "pode": commonWords(32) = "podem": commonWords(33) = "será": commonWords(34) = "serão"
	commonWords(35) = "está": commonWords(36) = "estão": commonWords(37) = "foram": commonWords(38) = "sendo": commonWords(39) = "tendo"
	commonWords(40) = "onde": commonWords(41) = "quando": commonWords(42) = "como": commonWords(43) = "porque": commonWords(44) = "portanto"
	commonWords(45) = "assim": commonWords(46) = "então": commonWords(47) = "ainda": commonWords(48) = "também": commonWords(49) = "apenas"
	word = LCase(Trim(word))
	For i = 0 To UBound(commonWords)
		If word = commonWords(i) Then IsCommonWord = True: Exit Function
	Next i
	IsCommonWord = False
End Function
