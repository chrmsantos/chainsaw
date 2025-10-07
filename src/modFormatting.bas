Attribute VB_Name = "modFormatting"
'================================================================================
' MODULE: modFormatting
' PURPOSE: Paragraph, font, title, numbering, and special token formatting routines.
'================================================================================
Option Explicit

' NOTE: Centralized formatting implementation. Remaining formatting bodies migrated from chainsaw.bas.

' Public visual content detector (wrapper around safe logic; avoids dependency on legacy modMain private helper)
Public Function HasVisualContent(para As Paragraph) As Boolean
	On Error GoTo Fallback
	HasVisualContent = (para.Range.InlineShapes.Count > 0 Or para.Range.ShapeRange.Count > 0)
	Exit Function
Fallback:
	On Error Resume Next
	HasVisualContent = (para.Range.InlineShapes.Count > 0)
End Function

'--------------------------------------------------------------------------------
' SECOND PARAGRAPH UTILITIES (migrated)
'--------------------------------------------------------------------------------
Private Function CountBlankLinesBefore(doc As Document, paraIndex As Long) As Long
	On Error GoTo ErrHandler
	Dim count As Long, i As Long, para As Paragraph, paraText As String
	For i = paraIndex - 1 To 1 Step -1
		If i <= 0 Then Exit For
		Set para = doc.Paragraphs(i)
		paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
		If paraText = "" And Not HasVisualContent(para) Then
			count = count + 1
		Else
			Exit For
		End If
		If count >= 5 Then Exit For
	Next i
	CountBlankLinesBefore = count: Exit Function
ErrHandler:
	CountBlankLinesBefore = 0
End Function

Private Function CountBlankLinesAfter(doc As Document, paraIndex As Long) As Long
	On Error GoTo ErrHandler
	Dim count As Long, i As Long, para As Paragraph, paraText As String
	For i = paraIndex + 1 To doc.Paragraphs.Count
		If i > doc.Paragraphs.Count Then Exit For
		Set para = doc.Paragraphs(i)
		paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
		If paraText = "" And Not HasVisualContent(para) Then
			count = count + 1
		Else
			Exit For
		End If
		If count >= 5 Then Exit For
	Next i
	CountBlankLinesAfter = count: Exit Function
ErrHandler:
	CountBlankLinesAfter = 0
End Function

Private Function GetSecondParagraphIndex(doc As Document) As Long
	On Error GoTo ErrHandler
	Dim para As Paragraph, paraText As String, i As Long, actualParaIndex As Long
	For i = 1 To doc.Paragraphs.Count
		Set para = doc.Paragraphs(i)
		paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
		If paraText <> "" Or HasVisualContent(para) Then
			actualParaIndex = actualParaIndex + 1
			If actualParaIndex = 2 Then GetSecondParagraphIndex = i: Exit Function
		End If
		If i > 50 Then Exit For
	Next i
	GetSecondParagraphIndex = 0: Exit Function
ErrHandler:
	GetSecondParagraphIndex = 0
End Function

Public Function EnsureSecondParagraphBlankLines(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim secondParaIndex As Long, blankLinesBefore As Long, blankLinesAfter As Long
	secondParaIndex = GetSecondParagraphIndex(doc)
	If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.Count Then
		Dim para As Paragraph: Set para = doc.Paragraphs(secondParaIndex)
		blankLinesBefore = CountBlankLinesBefore(doc, secondParaIndex)
		If blankLinesBefore < 2 Then
			Dim insertionPoint As Range: Set insertionPoint = para.Range: insertionPoint.Collapse wdCollapseStart
			insertionPoint.InsertBefore String(2 - blankLinesBefore, vbCrLf)
			secondParaIndex = secondParaIndex + (2 - blankLinesBefore)
			Set para = doc.Paragraphs(secondParaIndex)
		End If
		blankLinesAfter = CountBlankLinesAfter(doc, secondParaIndex)
		If blankLinesAfter < 2 Then
			Dim insertionPointAfter As Range: Set insertionPointAfter = para.Range: insertionPointAfter.Collapse wdCollapseEnd
			insertionPointAfter.InsertAfter String(2 - blankLinesAfter, vbCrLf)
		End If
	End If
	EnsureSecondParagraphBlankLines = True: Exit Function
ErrHandler:
	EnsureSecondParagraphBlankLines = False
End Function

Public Function FormatFirstParagraph(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim para As Paragraph, paraText As String, i As Long, actualParaIndex As Long, firstParaIndex As Long
	For i = 1 To doc.Paragraphs.Count
		Set para = doc.Paragraphs(i)
		paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
		If paraText <> "" Or HasVisualContent(para) Then
			actualParaIndex = actualParaIndex + 1
			If actualParaIndex = 1 Then firstParaIndex = i: Exit For
		End If
		If i > 20 Then Exit For
	Next i
	If firstParaIndex > 0 And firstParaIndex <= doc.Paragraphs.Count Then
		Set para = doc.Paragraphs(firstParaIndex)
		If HasVisualContent(para) Then
			Dim n As Long, charCount4 As Long: charCount4 = SafeGetCharacterCount(para.Range)
			If charCount4 > 0 Then
				For n = 1 To charCount4
					Dim charRange3 As Range: Set charRange3 = para.Range.Characters(n)
					If charRange3.InlineShapes.Count = 0 Then
						With charRange3.Font: .AllCaps = True: .Bold = True: .Underline = wdUnderlineSingle: End With
					End If
				Next n
			End If
		Else
			With para.Range.Font: .AllCaps = True: .Bold = True: .Underline = wdUnderlineSingle: End With
		End If
		With para.Format: .Alignment = wdAlignParagraphCenter: .LeftIndent = 0: .FirstLineIndent = 0: .RightIndent = 0: End With
	End If
	FormatFirstParagraph = True: Exit Function
ErrHandler:
	FormatFirstParagraph = False
End Function


Public Function EnableHyphenation(doc As Document) As Boolean
	On Error GoTo ErrHandler
	If Not doc.AutoHyphenation Then
		doc.AutoHyphenation = True
		doc.HyphenationZone = CentimetersToPoints(0.63)
		doc.HyphenateCaps = True
	End If
	EnableHyphenation = True: Exit Function
ErrHandler:
	EnableHyphenation = False
End Function

Public Function RemoveWatermark(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim sec As Section, header As HeaderFooter, shp As Shape, i As Long
	For Each sec In doc.Sections
		For Each header In sec.Headers
			If header.Exists And header.Shapes.Count > 0 Then
				For i = header.Shapes.Count To 1 Step -1
					Set shp = header.Shapes(i)
					If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
						If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Or _
						   InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then shp.Delete
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
						   InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then shp.Delete
					End If
				Next i
			End If
		Next header
	Next sec
	RemoveWatermark = True: Exit Function
ErrHandler:
	RemoveWatermark = False
End Function

'--------------------------------------------------------------------------------
' SPECIAL PARAGRAPH FORMATTING: CONSIDERANDO / JUSTIFICATIVA / ANEXO
'--------------------------------------------------------------------------------
Public Function FormatConsiderandoParagraphs(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim para As Paragraph, rawText As String, textNoCrLf As String
	Dim i As Long, startIdx As Long, n As Long, ch As String, code As Long
	Dim totalFormatted As Long
	For i = 1 To doc.Paragraphs.Count
		Set para = doc.Paragraphs(i)
		rawText = para.Range.Text
		textNoCrLf = Replace(Replace(rawText, vbCr, ""), vbLf, "")
		If Len(textNoCrLf) >= 12 Then
			startIdx = 1
			For n = 1 To Len(textNoCrLf)
				ch = Mid$(textNoCrLf, n, 1)
				code = AscW(ch)
				' Skip leading spaces/tabs/quotes/dashes/parentheses/control (<33) and punctuation
				If Not (code = 32 Or code = 9 Or code = 34 Or code = 39 Or code = 45 Or _
						code = 8211 Or code = 8212 Or code = 40 Or code = 41 Or code < 33) Then
					startIdx = n: Exit For
				End If
			Next n
			If startIdx + 11 <= Len(textNoCrLf) Then
				If LCase$(Mid$(textNoCrLf, startIdx, 12)) = "considerando" Then
					Dim rng As Range: Set rng = para.Range.Duplicate
					rng.SetRange rng.Start + (startIdx - 1), rng.Start + (startIdx - 1) + 12
					rng.Text = "CONSIDERANDO"
					rng.Font.Bold = True
					totalFormatted = totalFormatted + 1
				End If
			End If
		End If
	Next i
	FormatConsiderandoParagraphs = True: Exit Function
ErrHandler:
	FormatConsiderandoParagraphs = False
End Function

Public Function FormatJustificativaAnexoParagraphs(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim para As Paragraph, paraText As String, cleanText As String
	Dim i As Long, formattedCount As Long
	For i = 1 To doc.Paragraphs.Count
		Set para = doc.Paragraphs(i)
		If Not HasVisualContent(para) Then
			paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
			cleanText = paraText
			Do While Len(cleanText) > 0 And _
				(Right(cleanText, 1) = "." Or _
				 Right(cleanText, 1) = "," Or _
				 Right(cleanText, 1) = ":" Or _
				 Right(cleanText, 1) = ";")
				cleanText = Left(cleanText, Len(cleanText) - 1)
			Loop
			cleanText = Trim(LCase(cleanText))
			If cleanText = "justificativa" Then
				With para.Format: .LeftIndent = 0: .FirstLineIndent = 0: .Alignment = wdAlignParagraphCenter: End With
				para.Range.Font.Bold = True
				Dim originalEnd As String: originalEnd = ""
				If Len(paraText) > Len(cleanText) Then originalEnd = Right(paraText, Len(paraText) - Len(cleanText))
				para.Range.Text = "Justificativa" & originalEnd & vbCrLf
				formattedCount = formattedCount + 1
			ElseIf IsAnexoPattern(cleanText) Then
				With para.Format: .LeftIndent = 0: .FirstLineIndent = 0: .RightIndent = 0: .Alignment = wdAlignParagraphLeft: End With
				para.Range.Font.Bold = True
				Dim anexoEnd As String: anexoEnd = ""
				If Len(paraText) > Len(cleanText) Then anexoEnd = Right(paraText, Len(paraText) - Len(cleanText))
				Dim anexoText As String: anexoText = IIf(cleanText = "anexo", "Anexo", "Anexos")
				para.Range.Text = anexoText & anexoEnd & vbCrLf
				formattedCount = formattedCount + 1
			ElseIf IsAnteOExpostoPattern(paraText) Then
				para.Range.Font.Bold = True
				formattedCount = formattedCount + 1
			End If
		End If
	Next i
	FormatJustificativaAnexoParagraphs = True: Exit Function
ErrHandler:
	FormatJustificativaAnexoParagraphs = False
End Function

Private Function IsAnexoPattern(text As String) As Boolean
	Dim cleanText As String: cleanText = LCase(Trim(text))
	IsAnexoPattern = (cleanText = "anexo" Or cleanText = "anexos")
End Function

Private Function IsAnteOExpostoPattern(text As String) As Boolean
	Dim cleanText As String: cleanText = LCase(Trim(text))
	If Len(cleanText) = 0 Then IsAnteOExpostoPattern = False: Exit Function
	If Len(cleanText) >= 13 And Left(cleanText, 13) = "ante o exposto" Then
		IsAnteOExpostoPattern = True
	Else
		IsAnteOExpostoPattern = False
	End If
End Function

Public Function InsertHeaderstamp(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim sec As Section, header As HeaderFooter, imgFile As String, imgWidth As Single, imgHeight As Single
	Dim shp As Shape, imgFound As Boolean, baseFolder As String
	imgFile = Trim(Config.headerImagePath)
	If Len(imgFile) = 0 Then
		baseFolder = IIf(doc.Path <> "", doc.Path, Environ("USERPROFILE") & "\Documents")
		If Right(baseFolder, 1) <> "\" Then baseFolder = baseFolder & "\"
		If Dir(baseFolder & "assets\stamp.png") <> "" Then
			imgFile = baseFolder & "assets\stamp.png"
		ElseIf Dir(Environ("USERPROFILE") & "\Documents\chainsaw\assets\stamp.png") <> "" Then
			imgFile = Environ("USERPROFILE") & "\Documents\chainsaw\assets\stamp.png"
		End If
	ElseIf InStr(1, imgFile, ":", vbTextCompare) = 0 And Left(imgFile, 2) <> "\\" Then
		baseFolder = IIf(doc.Path <> "", doc.Path, Environ("USERPROFILE") & "\Documents")
		If Right(baseFolder, 1) <> "\" Then baseFolder = baseFolder & "\"
		If Dir(baseFolder & imgFile) <> "" Then
			imgFile = baseFolder & imgFile
		ElseIf Dir(Environ("USERPROFILE") & "\Documents\chainsaw\" & imgFile) <> "" Then
			imgFile = Environ("USERPROFILE") & "\Documents\chainsaw\" & imgFile
		End If
	End If
	If Dir(imgFile) = "" Then InsertHeaderstamp = False: Exit Function
	imgWidth = CentimetersToPoints(HEADER_IMAGE_MAX_WIDTH_CM)
	imgHeight = imgWidth * HEADER_IMAGE_HEIGHT_RATIO
	For Each sec In doc.Sections
		Set header = sec.Headers(wdHeaderFooterPrimary)
		If header.Exists Then
			header.LinkToPrevious = False
			header.Range.Delete
			Set shp = header.Shapes.AddPicture(FileName:=imgFile, LinkToFile:=False, SaveWithDocument:=msoTrue)
			If Not shp Is Nothing Then
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
			End If
		End If
	Next sec
	InsertHeaderstamp = imgFound: Exit Function
ErrHandler:
	InsertHeaderstamp = False
End Function

Public Function InsertFooterstamp(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim sec As Section, footer As HeaderFooter, rng As Range
	For Each sec In doc.Sections
		Set footer = sec.Footers(wdHeaderFooterPrimary)
		If footer.Exists Then
			footer.LinkToPrevious = False
			Set rng = footer.Range: rng.Delete: Set rng = footer.Range: rng.Collapse Direction:=wdCollapseEnd
			rng.Fields.Add Range:=rng, Type:=wdFieldPage
			Set rng = footer.Range: rng.Collapse Direction:=wdCollapseEnd: rng.Text = "-"
			Set rng = footer.Range: rng.Collapse Direction:=wdCollapseEnd: rng.Fields.Add Range:=rng, Type:=wdFieldNumPages
			With footer.Range
				.Font.Name = STANDARD_FONT: .Font.Size = FOOTER_FONT_SIZE
				.ParagraphFormat.Alignment = wdAlignParagraphCenter
				.Fields.Update
			End With
		End If
	Next sec
	InsertFooterstamp = True: Exit Function
ErrHandler:
	InsertFooterstamp = False
End Function

Public Function FormatNumberedParagraphs(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim para As Paragraph, paraText As String, i As Long
	For i = 1 To doc.Paragraphs.Count
		Set para = doc.Paragraphs(i)
		If Not HasVisualContent(para) Then
			paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
			If IsNumberedParagraph(paraText) Then
				With para.Range.ListFormat
					.RemoveNumbers
					.ApplyNumberDefault
				End With
				Dim cleanedText As String: cleanedText = RemoveManualNumber(paraText)
				para.Range.Text = cleanedText & vbCrLf
			End If
		End If
	Next i
	FormatNumberedParagraphs = True: Exit Function
ErrHandler:
	FormatNumberedParagraphs = False
End Function

Private Function IsNumberedParagraph(text As String) As Boolean
	Dim pattern As String
	pattern = "^[0-9]{1,3}[\.|\)| ]"
	Dim re As Object: Set re = CreateObject("VBScript.RegExp")
	re.Pattern = pattern: re.IgnoreCase = True: re.Global = False
	IsNumberedParagraph = re.Test(text)
End Function

Private Function RemoveManualNumber(text As String) As String
	Dim re As Object: Set re = CreateObject("VBScript.RegExp")
	re.Pattern = "^[0-9]{1,3}[\.|\)| ]+"
	re.IgnoreCase = True: re.Global = False
	RemoveManualNumber = re.Replace(text, "")
End Function

Public Function ApplyPageSetup(doc As Document) As Boolean
	On Error GoTo ErrHandler
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
	ApplyPageSetup = True: Exit Function
ErrHandler:
	ApplyPageSetup = False
End Function

Public Function ApplyStdFont(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim para As Paragraph, i As Long, hasInlineImage As Boolean
	Dim formattedCount As Long, skippedCount As Long
	For i = doc.Paragraphs.Count To 1 Step -1
		Set para = doc.Paragraphs(i)
		hasInlineImage = (para.Range.InlineShapes.Count > 0)
		If Not hasInlineImage And Not HasVisualContent(para) Then
			' Fast path if already correct
			With para.Range.Font
				If .Name = STANDARD_FONT And .Size = STANDARD_FONT_SIZE And .Color = wdColorAutomatic Then GoTo NextPara
			End With
		End If
		If Not hasInlineImage Then
			If SafeSetFont(para.Range, STANDARD_FONT, STANDARD_FONT_SIZE) Then
				formattedCount = formattedCount + 1
			End If
		Else
			formattedCount = formattedCount + 1 ' Conservatively count
		End If
NextPara:
	Next i
	ApplyStdFont = True: Exit Function
ErrHandler:
	ApplyStdFont = False
End Function

Public Sub FormatCharacterByCharacter(para As Paragraph, fontName As String, fontSize As Long, fontColor As Long, removeUnderline As Boolean, removeBold As Boolean)
	On Error Resume Next
	Dim j As Long, charCount As Long, charRange As Range
	charCount = SafeGetCharacterCount(para.Range)
	If charCount = 0 Then Exit Sub
	For j = 1 To charCount
		Set charRange = para.Range.Characters(j)
		If charRange.InlineShapes.Count = 0 Then
			With charRange.Font
				If fontName <> "" Then .Name = fontName
				If fontSize > 0 Then .Size = fontSize
				If fontColor >= 0 Then .Color = fontColor
				If removeUnderline Then .Underline = wdUnderlineNone
				If removeBold Then .Bold = False
			End With
		End If
	Next j
End Sub

'================================================================================
' PARAGRAPH FORMATTING
'================================================================================

Public Function FormatSecondParagraph(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim para As Paragraph, paraText As String, i As Long
	Dim actualParaIndex As Long, secondParaIndex As Long
	For i = 1 To doc.Paragraphs.Count
		Set para = doc.Paragraphs(i)
		paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
		If paraText <> "" Or HasVisualContent(para) Then
			actualParaIndex = actualParaIndex + 1
			If actualParaIndex = 2 Then secondParaIndex = i: Exit For
		End If
		If i > 10 Then Exit For
	Next i
	If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.Count Then
		Set para = doc.Paragraphs(secondParaIndex)
		Dim insertionPoint As Range: Set insertionPoint = para.Range: insertionPoint.Collapse wdCollapseStart
		Dim blankLinesBefore As Long: blankLinesBefore = CountBlankLinesBefore(doc, secondParaIndex)
		If blankLinesBefore < 2 Then
			Dim linesToAdd As Long: linesToAdd = 2 - blankLinesBefore
			insertionPoint.InsertBefore String(linesToAdd, vbCrLf)
			secondParaIndex = secondParaIndex + linesToAdd
			Set para = doc.Paragraphs(secondParaIndex)
		End If
		With para.Format
			.LeftIndent = CentimetersToPoints(9)
			.FirstLineIndent = 0
			.RightIndent = 0
			.Alignment = wdAlignParagraphJustify
		End With
		Dim insertionPointAfter As Range: Set insertionPointAfter = para.Range: insertionPointAfter.Collapse wdCollapseEnd
		Dim blankLinesAfter As Long: blankLinesAfter = CountBlankLinesAfter(doc, secondParaIndex)
		If blankLinesAfter < 2 Then
			insertionPointAfter.InsertAfter String(2 - blankLinesAfter, vbCrLf)
		End If
	End If
	FormatSecondParagraph = True: Exit Function
ErrHandler:
	FormatSecondParagraph = False
End Function

Public Function CentimetersToPoints(ByVal cm As Double) As Single
	On Error Resume Next
	CentimetersToPoints = Application.CentimetersToPoints(cm)
	If Err.Number <> 0 Then CentimetersToPoints = cm * 28.35
End Function

Public Function GetSafeUserName() As String
	On Error GoTo ErrHandler
	Dim rawName As String, c As String, i As Integer, safeName As String
	rawName = Environ("USERNAME"): If rawName = "" Then rawName = Environ("USER")
	If rawName = "" Then On Error Resume Next: rawName = CreateObject("WScript.Network").Username: On Error GoTo ErrHandler
	For i = 1 To Len(rawName)
		c = Mid$(rawName, i, 1)
		If c Like "[A-Za-z0-9_-]" Then safeName = safeName & c
	Next i
	If Len(safeName) = 0 Then safeName = "user"
	GetSafeUserName = safeName: Exit Function
ErrHandler:
	GetSafeUserName = "user"
End Function

'================================================================================
' ADDITIONAL MIGRATED ROUTINES (extracted from modMain)
'================================================================================
Public Function CleanDocumentStructure(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim para As Paragraph, i As Long, firstTextParaIndex As Long
	Dim emptyLinesRemoved As Long, leadingSpacesRemoved As Long, paraCount As Long
	paraCount = doc.Paragraphs.Count: firstTextParaIndex = -1
	For i = 1 To paraCount
		If i > doc.Paragraphs.Count Then Exit For
		Set para = doc.Paragraphs(i)
		Dim paraTextCheck As String
		paraTextCheck = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
		If paraTextCheck <> "" Then firstTextParaIndex = i: Exit For
		If i > 50 Then Exit For
	Next i
	If firstTextParaIndex > 1 Then
		For i = firstTextParaIndex - 1 To 1 Step -1
			If i > doc.Paragraphs.Count Or i < 1 Then Exit For
			Set para = doc.Paragraphs(i)
			Dim paraTextEmpty As String
			paraTextEmpty = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
			If paraTextEmpty = "" Then
				If Not HasVisualContent(para) Then
					para.Range.Delete
					emptyLinesRemoved = emptyLinesRemoved + 1
					paraCount = paraCount - 1
				End If
			End If
		Next i
	End If
	Dim rng As Range: Set rng = doc.Range
	With rng.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Forward = True
		.Wrap = wdFindContinue
		.Format = False
		.MatchWildcards = False
		.Text = "^p "
		.Replacement.Text = "^p"
		Do While .Execute(Replace:=True)
			leadingSpacesRemoved = leadingSpacesRemoved + 1
			If leadingSpacesRemoved > 1000 Then Exit Do
		Loop
		.Text = "^p^t"
		.Replacement.Text = "^p"
		Do While .Execute(Replace:=True)
			leadingSpacesRemoved = leadingSpacesRemoved + 1
			If leadingSpacesRemoved > 1000 Then Exit Do
		Loop
	End With
	Set rng = doc.Range: With rng.Find
		.ClearFormatting: .Replacement.ClearFormatting: .Forward = True: .Wrap = wdFindStop: .Format = False: .MatchWildcards = False
		rng.Start = 0: rng.End = 1
		If rng.Text = " " Or rng.Text = vbTab Then
			Do While rng.End <= doc.Range.End And (Right(rng.Text, 1) = " " Or Right(rng.Text, 1) = vbTab)
				rng.End = rng.End + 1: leadingSpacesRemoved = leadingSpacesRemoved + 1: If leadingSpacesRemoved > 100 Then Exit Do
			Loop
			If rng.Start < rng.End - 1 Then rng.Delete
		End If
	End With
	CleanDocumentStructure = True: Exit Function
ErrHandler:
	CleanDocumentStructure = False
End Function

'================================================================================
' FINAL STRUCTURAL / CLEANUP ROUTINES (migrated from modMain)
'================================================================================
Public Function ApplyStdParagraphs(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim para As Paragraph, hasInlineImage As Boolean
	Dim paragraphIndent As Single, firstIndent As Single
	Dim rightMarginPoints As Single, i As Long
	Dim paraText As String, skippedCount As Long
	rightMarginPoints = 0
	For i = doc.Paragraphs.Count To 1 Step -1
		Set para = doc.Paragraphs(i)
		hasInlineImage = (para.Range.InlineShapes.Count > 0) Or HasVisualContent(para)
		' Text cleanup (avoid altering paragraphs with images)
		If Not hasInlineImage Then
			Dim cleanText As String: cleanText = para.Range.Text
			If InStr(cleanText, "  ") > 0 Or InStr(cleanText, vbTab) > 0 Then
				Do While InStr(cleanText, "  ") > 0
					cleanText = Replace(cleanText, "  ", " ")
				Loop
				cleanText = Replace(cleanText, " " & vbCr, vbCr)
				cleanText = Replace(cleanText, vbCr & " ", vbCr)
				cleanText = Replace(cleanText, vbCr & " ", vbCr) ' duplicate intentional (legacy behavior)
				cleanText = Replace(cleanText, vbTab, " ")
				Do While InStr(cleanText, "  ") > 0
					cleanText = Replace(cleanText, "  ", " ")
				Loop
				If cleanText <> para.Range.Text Then para.Range.Text = cleanText
			End If
		Else
			skippedCount = skippedCount + 1
		End If
		With para.Format
			.LineSpacingRule = wdLineSpacingMultiple
			.LineSpacing = LINE_SPACING
			.RightIndent = rightMarginPoints
			.SpaceBefore = 0: .SpaceAfter = 0
			If para.Alignment = wdAlignParagraphCenter Then
				.LeftIndent = 0: .FirstLineIndent = 0
			Else
				firstIndent = .FirstLineIndent: paragraphIndent = .LeftIndent
				If paragraphIndent >= CentimetersToPoints(5) Then
					.LeftIndent = CentimetersToPoints(9.5)
				ElseIf firstIndent < CentimetersToPoints(5) Then
					.LeftIndent = 0: .FirstLineIndent = CentimetersToPoints(1.5)
				End If
			End If
		End With
		If para.Alignment = wdAlignParagraphLeft Then para.Alignment = wdAlignParagraphJustify
	Next i
	ApplyStdParagraphs = True: Exit Function
ErrHandler:
	ApplyStdParagraphs = False
End Function

Public Function CleanMultipleSpaces(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim rng As Range, spacesRemoved As Long, totalOps As Long
	Set rng = doc.Range
	With rng.Find
		.ClearFormatting: .Replacement.ClearFormatting
		.Forward = True: .Wrap = wdFindContinue
		.Format = False: .MatchWildcards = False
		Do
			.Text = "  ": .Replacement.Text = " "
			Dim cnt As Long: cnt = 0
			Do While .Execute(Replace:=wdReplaceOne)
				cnt = cnt + 1: spacesRemoved = spacesRemoved + 1: rng.Collapse wdCollapseEnd
				If cnt Mod 200 = 0 Then DoEvents
				If spacesRemoved > 2000 Then Exit Do
			Loop
			totalOps = totalOps + 1
			If cnt = 0 Or totalOps > 10 Then Exit Do
		Loop
	End With
	' Tabs & residual doubles
	Set rng = doc.Range
	With rng.Find
		.ClearFormatting: .Replacement.ClearFormatting
		.MatchWildcards = False: .Forward = True: .Wrap = wdFindContinue
		.Text = "^t^t": .Replacement.Text = "^t": Do While .Execute(Replace:=wdReplaceOne): spacesRemoved = spacesRemoved + 1: rng.Collapse wdCollapseEnd: If spacesRemoved > 2000 Then Exit Do: Loop
		.Text = "^t": .Replacement.Text = " ": Do While .Execute(Replace:=wdReplaceOne): spacesRemoved = spacesRemoved + 1: If spacesRemoved > 2000 Then Exit Do: Loop
	End With
	CleanMultipleSpaces = True: Exit Function
ErrHandler:
	CleanMultipleSpaces = False
End Function

Public Function LimitSequentialEmptyLines(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim rng As Range, linesRemoved As Long, totalRepl As Long
	Set rng = doc.Range
	With rng.Find
		.ClearFormatting: .Replacement.ClearFormatting
		.Forward = True: .Wrap = wdFindContinue
		.Format = False: .MatchWildcards = False
		.Text = "^p^p^p^p": .Replacement.Text = "^p^p"
		Do While .Execute(Replace:=True)
			linesRemoved = linesRemoved + 1: totalRepl = totalRepl + 1: If totalRepl > 500 Then Exit Do
		Loop
		.Text = "^p^p^p": .Replacement.Text = "^p^p"
		Do While .Execute(Replace:=True)
			linesRemoved = linesRemoved + 1: totalRepl = totalRepl + 1: If totalRepl > 500 Then Exit Do
		Loop
	End With
	LimitSequentialEmptyLines = True: Exit Function
ErrHandler:
	LimitSequentialEmptyLines = False
End Function

Public Function EnsureParagraphSeparation(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim i As Long, para As Paragraph, nextPara As Paragraph, inserted As Long
	For i = 1 To doc.Paragraphs.Count - 1
		Set para = doc.Paragraphs(i): Set nextPara = doc.Paragraphs(i + 1)
		Dim t1 As String, t2 As String
		t1 = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
		t2 = Trim(Replace(Replace(nextPara.Range.Text, vbCr, ""), vbLf, ""))
		If t1 <> "" And t2 <> "" Then
			If nextPara.Range.Start - para.Range.End <= 1 Then
				Dim r As Range: Set r = doc.Range(para.Range.End - 1, para.Range.End - 1)
				r.Text = vbCrLf: inserted = inserted + 1
			End If
		End If
		If i Mod 500 = 0 Then DoEvents
	Next i
	EnsureParagraphSeparation = True: Exit Function
ErrHandler:
	EnsureParagraphSeparation = False
End Function

Public Function ConfigureDocumentView(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim w As Window: Set w = doc.ActiveWindow
	w.View.Zoom.Percentage = 110
	ConfigureDocumentView = True: Exit Function
ErrHandler:
	ConfigureDocumentView = False
End Function


Public Function FormatDocumentTitle(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim firstPara As Paragraph, paraText As String, words() As String, i As Long, newText As String
	For i = 1 To doc.Paragraphs.Count
		Set firstPara = doc.Paragraphs(i)
		paraText = Trim(Replace(Replace(firstPara.Range.Text, vbCr, ""), vbLf, ""))
		If paraText <> "" Then Exit For
	Next i
	If paraText = "" Then FormatDocumentTitle = True: Exit Function
	If Right(paraText, 1) = "." Then paraText = Left(paraText, Len(paraText) - 1)
	Dim isProp As Boolean, firstWord As String
	words = Split(paraText, " ")
	If UBound(words) >= 0 Then
		firstWord = LCase(Trim(words(0)))
		If firstWord = "indicação" Or firstWord = "requerimento" Or firstWord = "moção" Then isProp = True
	End If
	If isProp And UBound(words) >= 0 Then
		For i = 0 To UBound(words) - 1
			If i > 0 Then newText = newText & " "
			newText = newText & words(i)
		Next i
		If newText <> "" Then newText = newText & " "
		newText = newText & "$NUMERO$/$ANO$"
	Else
		newText = paraText
	End If
	firstPara.Range.Text = UCase(newText) & vbCrLf
	With firstPara.Range.Font: .Bold = True: .Underline = wdUnderlineSingle: End With
	With firstPara.Format: .Alignment = wdAlignParagraphCenter: .LeftIndent = 0: .FirstLineIndent = 0: .RightIndent = 0: .SpaceBefore = 0: .SpaceAfter = 6: End With
	FormatDocumentTitle = True: Exit Function
ErrHandler:
	FormatDocumentTitle = False
End Function

Private Function HasSubstantiveTextAfterNumber(fullText As String, numberToken As String) As Boolean
	On Error GoTo ErrHandler
	Dim remainder As String
	remainder = Mid(fullText, Len(numberToken) + 1)
	remainder = Trim(remainder)
	If Len(remainder) = 0 Then HasSubstantiveTextAfterNumber = False: Exit Function
	Dim firstWord As String, spacePos As Long
	spacePos = InStr(remainder, " ")
	If spacePos > 0 Then
		firstWord = Left(remainder, spacePos - 1)
	Else
		firstWord = remainder
	End If
	If Len(firstWord) < 2 Then HasSubstantiveTextAfterNumber = False: Exit Function
	HasSubstantiveTextAfterNumber = True: Exit Function
ErrHandler:
	HasSubstantiveTextAfterNumber = False
End Function
