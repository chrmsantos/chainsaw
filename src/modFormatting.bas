'================================================================================
' MODULE: modFormatting
' PURPOSE: Paragraph, font, title, numbering, and special token formatting routines.
'================================================================================
Option Explicit

' NOTE: Centralized formatting implementation. Remaining formatting bodies migrated from chainsaw.bas.

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

Public Function FormatSecondParagraph(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim para As Paragraph, paraText As String, i As Long, actualParaIndex As Long, secondParaIndex As Long
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
		With para.Format: .LeftIndent = CentimetersToPoints(9): .FirstLineIndent = 0: .RightIndent = 0: .Alignment = wdAlignParagraphJustify: End With
		Dim insertionPointAfter As Range: Set insertionPointAfter = para.Range: insertionPointAfter.Collapse wdCollapseEnd
		Dim blankLinesAfter As Long: blankLinesAfter = CountBlankLinesAfter(doc, secondParaIndex)
		If blankLinesAfter < 2 Then insertionPointAfter.InsertAfter String(2 - blankLinesAfter, vbCrLf)
	End If
	FormatSecondParagraph = True: Exit Function
ErrHandler:
	FormatSecondParagraph = False
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
Public Function ApplyStdParagraphs(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim para As Paragraph, i As Long
	Dim hasInlineImage As Boolean, skippedCount As Long, rightMarginPoints As Single
	rightMarginPoints = 0
	For i = doc.Paragraphs.Count To 1 Step -1
		Set para = doc.Paragraphs(i)
		hasInlineImage = (para.Range.InlineShapes.Count > 0)
		If Not hasInlineImage And HasVisualContent(para) Then
			hasInlineImage = True: skippedCount = skippedCount + 1
		End If
		Dim cleanText As String: cleanText = para.Range.Text
		If InStr(cleanText, "  ") > 0 Or InStr(cleanText, vbTab) > 0 Then
			Do While InStr(cleanText, "  ") > 0: cleanText = Replace(cleanText, "  ", " "): Loop
			cleanText = Replace(cleanText, " " & vbCr, vbCr)
			cleanText = Replace(cleanText, vbCr & " ", vbCr)
			cleanText = Replace(cleanText, " " & vbLf, vbLf)
			cleanText = Replace(cleanText, vbLf & " ", vbLf)
			Do While InStr(cleanText, vbTab & vbTab) > 0: cleanText = Replace(cleanText, vbTab & vbTab, vbTab): Loop
			cleanText = Replace(cleanText, vbTab, " ")
			Do While InStr(cleanText, "  ") > 0: cleanText = Replace(cleanText, "  ", " "): Loop
		End If
		If cleanText <> para.Range.Text And Not hasInlineImage Then para.Range.Text = cleanText
		With para.Format
			.LineSpacingRule = wdLineSpacingMultiple
			.LineSpacing = LINE_SPACING
			.RightIndent = rightMarginPoints
			.SpaceBefore = 0
			.SpaceAfter = 0
			If para.Alignment = wdAlignParagraphCenter Then
				.LeftIndent = 0: .FirstLineIndent = 0
			Else
				If .LeftIndent >= CentimetersToPoints(5) Then
					.LeftIndent = CentimetersToPoints(9.5)
				ElseIf .FirstLineIndent < CentimetersToPoints(5) Then
					.LeftIndent = 0
					.FirstLineIndent = CentimetersToPoints(1.5)
				End If
			End If
		End With
		If para.Alignment = wdAlignParagraphLeft Then para.Alignment = wdAlignParagraphJustify
	Next i
	ApplyStdParagraphs = True: Exit Function
ErrHandler:
	ApplyStdParagraphs = False
End Function

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

Private Function EnsureSecondParagraphBlankLines(doc As Document) As Boolean
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
						With charRange3.Font
							.AllCaps = True: .Bold = True: .Underline = wdUnderlineSingle
						End With
					End If
				Next n
			End If
		Else
			With para.Range.Font: .AllCaps = True: .Bold = True: .Underline = wdUnderlineSingle: End With
		End If
		With para.Format
			.Alignment = wdAlignParagraphCenter: .LeftIndent = 0: .FirstLineIndent = 0: .RightIndent = 0
		End With
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
		EnableHyphenation = True
	Else
		EnableHyphenation = True
	End If
	Exit Function
ErrHandler:
	EnableHyphenation = False
End Function

Public Function RemoveWatermark(doc As Document) As Boolean
	On Error GoTo ErrHandler
	Dim sec As Section, header As HeaderFooter, shp As Shape, i As Long, removedCount As Long
	For Each sec In doc.Sections
		For Each header In sec.Headers
			If header.Exists And header.Shapes.Count > 0 Then
				For i = header.Shapes.Count To 1 Step -1
					Set shp = header.Shapes(i)
					If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
						If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Or _
						   InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then
							shp.Delete: removedCount = removedCount + 1
						End If
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
						   InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then
							shp.Delete: removedCount = removedCount + 1
						End If
					End If
				Next i
			End If
		Next header
	Next sec
	RemoveWatermark = True: Exit Function
ErrHandler:
	RemoveWatermark = False
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
	Dim sec As Section, footer As HeaderFooter, rng As Range, sectionsProcessed As Long
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
			sectionsProcessed = sectionsProcessed + 1
		End If
	Next sec
	InsertFooterstamp = True: Exit Function
ErrHandler:
	InsertFooterstamp = False
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
