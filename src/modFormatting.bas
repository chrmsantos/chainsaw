'================================================================================
' MODULE: modFormatting
' PURPOSE: Paragraph, font, title, numbering, and special token formatting routines.
'================================================================================
Option Explicit

' NOTE: Implementation bodies still reside in chainsaw.bas and will be migrated carefully.

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
