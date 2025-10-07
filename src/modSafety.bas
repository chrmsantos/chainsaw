Attribute VB_Name = "modSafety"
'================================================================================
' MODULE: modSafety
' PURPOSE: Safe wrapper functions for Word object model interactions.
'================================================================================
Option Explicit

Public Function SafeGetCharacterCount(targetRange As Range) As Long
	On Error GoTo Fallback
	SafeGetCharacterCount = targetRange.Characters.Count
	Exit Function
Fallback:
	On Error GoTo Fail
	SafeGetCharacterCount = Len(targetRange.Text)
	Exit Function
Fail:
	SafeGetCharacterCount = 0
End Function

Public Function SafeSetFont(targetRange As Range, fontName As String, fontSize As Long) As Boolean
	On Error GoTo Fail
	With targetRange.Font
		If fontName <> "" Then .Name = fontName
		If fontSize > 0 Then .Size = fontSize
		.Color = wdColorAutomatic
	End With
	SafeSetFont = True
	Exit Function
Fail:
	SafeSetFont = False
End Function

Public Function SafeSetParagraphFormat(para As Paragraph, alignment As Long, leftIndent As Single, firstLineIndent As Single) As Boolean
	On Error GoTo Fail
	With para.Format
		If alignment >= 0 Then .Alignment = alignment
		If leftIndent >= 0 Then .LeftIndent = leftIndent
		If firstLineIndent >= 0 Then .FirstLineIndent = firstLineIndent
	End With
	SafeSetParagraphFormat = True
	Exit Function
Fail:
	SafeSetParagraphFormat = False
End Function

Public Function SafeHasVisualContent(para As Paragraph) As Boolean
	On Error GoTo Alt
	Dim hasImages As Boolean, hasShapes As Boolean, shp As Shape
	hasImages = (para.Range.InlineShapes.Count > 0)
	hasShapes = False
	If Not hasImages Then
		For Each shp In para.Range.ShapeRange
			hasShapes = True: Exit For
		Next shp
	End If
	SafeHasVisualContent = hasImages Or hasShapes
	Exit Function
Alt:
	On Error GoTo Final
	SafeHasVisualContent = (para.Range.InlineShapes.Count > 0)
	Exit Function
Final:
	SafeHasVisualContent = False
End Function

Public Function SafeFindReplace(doc As Document, findText As String, replaceText As String, Optional useWildcards As Boolean = False) As Long
	On Error GoTo Fail
	Dim count As Long: count = 0
	With doc.Range.Find
		.ClearFormatting: .Replacement.ClearFormatting
		.Text = findText: .Replacement.Text = replaceText
		.Forward = True: .Wrap = wdFindContinue
		.Format = False: .MatchCase = False: .MatchWholeWord = False
		.MatchWildcards = useWildcards: .MatchSoundsLike = False: .MatchAllWordForms = False
		Do While .Execute(Replace:=True)
			count = count + 1
			If count > 10000 Then Exit Do
		Loop
	End With
	SafeFindReplace = count
	Exit Function
Fail:
	SafeFindReplace = 0
End Function

Public Function SafeGetLastCharacter(rng As Range) As String
	On Error GoTo Alt
	Dim c As Long: c = SafeGetCharacterCount(rng)
	If c > 0 Then SafeGetLastCharacter = rng.Characters(c).Text Else SafeGetLastCharacter = ""
	Exit Function
Alt:
	On Error GoTo Final
	SafeGetLastCharacter = Right(rng.Text, 1)
	Exit Function
Final:
	SafeGetLastCharacter = ""
End Function
