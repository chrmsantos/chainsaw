' SPDX-License-Identifier: GPL-3.0-or-later
' =============================================================================
' MODULE: modFormatting
' PURPOSE: Encapsulates all document formatting, header/footer insertion,
'          paragraph-specific transformations, replacements, and related
'          helper routines extracted from the legacy monolith.
' NOTE:    Behavior preserved verbatim; logging/backups removed earlier.
' =============================================================================
Option Explicit
Option Private Module

' Formatting relies on public constants exposed by modOrquestration (now Public Const)
' to avoid duplication (STANDARD_FONT, margins, image sizes, etc.).

' ==================================================================================
' SAFE HELPERS (character count, font set, visual content detection)
' ==================================================================================
Public Function SafeGetCharacterCount(targetRange As Range) As Long
    On Error GoTo ErrorHandler
    SafeGetCharacterCount = targetRange.Characters.Count
    Exit Function
ErrorHandler:
    On Error Resume Next
    SafeGetCharacterCount = Len(targetRange.Text)
End Function

Public Function SafeSetFont(targetRange As Range, fontName As String, fontSize As Long) As Boolean
    On Error GoTo ErrorHandler
    With targetRange.Font
        .Name = fontName
        .Size = fontSize
        .Color = wdColorAutomatic
    End With
    SafeSetFont = True
    Exit Function
ErrorHandler:
    SafeSetFont = False
End Function

Private Function SafeHasVisualContent(para As Paragraph) As Boolean
    On Error GoTo ErrorHandler
    Dim hasImages As Boolean, hasShapes As Boolean
    If para.Range.InlineShapes.Count > 0 Then hasImages = True
    If para.Range.ShapeRange.Count > 0 Then hasShapes = True
    SafeHasVisualContent = hasImages Or hasShapes
    Exit Function
ErrorHandler:
    On Error Resume Next
    SafeHasVisualContent = (para.Range.InlineShapes.Count > 0)
End Function

Public Function HasVisualContent(para As Paragraph) As Boolean
    On Error Resume Next
    HasVisualContent = SafeHasVisualContent(para)
End Function

' ==================================================================================
' PAGE SETUP
' ==================================================================================
Public Function ApplyPageSetup(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    With doc.PageSetup
        .TopMargin = Application.CentimetersToPoints(TOP_MARGIN_CM)
        .BottomMargin = Application.CentimetersToPoints(BOTTOM_MARGIN_CM)
        .LeftMargin = Application.CentimetersToPoints(LEFT_MARGIN_CM)
        .RightMargin = Application.CentimetersToPoints(RIGHT_MARGIN_CM)
        .HeaderDistance = Application.CentimetersToPoints(HEADER_DISTANCE_CM)
        .FooterDistance = Application.CentimetersToPoints(FOOTER_DISTANCE_CM)
        .Gutter = 0
        .Orientation = wdOrientPortrait
    End With
    ApplyPageSetup = True
    Exit Function
ErrorHandler:
    ApplyPageSetup = False
End Function

' ==================================================================================
' FONT FORMATTING (standard font for all paragraphs, image-safe)
' ==================================================================================
Public Function ApplyStdFont(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim para As Paragraph, hasInlineImage As Boolean
    Dim i As Long, formattedCount As Long, skippedCount As Long
    Dim paraFont As Font, needsFontFormatting As Boolean
    Dim needsUnderlineRemoval As Boolean, needsBoldRemoval As Boolean
    Dim paraFullText As String, isTitle As Boolean, isSpecialParagraph As Boolean
    Dim inlineShapesCount As Long

    For i = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(i)
        hasInlineImage = False: isTitle = False: needsUnderlineRemoval = False: needsBoldRemoval = False
        Set paraFont = para.Range.Font
        needsFontFormatting = (paraFont.Name <> STANDARD_FONT) Or (paraFont.Size <> STANDARD_FONT_SIZE) Or (paraFont.Color <> wdColorAutomatic)
        needsUnderlineRemoval = (paraFont.Underline <> wdUnderlineNone)
        needsBoldRemoval = (paraFont.Bold = True)
        inlineShapesCount = para.Range.InlineShapes.Count
        If Not needsFontFormatting And Not needsUnderlineRemoval And Not needsBoldRemoval And inlineShapesCount = 0 Then
            formattedCount = formattedCount + 1
            GoTo NextParagraph
        End If
        If inlineShapesCount > 0 Then
            hasInlineImage = True: skippedCount = skippedCount + 1
        End If
        If Not hasInlineImage And (needsFontFormatting Or needsUnderlineRemoval Or needsBoldRemoval) Then
            If HasVisualContent(para) Then hasInlineImage = True: skippedCount = skippedCount + 1
        End If
        If needsUnderlineRemoval Or needsBoldRemoval Then
            paraFullText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            If i <= 3 And para.Format.Alignment = wdAlignParagraphCenter And paraFullText <> "" Then isTitle = True
            Dim cleanParaText As String: cleanParaText = paraFullText
            Do While Len(cleanParaText) > 0 And (Right(cleanParaText, 1) = "." Or Right(cleanParaText, 1) = "," Or Right(cleanParaText, 1) = ":" Or Right(cleanParaText, 1) = ";")
                cleanParaText = Left(cleanParaText, Len(cleanParaText) - 1)
            Loop
            cleanParaText = Trim(LCase(cleanParaText))
            If cleanParaText = "justificativa:" Or IsAnexoPattern(cleanParaText) Then isSpecialParagraph = True
        End If
        If needsFontFormatting Then
            If Not hasInlineImage Then
                If SafeSetFont(para.Range, STANDARD_FONT, STANDARD_FONT_SIZE) Then
                    formattedCount = formattedCount + 1
                Else
                    With paraFont: .Name = STANDARD_FONT: .Size = STANDARD_FONT_SIZE: .Color = wdColorAutomatic: End With
                    formattedCount = formattedCount + 1
                End If
            Else
                FormatCharacterByCharacter para, STANDARD_FONT, STANDARD_FONT_SIZE, wdColorAutomatic, False, False
                formattedCount = formattedCount + 1
            End If
        End If
        If needsUnderlineRemoval Or needsBoldRemoval Then
            Dim removeUnderline As Boolean, removeBold As Boolean
            removeUnderline = needsUnderlineRemoval And Not isTitle
            removeBold = needsBoldRemoval And Not isTitle And Not isSpecialParagraph
            If removeUnderline Or removeBold Then
                If Not hasInlineImage Then
                    If removeUnderline Then paraFont.Underline = wdUnderlineNone
                    If removeBold Then paraFont.Bold = False
                Else
                    FormatCharacterByCharacter para, "", 0, 0, removeUnderline, removeBold
                End If
            End If
        End If
NextParagraph:
    Next i
    ApplyStdFont = True
    Exit Function
ErrorHandler:
    ApplyStdFont = False
End Function

Private Sub FormatCharacterByCharacter(para As Paragraph, fontName As String, fontSize As Long, fontColor As Long, removeUnderline As Boolean, removeBold As Boolean)
    On Error Resume Next
    Dim j As Long, charCount As Long, charRange As Range
    charCount = SafeGetCharacterCount(para.Range)
    If charCount > 0 Then
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
    End If
End Sub

' ==================================================================================
' PARAGRAPH FORMATTING (line spacing, indents, justification)
' ==================================================================================
Public Function ApplyStdParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim para As Paragraph, hasInlineImage As Boolean
    Dim paragraphIndent As Single, firstIndent As Single, rightMarginPoints As Single
    Dim i As Long, formattedCount As Long, skippedCount As Long, paraText As String
    rightMarginPoints = 0
    For i = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(i)
        hasInlineImage = False
        If para.Range.InlineShapes.Count > 0 Then hasInlineImage = True: skippedCount = skippedCount + 1
        If Not hasInlineImage And HasVisualContent(para) Then hasInlineImage = True: skippedCount = skippedCount + 1
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
        paraText = Trim(LCase(Replace(Replace(Replace(para.Range.Text, ".", ""), ",", ""), ";", "")))
        paraText = Replace(paraText, vbCr, "")
        paraText = Replace(paraText, vbLf, "")
        paraText = Replace(paraText, " ", "")
        With para.Format
            .LineSpacingRule = wdLineSpacingMultiple
            .LineSpacing = LINE_SPACING
            .RightIndent = rightMarginPoints
            .SpaceBefore = 0: .SpaceAfter = 0
            If para.Alignment = wdAlignParagraphCenter Then
                .LeftIndent = 0: .FirstLineIndent = 0
            Else
                firstIndent = .FirstLineIndent: paragraphIndent = .LeftIndent
                If paragraphIndent >= Application.CentimetersToPoints(5) Then
                    .LeftIndent = Application.CentimetersToPoints(9.5)
                ElseIf firstIndent < Application.CentimetersToPoints(5) Then
                    .LeftIndent = 0: .FirstLineIndent = Application.CentimetersToPoints(1.5)
                End If
            End If
        End With
        If para.Alignment = wdAlignParagraphLeft Then para.Alignment = wdAlignParagraphJustify
        formattedCount = formattedCount + 1
    Next i
    ApplyStdParagraphs = True
    Exit Function
ErrorHandler:
    ApplyStdParagraphs = False
End Function

' ==================================================================================
' FIRST & SECOND PARAGRAPH SPECIAL FORMATTING
' ==================================================================================
Public Function FormatFirstParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim para As Paragraph, paraText As String, i As Long, actualParaIndex As Long, firstParaIndex As Long
    actualParaIndex = 0: firstParaIndex = 0
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
            Dim n As Long, charCount4 As Long, charRange3 As Range
            charCount4 = SafeGetCharacterCount(para.Range)
            If charCount4 > 0 Then
                For n = 1 To charCount4
                    Set charRange3 = para.Range.Characters(n)
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
            .Alignment = wdAlignParagraphCenter
            .LeftIndent = 0: .FirstLineIndent = 0: .RightIndent = 0
        End With
    End If
    FormatFirstParagraph = True
    Exit Function
ErrorHandler:
    FormatFirstParagraph = False
End Function

Public Function CountBlankLinesBefore(doc As Document, paraIndex As Long) As Long
    On Error GoTo ErrorHandler
    Dim count As Long, i As Long, para As Paragraph, paraText As String
    count = 0
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
    CountBlankLinesBefore = count
    Exit Function
ErrorHandler:
    CountBlankLinesBefore = 0
End Function

Public Function CountBlankLinesAfter(doc As Document, paraIndex As Long) As Long
    On Error GoTo ErrorHandler
    Dim count As Long, i As Long, para As Paragraph, paraText As String
    count = 0
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
    CountBlankLinesAfter = count
    Exit Function
ErrorHandler:
    CountBlankLinesAfter = 0
End Function

Public Function FormatSecondParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim para As Paragraph, paraText As String, i As Long
    Dim actualParaIndex As Long, secondParaIndex As Long
    actualParaIndex = 0: secondParaIndex = 0
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
        Dim insertionPoint As Range, blankLinesBefore As Long
        Set insertionPoint = para.Range: insertionPoint.Collapse wdCollapseStart
        blankLinesBefore = CountBlankLinesBefore(doc, secondParaIndex)
        If blankLinesBefore < 2 Then
            Dim linesToAdd As Long, newLines As String
            linesToAdd = 2 - blankLinesBefore
            newLines = String(linesToAdd, vbCrLf)
            insertionPoint.InsertBefore newLines
            secondParaIndex = secondParaIndex + linesToAdd
            Set para = doc.Paragraphs(secondParaIndex)
        End If
        With para.Format
            .LeftIndent = Application.CentimetersToPoints(9)
            .FirstLineIndent = 0
            .RightIndent = 0
            .Alignment = wdAlignParagraphJustify
        End With
        Dim insertionPointAfter As Range, blankLinesAfter As Long
        Set insertionPointAfter = para.Range: insertionPointAfter.Collapse wdCollapseEnd
        blankLinesAfter = CountBlankLinesAfter(doc, secondParaIndex)
        If blankLinesAfter < 2 Then
            Dim linesToAddAfter As Long, newLinesAfter As String
            linesToAddAfter = 2 - blankLinesAfter
            newLinesAfter = String(linesToAddAfter, vbCrLf)
            insertionPointAfter.InsertAfter newLinesAfter
        End If
    End If
    FormatSecondParagraph = True
    Exit Function
ErrorHandler:
    FormatSecondParagraph = False
End Function

' ==================================================================================
' HYPHENATION
' ==================================================================================
Public Function EnableHyphenation(doc As Document) As Boolean
    On Error GoTo ErrHandler
    doc.Hyphenation = True
    EnableHyphenation = True
    Exit Function
ErrHandler:
    EnableHyphenation = True
End Function

' ==================================================================================
' WATERMARK REMOVAL
' ==================================================================================
Public Function RemoveWatermark(doc As Document) As Boolean
    On Error GoTo ErrorHandler
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
    RemoveWatermark = True
    Exit Function
ErrorHandler:
    RemoveWatermark = False
End Function

' ==================================================================================
' HEADER & FOOTER INSERTION
' ==================================================================================
Public Function InsertHeaderstamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim sec As Section, header As HeaderFooter, imgFile As String
    Dim imgWidth As Single, imgHeight As Single, shp As Shape, imgFound As Boolean
    Dim baseFolder As String
    imgFile = Trim(Config.headerImagePath)
    If Len(imgFile) = 0 Then
        On Error Resume Next
        baseFolder = IIf(doc.Path <> "", doc.Path, Environ("USERPROFILE") & "\Documents")
        On Error GoTo ErrorHandler
        If Right(baseFolder, 1) <> "\" Then baseFolder = baseFolder & "\"
        If Dir(baseFolder & "assets\stamp.png") <> "" Then
            imgFile = baseFolder & "assets\stamp.png"
        ElseIf Dir(Environ("USERPROFILE") & "\Documents\chainsaw\assets\stamp.png") <> "" Then
            imgFile = Environ("USERPROFILE") & "\Documents\chainsaw\assets\stamp.png"
        End If
    Else
        If InStr(1, imgFile, ":", vbTextCompare) = 0 And Left(imgFile, 2) <> "\\" Then
            On Error Resume Next
            baseFolder = IIf(doc.Path <> "", doc.Path, Environ("USERPROFILE") & "\Documents")
            On Error GoTo ErrorHandler
            If Right(baseFolder, 1) <> "\" Then baseFolder = baseFolder & "\"
            If Dir(baseFolder & imgFile) <> "" Then
                imgFile = baseFolder & imgFile
            ElseIf Dir(Environ("USERPROFILE") & "\Documents\chainsaw\" & imgFile) <> "" Then
                imgFile = Environ("USERPROFILE") & "\Documents\chainsaw\" & imgFile
            End If
        End If
    End If
    If Dir(imgFile) = "" Then
        Application.StatusBar = "Warning: Header image not found"
        InsertHeaderstamp = False
        Exit Function
    End If
    imgWidth = Application.CentimetersToPoints(HEADER_IMAGE_MAX_WIDTH_CM)
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
                    .Top = Application.CentimetersToPoints(HEADER_IMAGE_TOP_MARGIN_CM)
                    .WrapFormat.Type = wdWrapTopBottom
                    .ZOrder msoSendToBack
                End With
                imgFound = True
            End If
        End If
    Next sec
    InsertHeaderstamp = imgFound
    Exit Function
ErrorHandler:
    InsertHeaderstamp = False
End Function

Public Function InsertFooterstamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim sec As Section, footer As HeaderFooter, rng As Range
    For Each sec In doc.Sections
        Set footer = sec.Footers(wdHeaderFooterPrimary)
        If footer.Exists Then
            footer.LinkToPrevious = False
            Set rng = footer.Range
            rng.Delete
            Set rng = footer.Range: rng.Collapse Direction:=wdCollapseEnd: rng.Fields.Add Range:=rng, Type:=wdFieldPage
            Set rng = footer.Range: rng.Collapse Direction:=wdCollapseEnd: rng.Text = "-"
            Set rng = footer.Range: rng.Collapse Direction:=wdCollapseEnd: rng.Fields.Add Range:=rng, Type:=wdFieldNumPages
            With footer.Range
                .Font.Name = STANDARD_FONT
                .Font.Size = FOOTER_FONT_SIZE
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Fields.Update
            End With
        End If
    Next sec
    InsertFooterstamp = True
    Exit Function
ErrorHandler:
    InsertFooterstamp = False
End Function

' ==================================================================================
' CONSIDERANDO PARAGRAPH TOKEN FORMATTING
' ==================================================================================
Public Function FormatConsiderandoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim para As Paragraph, rawText As String, textNoCrLf As String
    Dim i As Long, totalFormatted As Long, startIdx As Long, n As Long, ch As String
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        rawText = para.Range.Text
        textNoCrLf = Replace(Replace(rawText, vbCr, ""), vbLf, "")
        If Len(textNoCrLf) >= 12 Then
            startIdx = 1
            For n = 1 To Len(textNoCrLf)
                ch = Mid$(textNoCrLf, n, 1)
                startIdx = n: Exit For
            Next n
            If startIdx + 11 <= Len(textNoCrLf) Then
                If LCase$(Mid$(textNoCrLf, startIdx, 12)) = "considerando" Then
                    Dim rng As Range
                    Set rng = para.Range.Duplicate
                    rng.SetRange rng.Start + (startIdx - 1), rng.Start + (startIdx - 1) + 12
                    rng.Text = "CONSIDERANDO"
                    rng.Font.Bold = True
                    totalFormatted = totalFormatted + 1
                End If
            End If
        End If
    Next i
    FormatConsiderandoParagraphs = True
    Exit Function
ErrorHandler:
    FormatConsiderandoParagraphs = False
End Function

' ==================================================================================
' TEXT REPLACEMENTS (GLOBAL + PARAGRAPH-SPECIFIC)
' ==================================================================================
Public Function ApplyTextReplacements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim rng As Range, replacementCount As Long, dOesteVariants() As String, i As Long
    Set rng = doc.Range
    ReDim dOesteVariants(0 To 15)
    dOesteVariants(0) = "d'O": dOesteVariants(1) = "d�O": dOesteVariants(2) = "d`O": dOesteVariants(3) = "d" & Chr(8220) & "O"
    dOesteVariants(4) = "d'o": dOesteVariants(5) = "d�o": dOesteVariants(6) = "d`o": dOesteVariants(7) = "d" & Chr(8220) & "o"
    dOesteVariants(8) = "D'O": dOesteVariants(9) = "D�O": dOesteVariants(10) = "D`O": dOesteVariants(11) = "D" & Chr(8220) & "O"
    dOesteVariants(12) = "D'o": dOesteVariants(13) = "D�o": dOesteVariants(14) = "D`o": dOesteVariants(15) = "D" & Chr(8220) & "o"
    For i = 0 To UBound(dOesteVariants)
        With rng.Find
            .ClearFormatting: .Replacement.ClearFormatting
            .Text = dOesteVariants(i) & "este"
            .Replacement.Text = "d'Oeste"
            .Forward = True: .Wrap = wdFindStop
            .Format = False: .MatchCase = False: .MatchWholeWord = False
            .MatchWildcards = False: .MatchSoundsLike = False: .MatchAllWordForms = False
            Do While .Execute(Replace:=wdReplaceOne)
                replacementCount = replacementCount + 1: rng.Collapse wdCollapseEnd
            Loop
        End With
    Next i
    Set rng = doc.Range
    Dim dashVariants() As String: ReDim dashVariants(0 To 2)
    dashVariants(0) = " - ": dashVariants(1) = " � ": dashVariants(2) = " � "
    For i = 0 To UBound(dashVariants)
        If dashVariants(i) <> " � " Then
            With rng.Find
                .ClearFormatting: .Replacement.ClearFormatting
                .Text = dashVariants(i): .Replacement.Text = " � "
                .Forward = True: .Wrap = wdFindStop
                .Format = False: .MatchCase = False: .MatchWholeWord = False
                .MatchWildcards = False: .MatchSoundsLike = False: .MatchAllWordForms = False
                Do While .Execute(Replace:=wdReplaceOne)
                    replacementCount = replacementCount + 1: rng.Collapse wdCollapseEnd
                Loop
            End With
        End If
    Next i
    Set rng = doc.Range
    Dim lineStartDashVariants() As String: ReDim lineStartDashVariants(0 To 1)
    lineStartDashVariants(0) = "^p- ": lineStartDashVariants(1) = "^p� "
    For i = 0 To UBound(lineStartDashVariants)
        With rng.Find
            .ClearFormatting: .Replacement.ClearFormatting
            .Text = lineStartDashVariants(i): .Replacement.Text = "^p� "
            .Forward = True: .Wrap = wdFindStop
            .Format = False: .MatchCase = False: .MatchWholeWord = False
            .MatchWildcards = False: .MatchSoundsLike = False: .MatchAllWordForms = False
            Do While .Execute(Replace:=wdReplaceOne)
                replacementCount = replacementCount + 1: rng.Collapse wdCollapseEnd
            Loop
        End With
    Next i
    Set rng = doc.Range
    Dim lineEndDashVariants() As String: ReDim lineEndDashVariants(0 To 1)
    lineEndDashVariants(0) = " -^p": lineEndDashVariants(1) = " �^p"
    For i = 0 To UBound(lineEndDashVariants)
        With rng.Find
            .ClearFormatting: .Replacement.ClearFormatting
            .Text = lineEndDashVariants(i): .Replacement.Text = " �^p"
            .Forward = True: .Wrap = wdFindStop
            .Format = False: .MatchCase = False: .MatchWholeWord = False
            .MatchWildcards = False: .MatchSoundsLike = False: .MatchAllWordForms = False
            Do While .Execute(Replace:=wdReplaceOne)
                replacementCount = replacementCount + 1: rng.Collapse wdCollapseEnd
            Loop
        End With
    Next i
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = "^l": .Replacement.Text = " "
        .Forward = True: .Wrap = wdFindStop
        .Format = False: .MatchCase = False: .MatchWholeWord = False
        .MatchWildcards = False: .MatchSoundsLike = False: .MatchAllWordForms = False
        Do While .Execute(Replace:=wdReplaceOne)
            replacementCount = replacementCount + 1: rng.Collapse wdCollapseEnd
        Loop
    End With
    On Error Resume Next
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = Chr(11): .Replacement.Text = " "
        .Forward = True: .Wrap = wdFindStop
        .Format = False: .MatchCase = False: .MatchWholeWord = False
        .MatchWildcards = False: .MatchSoundsLike = False: .MatchAllWordForms = False
        Do While .Execute(Replace:=wdReplaceOne)
            If Err.Number <> 0 Then Exit Do
            replacementCount = replacementCount + 1: rng.Collapse wdCollapseEnd
        Loop
    End With
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrorHandler
    On Error Resume Next
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = Chr(10): .Replacement.Text = " "
        .Forward = True: .Wrap = wdFindStop
        .Format = False: .MatchCase = False: .MatchWholeWord = False
        .MatchWildcards = False: .MatchSoundsLike = False: .MatchAllWordForms = False
        Do While .Execute(Replace:=wdReplaceOne)
            If Err.Number <> 0 Then Exit Do
            replacementCount = replacementCount + 1: rng.Collapse wdCollapseEnd
        Loop
    End With
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrorHandler
    ApplyTextReplacements = True
    Exit Function
ErrorHandler:
    ApplyTextReplacements = False
End Function

Public Function ApplySpecificParagraphReplacements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Application.StatusBar = "Applying specific replacements (per-paragraph and global)..."
    Dim replacementCount As Long, secondParaIndex As Long, thirdParaIndex As Long
    Dim para As Paragraph, paraText As String, i As Long, actualParaIndex As Long
    replacementCount = 0: actualParaIndex = 0: secondParaIndex = 0: thirdParaIndex = 0
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then
            actualParaIndex = actualParaIndex + 1
            If actualParaIndex = 2 Then
                secondParaIndex = i
            ElseIf actualParaIndex = 3 Then
                thirdParaIndex = i: Exit For
            End If
        End If
        If i > 50 Then Exit For
    Next i
    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.Count Then
        Set para = doc.Paragraphs(secondParaIndex)
        paraText = para.Range.Text
        Dim idxStart As Long: idxStart = 1
        Do While idxStart <= Len(paraText) And (Mid$(paraText, idxStart, 1) = " " Or Mid$(paraText, idxStart, 1) = vbTab)
            idxStart = idxStart + 1
        Loop
        If Len(paraText) >= idxStart + 6 Then
            Dim token As String: token = Mid$(paraText, idxStart, 7)
            If token = "Sugiro " Then
                Dim r1 As Range: Set r1 = para.Range.Duplicate
                r1.SetRange r1.Start + (idxStart - 1), r1.Start + (idxStart - 1) + 7
                r1.Text = "Requeiro ": replacementCount = replacementCount + 1
            ElseIf token = "Sugere " Then
                Dim r2 As Range: Set r2 = para.Range.Duplicate
                r2.SetRange r2.Start + (idxStart - 1), r2.Start + (idxStart - 1) + 7
                r2.Text = "Indica ": replacementCount = replacementCount + 1
            End If
        End If
    End If
    If thirdParaIndex > 0 And thirdParaIndex <= doc.Paragraphs.Count Then
        Set para = doc.Paragraphs(thirdParaIndex)
        paraText = para.Range.Text
        Dim originalText As String: originalText = paraText
        If InStr(paraText, " sugerir ") > 0 Then paraText = Replace(paraText, " sugerir ", " indicar "): replacementCount = replacementCount + 1
        If InStr(paraText, " Setor, ") > 0 Then paraText = Replace(paraText, " Setor, ", " setor competente, "): replacementCount = replacementCount + 1
        If paraText <> originalText Then para.Range.Text = paraText
    End If
    Dim rng As Range: Set rng = doc.Range
    With rng.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = " A C�MARA MUNICIPAL DE SANTA B�RBARA D'OESTE, ESTADO DE S�O PAULO "
        .Replacement.Text = " a C�mara Municipal de Santa B�rbara d'Oeste, estado de S�o Paulo, "
        .Forward = True: .Wrap = wdFindContinue
        .Format = False: .MatchCase = True: .MatchWholeWord = False: .MatchWildcards = False
        Do While .Execute(Replace:=wdReplaceOne)
            replacementCount = replacementCount + 1: rng.Collapse wdCollapseEnd
        Loop
    End With
    Dim wordsToUppercase() As String, j As Long
    ReDim wordsToUppercase(0 To 15)
    wordsToUppercase(0) = "aplaude": wordsToUppercase(1) = "Aplaude": wordsToUppercase(2) = "aplauso": wordsToUppercase(3) = "Aplauso"
    wordsToUppercase(4) = "protesta": wordsToUppercase(5) = "Protesta": wordsToUppercase(6) = "protesto": wordsToUppercase(7) = "Protesto"
    wordsToUppercase(8) = "apela": wordsToUppercase(9) = "Apela": wordsToUppercase(10) = "apelo": wordsToUppercase(11) = "Apelo"
    wordsToUppercase(12) = "apoia": wordsToUppercase(13) = "Apoia": wordsToUppercase(14) = "apoio": wordsToUppercase(15) = "Apoio"
    For j = 0 To UBound(wordsToUppercase)
        Set rng = doc.Range
        With rng.Find
            .ClearFormatting: .Replacement.ClearFormatting
            .Text = wordsToUppercase(j): .Replacement.Text = UCase(wordsToUppercase(j))
            .Forward = True: .Wrap = wdFindContinue
            .Format = False: .MatchCase = True: .MatchWholeWord = True: .MatchWildcards = False
            Do While .Execute(Replace:=wdReplaceOne)
                replacementCount = replacementCount + 1: rng.Collapse wdCollapseEnd
            Loop
        End With
    Next j
    ApplySpecificParagraphReplacements = True
    Exit Function
ErrorHandler:
    ApplySpecificParagraphReplacements = False
End Function

' ==================================================================================
' NUMBERED PARAGRAPH NORMALIZATION
' ==================================================================================
Public Function FormatNumberedParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim para As Paragraph, paraText As String, i As Long, formattedCount As Long
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            If IsNumberedParagraph(paraText) Then
                With para.Range.ListFormat
                    .RemoveNumbers
                    .ApplyNumberDefault
                End With
                Dim cleanedText As String
                cleanedText = RemoveManualNumber(paraText)
                para.Range.Text = cleanedText & vbCrLf
                formattedCount = formattedCount + 1
            End If
        End If
    Next i
    FormatNumberedParagraphs = True
    Exit Function
ErrorHandler:
    FormatNumberedParagraphs = False
End Function

' ==================================================================================
' JUSTIFICATIVA / ANEXO PARAGRAPH FORMATTING
' ==================================================================================
Public Function FormatJustificativaAnexoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    Dim para As Paragraph, paraText As String, cleanText As String, i As Long, formattedCount As Long
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
            cleanText = paraText
            Do While Len(cleanText) > 0 And (Right(cleanText, 1) = "." Or Right(cleanText, 1) = "," Or Right(cleanText, 1) = ":" Or Right(cleanText, 1) = ";")
                cleanText = Left(cleanText, Len(cleanText) - 1)
            Loop
            cleanText = Trim(LCase(cleanText))
            If cleanText = "justificativa" Then
                With para.Format
                    .LeftIndent = 0: .FirstLineIndent = 0: .Alignment = wdAlignParagraphCenter
                End With
                With para.Range.Font: .Bold = True: End With
                Dim originalEnd As String: originalEnd = ""
                If Len(paraText) > Len(cleanText) Then originalEnd = Right(paraText, Len(paraText) - Len(cleanText))
                para.Range.Text = "Justificativa" & originalEnd & vbCrLf
                formattedCount = formattedCount + 1
            ElseIf IsAnexoPattern(cleanText) Then
                With para.Format
                    .LeftIndent = 0: .FirstLineIndent = 0: .RightIndent = 0: .Alignment = wdAlignParagraphLeft
                End With
                With para.Range.Font: .Bold = True: End With
                Dim anexoEnd As String: anexoEnd = ""
                If Len(paraText) > Len(cleanText) Then anexoEnd = Right(paraText, Len(paraText) - Len(cleanText))
                Dim anexoText As String
                If cleanText = "anexo" Then anexoText = "Anexo" Else anexoText = "Anexos"
                para.Range.Text = anexoText & anexoEnd & vbCrLf
                formattedCount = formattedCount + 1
            ElseIf IsAnteOExpostoPattern(paraText) Then
                With para.Range.Font: .Bold = True: End With
                formattedCount = formattedCount + 1
            End If
        End If
    Next i
    FormatJustificativaAnexoParagraphs = True
    Exit Function
ErrorHandler:
    FormatJustificativaAnexoParagraphs = False
End Function

' (End of modFormatting)
'==================================================================================
' INTERNAL HELPER PATTERN & NUMBERING FUNCTIONS (migrated from modOrquestration)
'==================================================================================
Private Function IsAnexoPattern(textValue As String) As Boolean
    On Error GoTo ErrHandler
    Dim normalizedText As String
    normalizedText = LCase$(Trim$(textValue))
    IsAnexoPattern = (normalizedText = "anexo" Or normalizedText = "anexos" Or _
                      normalizedText = "anexo:" Or normalizedText = "anexos:")
    Exit Function
ErrHandler:
    IsAnexoPattern = False
End Function

Private Function IsAnteOExpostoPattern(textValue As String) As Boolean
    On Error GoTo ErrHandler
    Dim normalizedText As String
    normalizedText = LCase$(Trim$(textValue))
    ' Accept slight punctuation variants
    If Len(normalizedText) = 0 Then
        IsAnteOExpostoPattern = False
    ElseIf Left$(normalizedText, 13) = "ante o exposto" Then
        IsAnteOExpostoPattern = True
    Else
        IsAnteOExpostoPattern = False
    End If
    Exit Function
ErrHandler:
    IsAnteOExpostoPattern = False
End Function

Private Function IsNumberedParagraph(text As String) As Boolean
    On Error GoTo ErrHandler
    Dim cleanText As String, firstToken As String, spacePos As Long, numberPart As String, lastChar As String
    cleanText = Trim$(text)
    If Len(cleanText) = 0 Then GoTo NotPattern
    spacePos = InStr(cleanText, " ")
    If spacePos > 0 Then
        firstToken = Left$(cleanText, spacePos - 1)
    Else
        firstToken = cleanText
    End If
    ' Pattern 1: 1.
    If Len(firstToken) >= 2 And Right$(firstToken, 1) = "." Then
        numberPart = Left$(firstToken, Len(firstToken) - 1)
        If IsNumeric(numberPart) And Val(numberPart) > 0 Then If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then IsNumberedParagraph = True: Exit Function
    End If
    ' Pattern 2: 1)
    If Len(firstToken) >= 2 And Right$(firstToken, 1) = ")" Then
        numberPart = Left$(firstToken, Len(firstToken) - 1)
        If IsNumeric(numberPart) And Val(numberPart) > 0 Then If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then IsNumberedParagraph = True: Exit Function
    End If
    ' Pattern 3: (1)
    If Len(firstToken) >= 3 And Left$(firstToken, 1) = "(" And Right$(firstToken, 1) = ")" Then
        numberPart = Mid$(firstToken, 2, Len(firstToken) - 2)
        If IsNumeric(numberPart) And Val(numberPart) > 0 Then If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then IsNumberedParagraph = True: Exit Function
    End If
    ' Pattern 4: 1 <space>
    If IsNumeric(firstToken) And Val(firstToken) > 0 And spacePos > 0 Then
        If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then IsNumberedParagraph = True: Exit Function
    End If
    ' Pattern 5: 1- / 1: / 1;
    If Len(firstToken) >= 2 Then
        lastChar = Right$(firstToken, 1)
        If lastChar = "-" Or lastChar = ":" Or lastChar = ";" Then
            numberPart = Left$(firstToken, Len(firstToken) - 1)
            If IsNumeric(numberPart) And Val(numberPart) > 0 Then
                If HasSubstantiveTextAfterNumber(cleanText, firstToken) Then IsNumberedParagraph = True: Exit Function
            End If
        End If
    End If
NotPattern:
    IsNumberedParagraph = False
    Exit Function
ErrHandler:
    IsNumberedParagraph = False
End Function

Private Function HasSubstantiveTextAfterNumber(fullText As String, numberToken As String) As Boolean
    Dim startPos As Long, remainingText As String, words() As String, i As Long
    startPos = Len(numberToken) + 1
    If startPos > Len(fullText) Then HasSubstantiveTextAfterNumber = False: Exit Function
    remainingText = Trim$(Mid$(fullText, startPos))
    If Len(remainingText) = 0 Then HasSubstantiveTextAfterNumber = False: Exit Function
    words = Split(remainingText, " ")
    For i = 0 To UBound(words)
        If ContainsLetters(words(i)) And Len(words(i)) >= 2 Then HasSubstantiveTextAfterNumber = True: Exit Function
    Next i
    HasSubstantiveTextAfterNumber = False
End Function

Private Function ContainsLetters(text As String) As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(text)
        ch = LCase$(Mid$(text, i, 1))
        If ch >= "a" And ch <= "z" Then ContainsLetters = True: Exit Function
    Next i
    ContainsLetters = False
End Function

Private Function RemoveManualNumber(text As String) As String
    Dim cleanText As String, spacePos As Long, firstToken As String
    cleanText = Trim$(text)
    If Len(cleanText) = 0 Then RemoveManualNumber = text: Exit Function
    spacePos = InStr(cleanText, " ")
    If spacePos > 0 Then
        firstToken = Left$(cleanText, spacePos - 1)
        If (Len(firstToken) >= 2 And (Right$(firstToken, 1) = "." Or Right$(firstToken, 1) = ")")) Or _
           (Len(firstToken) >= 3 And Left$(firstToken, 1) = "(" And Right$(firstToken, 1) = ")") Or _
           (IsNumeric(firstToken) And Val(firstToken) > 0) Then
            RemoveManualNumber = Trim$(Mid$(cleanText, spacePos + 1))
        Else
            RemoveManualNumber = cleanText
        End If
    Else
        RemoveManualNumber = cleanText
    End If
End Function

