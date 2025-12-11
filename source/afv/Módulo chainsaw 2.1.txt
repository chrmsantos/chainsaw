Public Sub RemoverLinhasEmBrancoExtras(doc As Document)
    Dim i As Long

      ' --- Espaçamento simples em todos os parágrafos ---
    Dim p As Paragraph
    For Each p In doc.Paragraphs
        With p.Format
            .LineSpacingRule = wdLineSpaceSingle
            .LineSpacing = 12
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
    Next p

    ' --- Remove linhas em branco extras ---
    For i = doc.Paragraphs.count To 2 Step -1
        Dim txtAtual As String, txtAnterior As String
        txtAtual = Trim(Replace(doc.Paragraphs(i).Range.text, vbCr, ""))
        txtAnterior = Trim(Replace(doc.Paragraphs(i - 1).Range.text, vbCr, ""))
        
        If txtAtual = "" And txtAnterior = "" Then
            doc.Paragraphs(i).Range.Delete
        End If
    Next i

    ' --- Substituições no texto ---
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False

        .text = "por intermédio do Setor,"
        .Replacement.text = "por intermédio do Setor competente,"
        .Execute Replace:=wdReplaceAll

        .text = "Indica ao Poder Executivo Municipal efetue"
        .Replacement.text = "Indica ao Poder Executivo Municipal que efetue"
        .Execute Replace:=wdReplaceAll

        .text = "Fomos procurados por munícipes, solicitando essa providência, pois segundo eles,"
        .Replacement.text = "Fomos procurados por munícipes solicitando essa providência, pois, segundo eles,"
        .Execute Replace:=wdReplaceAll
    End With

    ' --- Ajustes por parágrafo ---
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        Dim cleanTxt As String
        cleanTxt = LCase(Trim(Replace(para.Range.text, vbCr, "")))
        cleanTxt = Replace(cleanTxt, "-", "")

        ' Espaçamento extra antes e depois da data
        If InStr(cleanTxt, "plenário ""dr. tancredo neves""") > 0 Then
            para.Format.SpaceBefore = 24
            para.Format.SpaceAfter = 24
        End If

        ' Centraliza nome, cargo e partido
        If Left(cleanTxt, 8) = "vereador" _
           Or Left(cleanTxt, 9) = "vereadora" _
           Or InStr(cleanTxt, "vicepresidente") > 0 Then

            ' Cargo
            With para.Format
                .leftIndent = 0
                .RightIndent = 0
                .firstLineIndent = 0
                .alignment = wdAlignParagraphCenter
            End With

            ' Nome (parágrafo anterior)
            If Not para.Previous Is Nothing Then
                With para.Previous.Format
                    .leftIndent = 0
                    .RightIndent = 0
                    .firstLineIndent = 0
                    .alignment = wdAlignParagraphCenter
                End With
                para.Previous.Range.Font.Bold = True
            End If

            ' Partido (parágrafo seguinte)
            If Not para.Next Is Nothing Then
                With para.Next.Format
                    .leftIndent = 0
                    .RightIndent = 0
                    .firstLineIndent = 0
                    .alignment = wdAlignParagraphCenter
                End With
            End If
        End If
    Next para
End Sub

