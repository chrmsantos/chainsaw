'================================================================================
' INSERÇÃO DE NÚMEROS DE PÁGINA NO RODAPÉ + SIGLA AFV
'================================================================================
Private Function InsertFooterStamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As Section
    Dim footer As HeaderFooter
    Dim rngAFV As Range
    Dim rngPage As Range
    Dim rngDash As Range
    Dim rngNum As Range
    Dim fPage As Field
    Dim fTotal As Field

    For Each sec In doc.Sections
        Set footer = sec.Footers(wdHeaderFooterPrimary)
        If footer.Exists Then
            footer.LinkToPrevious = False

            ' -------------------------------
            ' Limpa todo o rodapé
            ' -------------------------------
            footer.Range.Delete

            ' -------------------------------
            ' Insere "afv" à esquerda
            ' -------------------------------
            Set rngAFV = footer.Range
            rngAFV.Collapse Direction:=wdCollapseStart
            rngAFV.text = "afv"
            With rngAFV.Font
                .Name = "Arial"
                .size = 6
                .Color = RGB(128, 128, 128)
            End With
            rngAFV.ParagraphFormat.alignment = wdAlignParagraphLeft
            rngAFV.InsertParagraphAfter

            ' -------------------------------
            ' Insere números de página X-Y centralizados
            ' -------------------------------
            ' Cria um parágrafo limpo
            Set rngPage = footer.Range.Paragraphs.Last.Range
            rngPage.text = ""
            rngPage.Collapse Direction:=wdCollapseStart
           
            
            rngPage.text = "Página "
With rngPage.Font
    .Name = "Arial"
    .size = 9
End With
rngPage.Collapse Direction:=wdCollapseEnd


            ' Campo PAGE
            Set fPage = rngPage.Fields.Add(Range:=rngPage, Type:=wdFieldPage)

            ' De "de"
            Set rngDash = footer.Range.Paragraphs.Last.Range
            rngDash.Collapse Direction:=wdCollapseEnd
            rngDash.text = " de "

            ' Campo NUMPAGES
            Set rngNum = footer.Range.Paragraphs.Last.Range
            rngNum.Collapse Direction:=wdCollapseEnd
            Set fTotal = rngNum.Fields.Add(Range:=rngNum, Type:=wdFieldNumPages)

            ' Centraliza os números de página
            rngPage.ParagraphFormat.alignment = wdAlignParagraphCenter
            
            ' Formata os campos de número de página
            With fPage.result
             .Font.Name = "Arial"
            .Font.size = 9
            End With

            With fTotal.result
             .Font.Name = "Arial"
             .Font.size = 9
            End With

        End If
    Next sec

    InsertFooterStamp = True
    Exit Function

ErrorHandler:
    InsertFooterStamp = False
End Function

