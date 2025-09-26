' =============================================================================
' CHAINSAW PROPOSITURAS - VERSÃO SIMPLIFICADA
' =============================================================================
' Sistema básico de padronização de documentos legislativos no Microsoft Word
' Versão: 2.0-Simple | Data: 2025-09-26

Option Explicit

' =============================================================================
' FUNÇÃO PRINCIPAL - PADRONIZAÇÃO DE DOCUMENTO
' =============================================================================
Public Sub PadronizarDocumento()
    On Error GoTo TratarErro
    
    ' Verificar se há documento ativo
    If ActiveDocument Is Nothing Then
        MsgBox "Abra um documento antes de executar a padronização.", vbExclamation, "Chainsaw - Erro"
        Exit Sub
    End If
    
    ' Confirmar execução
    If MsgBox("Deseja padronizar o documento atual?" & vbCrLf & _
              "Esta ação irá modificar a formatação.", vbYesNo + vbQuestion, "Chainsaw - Confirmar") = vbNo Then
        Exit Sub
    End If
    
    ' Executar padronização
    Application.StatusBar = "Padronizando documento..."
    Call ExecutarPadronizacao
    
    ' Finalizar
    Application.StatusBar = "Padronização concluída!"
    MsgBox "Documento padronizado com sucesso!", vbInformation, "Chainsaw - Concluído"
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro durante a padronização: " & Err.Description, vbCritical, "Chainsaw - Erro"
    Application.StatusBar = ""
End Sub

' =============================================================================
' EXECUTAR PADRONIZAÇÃO - FUNCIONALIDADE PRINCIPAL
' =============================================================================
Private Sub ExecutarPadronizacao()
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Desabilitar atualizações de tela para performance
    Application.ScreenUpdating = False
    
    ' 1. Limpar formatação desnecessária
    Call LimparFormatacao(doc)
    
    ' 2. Aplicar formatação da primeira linha
    Call FormatarPrimeiraLinha(doc)
    
    ' 3. Aplicar formatação dos parágrafos 2-4
    Call FormatarParagrafos2a4(doc)
    
    ' 4. Padronizar "Considerando"
    Call PadronizarConsiderando(doc)
    
    ' 5. Limpar espaços e quebras excessivas
    Call LimparEspacos(doc)
    
    ' Reabilitar atualizações de tela
    Application.ScreenUpdating = True
End Sub

' =============================================================================
' LIMPAR FORMATAÇÃO DESNECESSÁRIA
' =============================================================================
Private Sub LimparFormatacao(doc As Document)
    Dim para As Paragraph
    
    ' Remover formatação excessiva de todos os parágrafos
    For Each para In doc.Paragraphs
        With para.Range.Font
            .Bold = False
            .Italic = False
            .Underline = wdUnderlineNone
            .Size = 12
            .Name = "Times New Roman"
        End With
        
        With para.Range.ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LeftIndent = 0
            .FirstLineIndent = 0
            .LineSpacing = LinesToPoints(1)
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
    Next para
End Sub

' =============================================================================
' FORMATAR PRIMEIRA LINHA
' =============================================================================
Private Sub FormatarPrimeiraLinha(doc As Document)
    If doc.Paragraphs.Count >= 1 Then
        With doc.Paragraphs(1).Range
            .Font.Bold = True
            .Font.Underline = wdUnderlineSingle
            .Text = UCase(.Text)
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
        End With
    End If
End Sub

' =============================================================================
' FORMATAR PARÁGRAFOS 2-4 (RECUO 9CM)
' =============================================================================
Private Sub FormatarParagrafos2a4(doc As Document)
    Dim i As Integer
    
    For i = 2 To 4
        If doc.Paragraphs.Count >= i Then
            With doc.Paragraphs(i).Range.ParagraphFormat
                .LeftIndent = CentimetersToPoints(9)
                .FirstLineIndent = 0
            End With
        End If
    Next i
End Sub

' =============================================================================
' PADRONIZAR "CONSIDERANDO"
' =============================================================================
Private Sub PadronizarConsiderando(doc As Document)
    Dim para As Paragraph
    
    For Each para In doc.Paragraphs
        If InStr(1, para.Range.Text, "considerando", vbTextCompare) > 0 Then
            ' Aplicar formatação especial para "Considerando"
            If InStr(1, Trim(para.Range.Text), "considerando", vbTextCompare) = 1 Then
                para.Range.Text = Replace(para.Range.Text, "considerando", "CONSIDERANDO", 1, 1, vbTextCompare)
                para.Range.Font.Bold = True
            End If
        End If
    Next para
End Sub

' =============================================================================
' LIMPAR ESPAÇOS E QUEBRAS EXCESSIVAS
' =============================================================================
Private Sub LimparEspacos(doc As Document)
    ' Remover múltiplos espaços
    With doc.Range.Find
        .ClearFormatting
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Remover múltiplas quebras de linha
    With doc.Range.Find
        .ClearFormatting
        .Text = "^p^p^p"
        .Replacement.Text = "^p^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
End Sub

' =============================================================================
' FUNÇÃO AUXILIAR - TESTE SIMPLES
' =============================================================================
Public Sub Teste()
    MsgBox "Chainsaw Proposituras - Versão Simplificada funcionando!" & vbCrLf & _
           "Execute PadronizarDocumento para usar o sistema.", vbInformation, "Teste"
End Sub