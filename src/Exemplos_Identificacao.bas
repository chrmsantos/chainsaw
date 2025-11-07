' =============================================================================
' EXEMPLOS DE USO - Sistema de Identificação de Elementos Estruturais
' =============================================================================
' Versão: 1.0
' Data: 07/11/2024
' =============================================================================
' INSTRUÇÕES:
' - Este arquivo contém exemplos de uso das funções de identificação
' - Cole os exemplos no editor VBA do Word para testá-los
' - Execute a macro PadronizarDocumentoMain primeiro para construir o cache
' =============================================================================

Option Explicit

'================================================================================
' EXEMPLO 1: Exibir informações sobre todos os elementos
'================================================================================
Sub Exemplo1_ExibirInformacoesCompletas()
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Exibe relatório completo
    Dim info As String
    info = GetElementInfo(doc)
    
    MsgBox info, vbInformation, "Estrutura do Documento"
End Sub

'================================================================================
' EXEMPLO 2: Selecionar e destacar o título
'================================================================================
Sub Exemplo2_SelecionarTitulo()
    Dim doc As Document
    Dim titulo As Range
    
    Set doc = ActiveDocument
    Set titulo = GetTituloRange(doc)
    
    If Not titulo Is Nothing Then
        titulo.Select
        MsgBox "Título selecionado!", vbInformation
    Else
        MsgBox "Título não encontrado!", vbExclamation
    End If
End Sub

'================================================================================
' EXEMPLO 3: Contar palavras de cada elemento
'================================================================================
Sub Exemplo3_ContarPalavrasPorElemento()
    Dim doc As Document
    Dim relatorio As String
    Dim rng As Range
    
    Set doc = ActiveDocument
    relatorio = "=== CONTAGEM DE PALAVRAS ===" & vbCrLf & vbCrLf
    
    ' Título
    Set rng = GetTituloRange(doc)
    If Not rng Is Nothing Then
        relatorio = relatorio & "Título: " & rng.Words.Count & " palavras" & vbCrLf
    End If
    
    ' Ementa
    Set rng = GetEmentaRange(doc)
    If Not rng Is Nothing Then
        relatorio = relatorio & "Ementa: " & rng.Words.Count & " palavras" & vbCrLf
    End If
    
    ' Proposição
    Set rng = GetProposicaoRange(doc)
    If Not rng Is Nothing Then
        relatorio = relatorio & "Proposição: " & rng.Words.Count & " palavras" & vbCrLf
    End If
    
    ' Justificativa
    Set rng = GetJustificativaRange(doc)
    If Not rng Is Nothing Then
        relatorio = relatorio & "Justificativa: " & rng.Words.Count & " palavras" & vbCrLf
    End If
    
    ' Assinatura
    Set rng = GetAssinaturaRange(doc)
    If Not rng Is Nothing Then
        relatorio = relatorio & "Assinatura: " & rng.Words.Count & " palavras" & vbCrLf
    End If
    
    ' Anexo
    Set rng = GetAnexoRange(doc)
    If Not rng Is Nothing Then
        relatorio = relatorio & "Anexo: " & rng.Words.Count & " palavras" & vbCrLf
    End If
    
    relatorio = relatorio & vbCrLf & "=============================="
    
    MsgBox relatorio, vbInformation, "Contagem de Palavras"
End Sub

'================================================================================
' EXEMPLO 4: Exportar proposição para novo documento
'================================================================================
Sub Exemplo4_ExportarProposicao()
    Dim docOriginal As Document
    Dim docNovo As Document
    Dim proposicao As Range
    
    Set docOriginal = ActiveDocument
    Set proposicao = GetProposicaoRange(docOriginal)
    
    If Not proposicao Is Nothing Then
        ' Cria novo documento
        Set docNovo = Documents.Add
        
        ' Copia a proposição
        proposicao.Copy
        docNovo.Range.Paste
        
        MsgBox "Proposição exportada para novo documento!", vbInformation
    Else
        MsgBox "Proposição não encontrada!", vbExclamation
    End If
End Sub

'================================================================================
' EXEMPLO 5: Adicionar marcadores de seção
'================================================================================
Sub Exemplo5_AdicionarMarcadoresSecao()
    Dim doc As Document
    Dim rng As Range
    
    Set doc = ActiveDocument
    
    ' Desabilita atualização de tela
    Application.ScreenUpdating = False
    
    ' Título
    Set rng = GetTituloRange(doc)
    If Not rng Is Nothing Then
        rng.Bookmarks.Add "secao_titulo"
    End If
    
    ' Ementa
    Set rng = GetEmentaRange(doc)
    If Not rng Is Nothing Then
        rng.Bookmarks.Add "secao_ementa"
    End If
    
    ' Proposição
    Set rng = GetProposicaoRange(doc)
    If Not rng Is Nothing Then
        rng.Bookmarks.Add "secao_proposicao"
    End If
    
    ' Justificativa
    Set rng = GetJustificativaRange(doc)
    If Not rng Is Nothing Then
        rng.Bookmarks.Add "secao_justificativa"
    End If
    
    ' Assinatura
    Set rng = GetAssinaturaRange(doc)
    If Not rng Is Nothing Then
        rng.Bookmarks.Add "secao_assinatura"
    End If
    
    ' Anexo
    Set rng = GetAnexoRange(doc)
    If Not rng Is Nothing Then
        rng.Bookmarks.Add "secao_anexo"
    End If
    
    Application.ScreenUpdating = True
    
    MsgBox "Marcadores de seção adicionados com sucesso!" & vbCrLf & _
           "Use Ctrl+G para navegar pelos marcadores.", vbInformation
End Sub

'================================================================================
' EXEMPLO 6: Validar estrutura do documento
'================================================================================
Sub Exemplo6_ValidarEstrutura()
    Dim doc As Document
    Dim relatorio As String
    Dim erros As Long
    
    Set doc = ActiveDocument
    erros = 0
    
    relatorio = "=== VALIDAÇÃO DA ESTRUTURA ===" & vbCrLf & vbCrLf
    
    ' Verifica título
    If GetTituloRange(doc) Is Nothing Then
        relatorio = relatorio & "[X] Título não encontrado" & vbCrLf
        erros = erros + 1
    Else
        relatorio = relatorio & "[OK] Título encontrado" & vbCrLf
    End If
    
    ' Verifica ementa
    If GetEmentaRange(doc) Is Nothing Then
        relatorio = relatorio & "[X] Ementa não encontrada" & vbCrLf
        erros = erros + 1
    Else
        relatorio = relatorio & "[OK] Ementa encontrada" & vbCrLf
    End If
    
    ' Verifica proposição
    If GetProposicaoRange(doc) Is Nothing Then
        relatorio = relatorio & "[X] Proposição não encontrada" & vbCrLf
        erros = erros + 1
    Else
        relatorio = relatorio & "[OK] Proposição encontrada" & vbCrLf
    End If
    
    ' Verifica justificativa
    If GetTituloJustificativaRange(doc) Is Nothing Then
        relatorio = relatorio & "[X] Título da Justificativa não encontrado" & vbCrLf
        erros = erros + 1
    Else
        relatorio = relatorio & "[OK] Título da Justificativa encontrado" & vbCrLf
    End If
    
    If GetJustificativaRange(doc) Is Nothing Then
        relatorio = relatorio & "[X] Justificativa não encontrada" & vbCrLf
        erros = erros + 1
    Else
        relatorio = relatorio & "[OK] Justificativa encontrada" & vbCrLf
    End If
    
    ' Verifica data
    If GetDataRange(doc) Is Nothing Then
        relatorio = relatorio & "[X] Data (Plenário) não encontrada" & vbCrLf
        erros = erros + 1
    Else
        relatorio = relatorio & "[OK] Data (Plenário) encontrada" & vbCrLf
    End If
    
    ' Verifica assinatura
    If GetAssinaturaRange(doc) Is Nothing Then
        relatorio = relatorio & "[X] Assinatura não encontrada" & vbCrLf
        erros = erros + 1
    Else
        relatorio = relatorio & "[OK] Assinatura encontrada" & vbCrLf
    End If
    
    ' Anexo é opcional
    If Not GetAnexoRange(doc) Is Nothing Then
        relatorio = relatorio & "[OK] Anexo encontrado (opcional)" & vbCrLf
    Else
        relatorio = relatorio & "[--] Anexo não presente (opcional)" & vbCrLf
    End If
    
    relatorio = relatorio & vbCrLf & "==============================" & vbCrLf
    
    If erros = 0 Then
        relatorio = relatorio & "Estrutura válida! ✓"
        MsgBox relatorio, vbInformation, "Validação de Estrutura"
    Else
        relatorio = relatorio & erros & " erro(s) encontrado(s)!"
        MsgBox relatorio, vbExclamation, "Validação de Estrutura"
    End If
End Sub

'================================================================================
' EXEMPLO 7: Aplicar cor de fundo aos elementos (para debug visual)
'================================================================================
Sub Exemplo7_DestacaElementosVisualmente()
    Dim doc As Document
    Dim rng As Range
    
    Set doc = ActiveDocument
    
    ' Desabilita atualização de tela
    Application.ScreenUpdating = False
    
    ' Limpa destaques anteriores
    doc.Range.HighlightColorIndex = 0
    
    ' Título - Amarelo
    Set rng = GetTituloRange(doc)
    If Not rng Is Nothing Then
        rng.HighlightColorIndex = 7 ' Amarelo
    End If
    
    ' Ementa - Verde claro
    Set rng = GetEmentaRange(doc)
    If Not rng Is Nothing Then
        rng.HighlightColorIndex = 4 ' Verde claro
    End If
    
    ' Proposição - Azul claro
    Set rng = GetProposicaoRange(doc)
    If Not rng Is Nothing Then
        rng.HighlightColorIndex = 5 ' Azul claro
    End If
    
    ' Justificativa - Rosa
    Set rng = GetJustificativaRange(doc)
    If Not rng Is Nothing Then
        rng.HighlightColorIndex = 6 ' Rosa
    End If
    
    ' Data - Cinza
    Set rng = GetDataRange(doc)
    If Not rng Is Nothing Then
        rng.HighlightColorIndex = 15 ' Cinza
    End If
    
    ' Assinatura - Laranja
    Set rng = GetAssinaturaRange(doc)
    If Not rng Is Nothing Then
        rng.HighlightColorIndex = 14 ' Laranja
    End If
    
    ' Anexo - Verde escuro
    Set rng = GetAnexoRange(doc)
    If Not rng Is Nothing Then
        rng.HighlightColorIndex = 11 ' Verde escuro
    End If
    
    Application.ScreenUpdating = True
    
    MsgBox "Elementos destacados visualmente!" & vbCrLf & vbCrLf & _
           "Amarelo: Título" & vbCrLf & _
           "Verde claro: Ementa" & vbCrLf & _
           "Azul claro: Proposição" & vbCrLf & _
           "Rosa: Justificativa" & vbCrLf & _
           "Cinza: Data" & vbCrLf & _
           "Laranja: Assinatura" & vbCrLf & _
           "Verde escuro: Anexo", vbInformation, "Elementos Destacados"
End Sub

'================================================================================
' EXEMPLO 8: Remover destaques visuais
'================================================================================
Sub Exemplo8_RemoverDestaques()
    ActiveDocument.Range.HighlightColorIndex = 0
    MsgBox "Destaques removidos!", vbInformation
End Sub

'================================================================================
' EXEMPLO 9: Gerar índice dos elementos
'================================================================================
Sub Exemplo9_GerarIndice()
    Dim doc As Document
    Dim docNovo As Document
    Dim rng As Range
    Dim indice As String
    
    Set doc = ActiveDocument
    
    indice = "ÍNDICE DOS ELEMENTOS ESTRUTURAIS" & vbCrLf
    indice = indice & String(50, "=") & vbCrLf & vbCrLf
    
    ' Título
    Set rng = GetTituloRange(doc)
    If Not rng Is Nothing Then
        indice = indice & "1. TÍTULO" & vbCrLf
        indice = indice & "   " & Left(rng.Text, 80) & "..." & vbCrLf & vbCrLf
    End If
    
    ' Ementa
    Set rng = GetEmentaRange(doc)
    If Not rng Is Nothing Then
        indice = indice & "2. EMENTA" & vbCrLf
        indice = indice & "   " & Left(rng.Text, 80) & "..." & vbCrLf & vbCrLf
    End If
    
    ' Proposição
    Set rng = GetProposicaoRange(doc)
    If Not rng Is Nothing Then
        indice = indice & "3. PROPOSIÇÃO" & vbCrLf
        indice = indice & "   Parágrafos: " & rng.Paragraphs.Count & vbCrLf
        indice = indice & "   Palavras: " & rng.Words.Count & vbCrLf & vbCrLf
    End If
    
    ' Justificativa
    Set rng = GetJustificativaRange(doc)
    If Not rng Is Nothing Then
        indice = indice & "4. JUSTIFICATIVA" & vbCrLf
        indice = indice & "   Parágrafos: " & rng.Paragraphs.Count & vbCrLf
        indice = indice & "   Palavras: " & rng.Words.Count & vbCrLf & vbCrLf
    End If
    
    ' Data
    Set rng = GetDataRange(doc)
    If Not rng Is Nothing Then
        indice = indice & "5. DATA (PLENÁRIO)" & vbCrLf
        indice = indice & "   " & Left(rng.Text, 80) & vbCrLf & vbCrLf
    End If
    
    ' Assinatura
    Set rng = GetAssinaturaRange(doc)
    If Not rng Is Nothing Then
        indice = indice & "6. ASSINATURA" & vbCrLf
        indice = indice & "   Parágrafos: " & rng.Paragraphs.Count & vbCrLf & vbCrLf
    End If
    
    ' Anexo
    Set rng = GetAnexoRange(doc)
    If Not rng Is Nothing Then
        indice = indice & "7. ANEXO" & vbCrLf
        indice = indice & "   Parágrafos: " & rng.Paragraphs.Count & vbCrLf
        indice = indice & "   Palavras: " & rng.Words.Count & vbCrLf & vbCrLf
    End If
    
    indice = indice & String(50, "=")
    
    ' Cria novo documento com o índice
    Set docNovo = Documents.Add
    docNovo.Range.Text = indice
    docNovo.Range.Font.Name = "Courier New"
    docNovo.Range.Font.Size = 10
    
    MsgBox "Índice gerado em novo documento!", vbInformation
End Sub

'================================================================================
' EXEMPLO 10: Navegar entre elementos
'================================================================================
Sub Exemplo10_NavegarProximoElemento()
    Static elementoAtual As Long
    Dim doc As Document
    Dim rng As Range
    
    Set doc = ActiveDocument
    
    elementoAtual = elementoAtual + 1
    If elementoAtual > 7 Then elementoAtual = 1
    
    Select Case elementoAtual
        Case 1
            Set rng = GetTituloRange(doc)
            If Not rng Is Nothing Then
                rng.Select
                MsgBox "TÍTULO", vbInformation, "Navegação"
            End If
        Case 2
            Set rng = GetEmentaRange(doc)
            If Not rng Is Nothing Then
                rng.Select
                MsgBox "EMENTA", vbInformation, "Navegação"
            End If
        Case 3
            Set rng = GetProposicaoRange(doc)
            If Not rng Is Nothing Then
                rng.Select
                MsgBox "PROPOSIÇÃO", vbInformation, "Navegação"
            End If
        Case 4
            Set rng = GetJustificativaRange(doc)
            If Not rng Is Nothing Then
                rng.Select
                MsgBox "JUSTIFICATIVA", vbInformation, "Navegação"
            End If
        Case 5
            Set rng = GetDataRange(doc)
            If Not rng Is Nothing Then
                rng.Select
                MsgBox "DATA (PLENÁRIO)", vbInformation, "Navegação"
            End If
        Case 6
            Set rng = GetAssinaturaRange(doc)
            If Not rng Is Nothing Then
                rng.Select
                MsgBox "ASSINATURA", vbInformation, "Navegação"
            End If
        Case 7
            Set rng = GetAnexoRange(doc)
            If Not rng Is Nothing Then
                rng.Select
                MsgBox "ANEXO", vbInformation, "Navegação"
            Else
                MsgBox "ANEXO não presente", vbInformation, "Navegação"
            End If
    End Select
End Sub

'================================================================================
' NOTAS IMPORTANTES:
'================================================================================
' 1. Execute PadronizarDocumentoMain ANTES de usar esses exemplos
' 2. O cache de parágrafos deve estar construído para as funções funcionarem
' 3. Todos os exemplos são seguros e não modificam o documento original
'    (exceto Exemplo 5 que adiciona marcadores e Exemplo 7 que adiciona destaques)
' 4. Para uso em produção, adicione tratamento de erros adequado
'================================================================================
