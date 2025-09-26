' =============================================================================
' TESTE SIMPLES DO CHAINSAW PROPOSITURAS
' =============================================================================
' Arquivo de teste para a versão simplificada

Sub TesteSimples()
    MsgBox "Chainsaw Proposituras - Teste básico" & vbCrLf & _
           "Data/Hora: " & Now, vbInformation, "Teste OK"
End Sub

Sub TestarPadronizacao()
    ' Verificar se há documento
    If ActiveDocument Is Nothing Then
        MsgBox "Abra um documento para testar a padronização.", vbExclamation, "Teste"
        Exit Sub
    End If
    
    ' Executar padronização
    Call PadronizarDocumento
End Sub

Sub CriarDocumentoTeste()
    ' Criar documento de exemplo para teste
    Dim doc As Document
    Set doc = Documents.Add
    
    ' Adicionar conteúdo de exemplo
    With doc.Range
        .Text = "PROPOSTA DE LEI ORDINÁRIA" & vbCrLf & _
                "    Autor: Deputado Exemplo" & vbCrLf & _
                "    Data: " & Date & vbCrLf & _
                "    Assunto: Exemplo de padronização" & vbCrLf & _
                vbCrLf & _
                "considerando que é necessário testar o sistema;" & vbCrLf & _
                "considerando que   múltiplos   espaços   devem   ser   removidos;" & vbCrLf & _
                vbCrLf & vbCrLf & vbCrLf & _
                "Este é um documento de teste com formatação irregular."
    End With
    
    MsgBox "Documento de teste criado!" & vbCrLf & _
           "Execute 'PadronizarDocumento' para testar a padronização.", vbInformation, "Teste"
End Sub