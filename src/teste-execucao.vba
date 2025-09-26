' =============================================================================
' TESTE DE EXECUÇÃO DO CHAINSAW PROPOSITURAS
' =============================================================================
' Use este código para testar se o sistema está executando corretamente
' Copie e cole no Editor VBA do Word

Sub TesteExecucaoChainsaw()
    ' Teste básico de execução
    MsgBox "Iniciando teste de execução do Chainsaw...", vbInformation, "Teste"
    
    On Error GoTo ErroTeste
    
    ' Verificar se o documento está aberto
    If ActiveDocument Is Nothing Then
        MsgBox "ERRO: Nenhum documento está aberto." & vbCrLf & _
               "Abra um documento antes de executar o teste.", vbExclamation, "Teste - Erro"
        Exit Sub
    End If
    
    ' Verificar se a subrotina principal existe
    MsgBox "Documento ativo encontrado: " & ActiveDocument.Name & vbCrLf & _
           "Tentando executar PadronizarDocumentoMain...", vbInformation, "Teste"
    
    ' Chamar a subrotina principal
    Call PadronizarDocumentoMain
    
    MsgBox "Teste concluído com sucesso!", vbInformation, "Teste"
    Exit Sub
    
ErroTeste:
    MsgBox "ERRO no teste: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Possíveis causas:" & vbCrLf & _
           "1. Subrotina PadronizarDocumentoMain não encontrada" & vbCrLf & _
           "2. Erro de compilação no código" & vbCrLf & _
           "3. Problema com configurações", vbCritical, "Teste - Erro"
End Sub

Sub TesteConfiguracao()
    ' Teste de carregamento de configuração
    MsgBox "Testando carregamento de configuração...", vbInformation, "Teste Config"
    
    On Error GoTo ErroConfig
    
    ' Verificar se a função LoadConfiguration existe
    Dim resultado As Boolean
    resultado = LoadConfiguration()
    
    If resultado Then
        MsgBox "Configuração carregada com sucesso!", vbInformation, "Teste Config"
    Else
        MsgBox "Falha ao carregar configuração, mas sem erro crítico.", vbExclamation, "Teste Config"
    End If
    
    Exit Sub
    
ErroConfig:
    MsgBox "ERRO ao carregar configuração: " & Err.Number & " - " & Err.Description, vbCritical, "Teste Config - Erro"
End Sub

Sub TesteSimples()
    ' Teste mínimo
    MsgBox "Se esta mensagem apareceu, o VBA está funcionando!" & vbCrLf & _
           "Data/Hora: " & Now, vbInformation, "Teste Simples"
End Sub