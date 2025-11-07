# GUIA RÁPIDO - Sistema de Identificação de Elementos v1.1

## 🚀 Começando em 5 Minutos

### 1. O Que é Isso?

O CHAINSAW v1.1 agora identifica automaticamente todos os elementos da sua propositura:
- Título, Ementa, Proposição, Justificativa, Data, Assinatura, Anexo

### 2. Como Funciona?

**Automático!** Quando você executa `PadronizarDocumentoMain`, o sistema:
1. Analisa o documento
2. Identifica todos os elementos
3. Registra no log
4. Disponibiliza funções de acesso

### 3. Primeiro Uso

#### Passo 1: Padronize o Documento
```vba
' Abra o documento
' Pressione Alt + F8
' Execute: PadronizarDocumentoMain
```

#### Passo 2: Veja os Elementos Identificados
```vba
Sub VerElementos()
    MsgBox GetElementInfo(ActiveDocument)
End Sub
```

### 4. Exemplos Práticos

#### Exemplo A: Selecionar a Proposição
```vba
Sub SelecionarProposicao()
    Dim rng As Range
    Set rng = GetProposicaoRange(ActiveDocument)
    If Not rng Is Nothing Then
        rng.Select
        MsgBox "Proposição selecionada!"
    End If
End Sub
```

#### Exemplo B: Contar Palavras da Justificativa
```vba
Sub ContarPalavrasJustificativa()
    Dim rng As Range
    Set rng = GetJustificativaRange(ActiveDocument)
    If Not rng Is Nothing Then
        MsgBox "Justificativa: " & rng.Words.Count & " palavras"
    End If
End Sub
```

#### Exemplo C: Validar Estrutura
```vba
Sub ValidarDocumento()
    Dim erros As Long
    erros = 0
    
    If GetTituloRange(ActiveDocument) Is Nothing Then erros = erros + 1
    If GetEmentaRange(ActiveDocument) Is Nothing Then erros = erros + 1
    If GetProposicaoRange(ActiveDocument) Is Nothing Then erros = erros + 1
    
    If erros = 0 Then
        MsgBox "Estrutura válida! ✓", vbInformation
    Else
        MsgBox erros & " erro(s) encontrado(s)!", vbExclamation
    End If
End Sub
```

### 5. Todas as Funções Disponíveis

| Função | Retorna | Descrição |
|--------|---------|-----------|
| `GetTituloRange(doc)` | Range | Título da propositura |
| `GetEmentaRange(doc)` | Range | Ementa |
| `GetProposicaoRange(doc)` | Range | Proposição completa |
| `GetTituloJustificativaRange(doc)` | Range | Título "Justificativa" |
| `GetJustificativaRange(doc)` | Range | Justificativa completa |
| `GetDataRange(doc)` | Range | Data do plenário |
| `GetAssinaturaRange(doc)` | Range | Assinatura + imagens |
| `GetTituloAnexoRange(doc)` | Range | Título do anexo |
| `GetAnexoRange(doc)` | Range | Conteúdo do anexo |
| `GetProposituraRange(doc)` | Range | Documento completo |
| `GetElementInfo(doc)` | String | Relatório completo |

### 6. Dica: Copiar Exemplos Prontos

Abra o arquivo `src/Exemplos_Identificacao.bas` no editor VBA e você encontrará:
- ✅ 10 exemplos completos e funcionais
- ✅ Código pronto para copiar e usar
- ✅ Comentários explicativos

### 7. O Que Fazer Se...

#### ...não encontrar um elemento?
A função retorna `Nothing`. Sempre verifique:
```vba
Dim rng As Range
Set rng = GetProposicaoRange(ActiveDocument)
If rng Is Nothing Then
    MsgBox "Proposição não encontrada!"
Else
    ' Use o rng aqui
End If
```

#### ...o documento não estiver padronizado?
Execute `PadronizarDocumentoMain` primeiro. A identificação só funciona após a padronização.

#### ...quiser ver o log?
Abra o arquivo de log na mesma pasta do documento:
`CHAINSAW_AAAAMMDD_HHMMSS_nomedocumento.log`

### 8. Recursos Adicionais

📖 **Documentação Completa:**
- `docs/IDENTIFICACAO_ELEMENTOS.md` - Guia técnico detalhado (200+ linhas)
- `docs/NOVIDADES_v1.1.md` - Resumo executivo
- `docs/RESUMO_IMPLEMENTACAO.md` - Relatório de implementação

💡 **Exemplos Práticos:**
- `src/Exemplos_Identificacao.bas` - 10 exemplos prontos (500+ linhas)

📝 **Histórico:**
- `CHANGELOG.md` - Todas as mudanças da v1.1

### 9. Casos de Uso Comuns

#### Navegação Rápida
```vba
' Cole os exemplos do arquivo Exemplos_Identificacao.bas
' Execute: Exemplo10_NavegarProximoElemento
' Pressione F5 repetidamente para navegar
```

#### Análise Estatística
```vba
' Execute: Exemplo3_ContarPalavrasPorElemento
' Veja quantas palavras tem cada seção
```

#### Debug Visual
```vba
' Execute: Exemplo7_DestacaElementosVisualmente
' Veja cada seção destacada com cor diferente
' Execute: Exemplo8_RemoverDestaques para limpar
```

#### Exportação
```vba
' Execute: Exemplo4_ExportarProposicao
' Cria novo documento só com a proposição
```

### 10. Precisa de Ajuda?

**Problema:** Função não encontra elemento  
**Solução:** Verifique se o documento segue o formato padrão

**Problema:** Erro ao executar  
**Solução:** Execute `PadronizarDocumentoMain` primeiro

**Problema:** Elemento identificado errado  
**Solução:** Verifique os critérios em `docs/IDENTIFICACAO_ELEMENTOS.md`

**Contato:** chrmsantos@protonmail.com

---

## ⚡ Início Ultra-Rápido (30 segundos)

1. Abra seu documento
2. `Alt + F8` → `PadronizarDocumentoMain` → `Executar`
3. `Alt + F8` → Cole e execute:

```vba
Sub Teste()
    MsgBox GetElementInfo(ActiveDocument)
End Sub
```

4. Veja a mágica acontecer! ✨

---

**Versão:** 1.1-RC1-202511071045  
**Última atualização:** 07/11/2024  
**Licença:** GNU GPLv3
