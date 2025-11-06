# Substituições Condicionais de Texto

## Resumo

Implementação de lógica condicional para substituições de texto no segundo parágrafo, baseada no **tipo de documento** (determinado pela primeira palavra).

---

## Modificações Realizadas

### 1. Nova Função: `GetFirstWordOfDocument`

**Localização**: `src/chainsaw.bas` (linhas 2147-2202)

**Propósito**: Identificar o tipo de documento através da primeira palavra.

**Funcionalidades**:
- Busca o primeiro parágrafo com conteúdo (ignora parágrafos vazios)
- Extrai a primeira palavra (tudo antes do primeiro espaço)
- Remove pontuação comum (`:`, `,`, `.`, `;`)
- Retorna em **MAIÚSCULAS** para comparação case-insensitive
- Proteção: analisa apenas os primeiros 10 parágrafos
- Tratamento de erros com log de warning

**Exemplo de Uso**:
```vba
Dim docType As String
docType = GetFirstWordOfDocument(doc)
' Retorna: "INDICAÇÃO", "REQUERIMENTO", "PROJETO", etc.
```

---

### 2. Substituições Condicionais no `FormatSecondParagraph`

**Localização**: `src/chainsaw.bas` (linhas 2255-2313)

#### Regra 1: "Sugere" → "Indica"
- **Condição**: Só ocorre se a 1ª palavra do documento for **"INDICAÇÃO"**
- **Comportamento anterior**: Sempre substituía
- **Comportamento atual**: Condicional com validação de tipo

```vba
' ANTES:
If lowerStart = "sugere" Then
    para.Range.text = "Indica" & Mid(paraFullText, 7) & vbCr
    LogMessage "Palavra inicial 'Sugere' substituída por 'Indica'..."
End If

' DEPOIS:
If lowerStart = "sugere" Then
    If docFirstWord = "INDICAÇÃO" Then
        para.Range.text = "Indica" & Mid(paraFullText, 7) & vbCr
        LogMessage "...(documento tipo INDICAÇÃO)", LOG_LEVEL_INFO
    Else
        LogMessage "..não substituída (documento não é INDICAÇÃO, é: " & docFirstWord & ")"
    End If
End If
```

#### Regra 2: "Pede" → "Requer"
- **Condição**: Só ocorre se a 1ª palavra do documento for **"REQUERIMENTO"**
- **Comportamento anterior**: Sempre substituía
- **Comportamento atual**: Condicional com validação de tipo

#### Regra 3: "Solicita" → "Requer"
- **Condição**: Só ocorre se a 1ª palavra do documento for **"REQUERIMENTO"**
- **Comportamento anterior**: Sempre substituía
- **Comportamento atual**: Condicional com validação de tipo

---

## Exemplos Práticos

### Cenário 1: Documento de Indicação
```
Primeira linha: "INDICAÇÃO N.º 123/2024"
Segunda linha: "Sugere a construção..."

RESULTADO:
[OK] "Sugere" → "Indica" (substituição ocorre)
[X] "Pede" → não substituído (documento não é REQUERIMENTO)
[X] "Solicita" → não substituído (documento não é REQUERIMENTO)
```

### Cenário 2: Documento de Requerimento
```
Primeira linha: "REQUERIMENTO"
Segunda linha: "Pede providências..."

RESULTADO:
[OK] "Pede" → "Requer" (substituição ocorre)
[OK] "Solicita" → "Requer" (substituição ocorre)
[X] "Sugere" → não substituído (documento não é INDICAÇÃO)
```

### Cenário 3: Outro Tipo de Documento
```
Primeira linha: "PROJETO DE LEI N.º 456/2024"
Segunda linha: "Sugere alteração..."

RESULTADO:
[X] "Sugere" → não substituído (documento não é INDICAÇÃO)
[X] "Pede" → não substituído (documento não é REQUERIMENTO)
[X] "Solicita" → não substituído (documento não é REQUERIMENTO)

LOG: "Palavra inicial 'Sugere' não substituída (documento não é INDICAÇÃO, é: PROJETO)"
```

---

## Características Técnicas

### Segurança
- [OK] Comparação **case-insensitive** (INDICAÇÃO = indicação = Indicação)
- [OK] Proteção contra documentos vazios
- [OK] Limite de 10 parágrafos para busca da primeira palavra
- [OK] Tratamento de erros com logging adequado
- [OK] Validação de comprimento antes de substring

### Performance
- [OK] Busca otimizada (para no primeiro parágrafo com conteúdo)
- [OK] Cache da primeira palavra (calcula uma vez por execução)
- [OK] Não afeta parágrafos além do 2º

### Logging
- [OK] Log informativo quando **substituição ocorre** (nível INFO)
- [OK] Log informativo quando **não ocorre** devido ao tipo de documento (nível INFO)
- [OK] Log de warning em caso de **erro** na função auxiliar

---

## Compatibilidade

- [OK] Compatível com Word 2010+
- [OK] Não quebra funcionalidade existente
- [OK] Retrocompatível (documentos sem tipo específico não causam erro)
- [OK] Funciona com acentuação (INDICAÇÃO com Ç)

---

## Manutenção

### Para Adicionar Novos Tipos de Documento
1. Editar `GetFirstWordOfDocument` se necessário normalização adicional
2. Adicionar nova condição em `FormatSecondParagraph`:
```vba
' Exemplo: "Propõe" → "Sugere" apenas em PROJETO DE LEI
If lowerStart = "propõe" Then
    If docFirstWord = "PROJETO" Then
        para.Range.text = "Sugere" & Mid(paraFullText, 7) & vbCr
        LogMessage "...(documento tipo PROJETO)", LOG_LEVEL_INFO
    Else
        LogMessage "..não substituída (não é PROJETO...)", LOG_LEVEL_INFO
    End If
End If
```

### Para Debug
- Verificar logs: busque por "documento tipo" ou "não substituída"
- Testar função isoladamente:
```vba
Debug.Print GetFirstWordOfDocument(ActiveDocument)
```

---

## Histórico de Mudanças

| Data | Versão | Descrição |
|------|--------|-----------|
| 2024 | 1.0 | Implementação inicial das substituições condicionais |

---

## Referências

- Função principal: `FormatSecondParagraph` (linha 2204)
- Função auxiliar: `GetFirstWordOfDocument` (linha 2147)
- Sistema de logging: `LogMessage` (definido anteriormente no código)

---

**Status**: [OK] Implementado e testado
**Próximas melhorias**: Considerar expansão para outros tipos de documentos conforme necessidade
