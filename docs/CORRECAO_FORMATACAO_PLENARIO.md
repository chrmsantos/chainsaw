# Correção: Formatação dos Parágrafos Após "Plenário Dr. Tancredo Neves"

## 🐛 Problema Identificado

Os parágrafos (linhas em branco) inseridos após a linha contendo "Plenário "Dr. Tancredo Neves", $DATAATUALEXTENSO$." não estavam sendo formatados corretamente:
- ❌ Não estavam centralizados
- ❌ Tinham recuos diferentes de zero

## ✅ Correção Implementada

### Arquivos Modificados

- **`src/chainsaw.bas`**

### Funções Alteradas

#### 1. `EnsurePlenarioBlankLines` (linhas ~2725-2760)

**Antes:**
```vba
' Insere EXATAMENTE 2 linhas em branco ANTES
Set para = doc.Paragraphs(plenarioIndex)
para.Range.InsertParagraphBefore
para.Range.InsertParagraphBefore

' Insere EXATAMENTE 2 linhas em branco DEPOIS
Set para = doc.Paragraphs(plenarioIndex + 2)
para.Range.InsertParagraphAfter
para.Range.InsertParagraphAfter
```

**Depois:**
```vba
' Insere EXATAMENTE 2 linhas em branco ANTES
Set para = doc.Paragraphs(plenarioIndex)
para.Range.InsertParagraphBefore
para.Range.InsertParagraphBefore

' Formata as linhas em branco inseridas ANTES: centralizado e recuos 0
Dim j As Long
For j = plenarioIndex To plenarioIndex + 1
    If j <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(j)
        With para.Format
            .leftIndent = 0
            .firstLineIndent = 0
            .RightIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
            .alignment = wdAlignParagraphCenter
        End With
    End If
Next j

' Insere EXATAMENTE 2 linhas em branco DEPOIS
Set para = doc.Paragraphs(plenarioIndex + 2)
para.Range.InsertParagraphAfter
para.Range.InsertParagraphAfter

' Formata as linhas em branco inseridas DEPOIS: centralizado e recuos 0
For j = plenarioIndex + 3 To plenarioIndex + 4
    If j <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(j)
        With para.Format
            .leftIndent = 0
            .firstLineIndent = 0
            .RightIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
            .alignment = wdAlignParagraphCenter
        End With
    End If
Next j
```

#### 2. `InsertBlankLines` (linhas ~4860-4920)

**Antes:**
```vba
' Insere EXATAMENTE 2 linhas em branco ANTES
Set para = doc.Paragraphs(plenarioIndex)
para.Range.InsertParagraphBefore
para.Range.InsertParagraphBefore

' Insere EXATAMENTE 2 linhas em branco DEPOIS
Set para = doc.Paragraphs(plenarioIndex + 2)
para.Range.InsertParagraphAfter
para.Range.InsertParagraphAfter
```

**Depois:**
```vba
' Insere EXATAMENTE 2 linhas em branco ANTES
Set para = doc.Paragraphs(plenarioIndex)
para.Range.InsertParagraphBefore
para.Range.InsertParagraphBefore

' Formata as linhas em branco inseridas ANTES: centralizado e recuos 0
For i = plenarioIndex To plenarioIndex + 1
    If i <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(i)
        With para.Format
            .leftIndent = 0
            .firstLineIndent = 0
            .RightIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
            .alignment = wdAlignParagraphCenter
        End With
    End If
Next i

' Insere EXATAMENTE 2 linhas em branco DEPOIS
Set para = doc.Paragraphs(plenarioIndex + 2)
para.Range.InsertParagraphAfter
para.Range.InsertParagraphAfter

' Formata as linhas em branco inseridas DEPOIS: centralizado e recuos 0
For i = plenarioIndex + 3 To plenarioIndex + 4
    If i <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(i)
        With para.Format
            .leftIndent = 0
            .firstLineIndent = 0
            .RightIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
            .alignment = wdAlignParagraphCenter
        End With
    End If
Next i
```

## 📋 O Que Foi Corrigido

### Formatação Aplicada aos Parágrafos

Para cada linha em branco inserida antes e depois do parágrafo "Plenário Dr. Tancredo Neves":

1. **Alinhamento**: `wdAlignParagraphCenter` (centralizado)
2. **Recuo à esquerda**: `0`
3. **Recuo de primeira linha**: `0`
4. **Recuo à direita**: `0`
5. **Espaçamento antes**: `0`
6. **Espaçamento depois**: `0`

### Parágrafos Afetados

A correção se aplica a **4 parágrafos** no total:

- **2 linhas em branco ANTES** do parágrafo do Plenário (índices: `plenarioIndex` e `plenarioIndex + 1`)
- **2 linhas em branco DEPOIS** do parágrafo do Plenário (índices: `plenarioIndex + 3` e `plenarioIndex + 4`)

### Estrutura Visual

```
[linha em branco] ← Formatada: centro, recuos 0
[linha em branco] ← Formatada: centro, recuos 0
Plenário "Dr. Tancredo Neves", $DATAATUALEXTENSO$. ← Já era formatado
[linha em branco] ← Formatada: centro, recuos 0
[linha em branco] ← Formatada: centro, recuos 0
```

## 🧪 Validação

### Verificação de Sintaxe

```
✓ Functions: 79 | End Function: 79
✓ Subs: 33 | End Sub: 33
```

Todas as funções e sub-rotinas estão corretamente fechadas.

## 🔍 Funções Relacionadas

As seguintes funções trabalham juntas para garantir a formatação correta:

1. **`ReplacePlenarioDateParagraph`** (linha ~6430)
   - Substitui o texto do parágrafo do Plenário
   - Aplica formatação ao próprio parágrafo (centralizado, recuos 0)

2. **`EnsurePlenarioBlankLines`** (linha ~2640)
   - Garante exatamente 2 linhas em branco antes e depois
   - **AGORA**: Formata essas linhas (centralizado, recuos 0)

3. **`InsertBlankLines`** (linha ~4700)
   - Insere linhas em branco estruturais no documento
   - **AGORA**: Formata as linhas inseridas ao redor do Plenário

4. **`CenterImageAfterPlenario`** (linha ~5850)
   - Centraliza imagens entre linhas 5-7 após o Plenário
   - Não modificada (já funcionava corretamente)

## 📊 Impacto

### Antes da Correção

```
Justificativa
[linha em branco - sem formatação específica]
[linha em branco - sem formatação específica]
Plenário "Dr. Tancredo Neves", $DATAATUALEXTENSO$. [centralizado]
[linha em branco - SEM FORMATAÇÃO] ← PROBLEMA
[linha em branco - SEM FORMATAÇÃO] ← PROBLEMA
[linha em branco - SEM FORMATAÇÃO] ← PROBLEMA (se houver 3ª)
[possível imagem centralizada]
Excelentíssimo Senhor Prefeito Municipal,
```

### Depois da Correção

```
Justificativa
[linha em branco - centralizada, recuos 0] ✓
[linha em branco - centralizada, recuos 0] ✓
Plenário "Dr. Tancredo Neves", $DATAATUALEXTENSO$. [centralizado] ✓
[linha em branco - centralizada, recuos 0] ✓
[linha em branco - centralizada, recuos 0] ✓
[possível imagem centralizada]
Excelentíssimo Senhor Prefeito Municipal,
```

## 🎯 Resultado Esperado

Ao executar a macro Chainsaw em um documento:

1. ✅ O parágrafo "Plenário Dr. Tancredo Neves" estará centralizado com recuos zero
2. ✅ As 2 linhas em branco ANTES estarão centralizadas com recuos zero
3. ✅ As 2 linhas em branco DEPOIS estarão centralizadas com recuos zero
4. ✅ Toda a seção terá formatação consistente e profissional

## 🚀 Próximos Passos

1. **Importar o módulo atualizado** no Word:
   - Abra o VBA Editor (Alt + F11)
   - Remova o módulo `chainsaw` antigo
   - Importe o arquivo `chainsaw.bas` atualizado

2. **Testar em documento de exemplo**:
   - Abra um documento de propositura
   - Execute a macro Chainsaw
   - Verifique a formatação do parágrafo do Plenário e linhas adjacentes

3. **Validar visualmente**:
   - Use a régua do Word para verificar recuos
   - Verifique o alinhamento (deve estar centralizado)
   - Confirme que não há espaçamentos extras

## 📝 Notas Técnicas

### Por que 4 linhas ao invés de 3?

O código insere **2 linhas antes** e **2 linhas depois** do parágrafo do Plenário, totalizando **4 linhas em branco** + 1 linha com texto = **5 linhas** na seção do Plenário.

Se você estava vendo 3 linhas com problemas, provavelmente eram:
- 1 linha antes (a 2ª linha antes do Plenário)
- 2 linhas depois (as 2 linhas logo após o Plenário)

Agora **TODAS as 4 linhas em branco** estão corretamente formatadas.

### Índices dos Parágrafos

```
plenarioIndex     → Linha em branco ANTES (1ª)
plenarioIndex + 1 → Linha em branco ANTES (2ª)
plenarioIndex + 2 → Parágrafo "Plenário Dr. Tancredo Neves"
plenarioIndex + 3 → Linha em branco DEPOIS (1ª) ← FORMATADA
plenarioIndex + 4 → Linha em branco DEPOIS (2ª) ← FORMATADA
```

## ✅ Status

- [x] Problema identificado
- [x] Código corrigido
- [x] Sintaxe validada
- [x] Documentação criada
- [ ] Teste em ambiente real (próximo passo)

---

**Correção aplicada em:** 05/11/2025  
**Arquivo modificado:** `src/chainsaw.bas`  
**Linhas alteradas:** ~2725-2760, ~4860-4920  
**Funções modificadas:** `EnsurePlenarioBlankLines`, `InsertBlankLines`  
**Status:** ✅ Pronto para teste
