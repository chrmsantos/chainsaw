# Chainsaw Proposituras

Sistema de padronização automática de documentos legislativos para Microsoft Word.

## Instalação

1. Baixe o arquivo `src/chainsaw-simples.bas`
2. Abra o Word e pressione Alt+F11
3. File > Import File > chainsaw-simples.bas
4. Execute: `Call Teste`

## Como Usar

```vba
' Abra um documento no Word, depois execute:
Call PadronizarDocumento
```

## O que faz

- **Primeira linha**: CAIXA ALTA, negrito, sublinhado, centralizada
- **Parágrafos 2-4**: Recuo de 9cm 
- **"Considerando"**: Transforma em CONSIDERANDO (negrito)
- **Limpeza**: Remove espaços múltiplos e quebras excessivas

## Exemplo

**Antes:**
```
proposta de lei ordinária
    Autor: João Silva
considerando que há necessidade...
```

**Depois:**
```
PROPOSTA DE LEI ORDINÁRIA
                 Autor: João Silva
CONSIDERANDO que há necessidade...
```

## Teste Completo

```vba
Call CriarDocumentoTeste  ' Cria exemplo
Call PadronizarDocumento  ' Aplica formatação
```

## Personalização

Para alterar o recuo dos parágrafos, edite a linha:

```vba
.LeftIndent = CentimetersToPoints(9)  ' Mude para seu valor
```

## Características

- **📄 Código**: 150 linhas apenas
- **⚡ Performance**: Execução instantânea  
- **🔧 Manutenção**: Código simples e claro
- **📦 Instalação**: 2 minutos

---

**Versão**: 2.0-Simple | **Licença**: Apache 2.0