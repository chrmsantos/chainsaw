# Chainsaw Proposituras - Simplificado ✂️

**Sistema automatizado de padronização de documentos legislativos para Microsoft Word**

## 🎯 Duas Versões Disponíveis

### ⚡ Versão SIMPLES (Recomendada)
- **📁 Arquivos**: 2 arquivos apenas
- **📄 Código**: 200 linhas
- **⏱️ Instalação**: 2 minutos
- **🚀 Performance**: Execução instantânea
- **👤 Ideal para**: Uso pessoal, aprendizado, modificações

### 🔧 Versão COMPLETA (Avançada)
- **📁 Arquivos**: 15+ arquivos
- **📄 Código**: 7.400+ linhas
- **⏱️ Instalação**: 30 minutos
- **🛡️ Recursos**: Backup, logging, configurações avançadas
- **🏢 Ideal para**: Ambiente corporativo, uso crítico

## 🚀 Instalação Rápida (Versão Simples)

### 1. Download dos Arquivos
```
src/chainsaw-simples.bas     # Módulo principal (150 linhas)
src/teste-simples.vba        # Testes (50 linhas)
```

### 2. Importar no Word
1. Abra o Word
2. Pressione `Alt+F11` (Editor VBA)
3. File > Import File > Selecione `chainsaw-simples.bas`
4. (Opcional) Import File > Selecione `teste-simples.vba`

### 3. Testar
```vba
' No Editor VBA, execute:
Call Teste
```

## 📝 Como Usar

### Uso Básico
```vba
' 1. Abra um documento no Word
' 2. No Editor VBA, execute:
Call PadronizarDocumento
```

### Teste com Exemplo
```vba
' 1. Criar documento de teste:
Call CriarDocumentoTeste

' 2. Padronizar:
Call PadronizarDocumento
```

## ⚙️ O que o Sistema Faz

### ✅ Formatação Automática
1. **Primeira Linha**
   - Transforma em CAIXA ALTA
   - Aplica negrito + sublinhado
   - Centraliza o texto

2. **Parágrafos 2-4**
   - Recuo esquerdo de 9cm
   - Remove recuo da primeira linha

3. **Padronização "Considerando"**
   - "considerando" → "CONSIDERANDO"
   - Aplica negrito automaticamente

4. **Limpeza Geral**
   - Remove múltiplos espaços
   - Remove quebras de linha excessivas
   - Padroniza fonte (Times New Roman, 12pt)

### 🎯 Exemplo Prático

**ANTES** (Documento despadronizado):
```
proposta de lei ordinária
    Autor: Deputado João
    Data: 01/01/2025
    Assunto: regulamentação

considerando que   há    múltiplos    espaços;


considerando que há quebras excessivas;
```

**DEPOIS** (Automaticamente padronizado):
```
PROPOSTA DE LEI ORDINÁRIA
                 Autor: Deputado João
                 Data: 01/01/2025
                 Assunto: regulamentação

CONSIDERANDO que há múltiplos espaços;

CONSIDERANDO que há quebras excessivas;
```

## 📊 Comparação das Versões

| Característica | Versão Simples | Versão Completa |
|----------------|----------------|-----------------|
| **Linhas de Código** | 200 | 7.400+ |
| **Arquivos** | 2 | 15+ |
| **Instalação** | 2 min | 30 min |
| **Execução** | Instantânea | 5-30 seg |
| **Configuração** | Código direto | Arquivo INI |
| **Backup** | Manual | Automático |
| **Logs** | Básico | Detalhado |
| **Validações** | Essenciais | Avançadas |
| **Personalização** | Código | Interface |
| **Manutenção** | Trivial | Complexa |

## 🔧 Personalização (Versão Simples)

### Alterar Recuo dos Parágrafos
```vba
' Em FormatarParagrafos2a4(), linha:
.LeftIndent = CentimetersToPoints(9)  ' Mude para seu valor
```

### Alterar Fonte Padrão  
```vba
' Em LimparFormatacao():
.Size = 12                    # Tamanho da fonte
.Name = "Times New Roman"     # Nome da fonte
```

### Adicionar Mais Parágrafos com Recuo
```vba
' Em FormatarParagrafos2a4():
For i = 2 To 4    # Mude para: For i = 2 To 6
```

## 🧪 Testes Disponíveis

```vba
Call TesteSimples          # Teste básico de funcionamento
Call CriarDocumentoTeste   # Criar documento de exemplo  
Call TestarPadronizacao    # Testar com documento ativo
Call PadronizarDocumento   # Executar padronização
```

## 🆚 Qual Versão Escolher?

### 🎯 Escolha a Versão SIMPLES se:
- ✅ Quer algo que funcione imediatamente
- ✅ Não precisa de configurações complexas
- ✅ Prefere código fácil de entender/modificar
- ✅ Uso pessoal ou em pequena escala
- ✅ Performance é importante

### 🔧 Escolha a Versão COMPLETA se:
- ⚙️ Precisa de configurações detalhadas
- 🛡️ Precisa de backups automáticos  
- 📊 Precisa de logs detalhados para auditoria
- 🏢 Uso em ambiente corporativo
- 🔍 Precisa de validações avançadas

## 📁 Estrutura do Projeto

```
chainsaw/
├── src/
│   ├── chainsaw-simples.bas    # ⚡ Versão Simples (150 linhas)
│   ├── teste-simples.vba       # ⚡ Testes Simples (50 linhas)  
│   ├── chainsaw0.bas           # 🔧 Versão Completa (7.400 linhas)
│   └── teste-execucao.vba      # 🔧 Testes Completos (233 linhas)
├── docs/
│   ├── README-VERSAO-SIMPLES.md      # Documentação da versão simples
│   ├── GUIA-SIMPLIFICACAO.md         # Como foi simplificado
│   └── [outros documentos]           # Docs da versão completa
└── config/
    └── chainsaw-config.ini           # Apenas para versão completa
```

## 🎯 Recomendação

**Para 90% dos usuários**: Use a **Versão Simples**
- Mais rápida, confiável e fácil
- 80% da funcionalidade com 3% da complexidade  
- Instalação em 2 minutos
- Modificação trivial

**Para uso corporativo crítico**: Use a **Versão Completa**
- Recursos avançados de auditoria e backup
- Configurações detalhadas
- Validações completas

## 📞 Suporte

### Versão Simples
- **Código**: Apenas 200 linhas - fácil de debugar
- **Modificação**: Editar diretamente o código
- **Problemas**: Geralmente VBA básico

### Versão Completa  
- **Documentação**: Extensa documentação disponível
- **Configuração**: Sistema INI configurável
- **Logs**: Sistema de auditoria detalhado

---

**🚀 Comece agora**: Baixe `chainsaw-simples.bas` e execute `Call Teste`!

**📌 Filosofia**: "Simplicidade é a sofisticação suprema" - Leonardo da Vinci